import win32com.client
import csv
import math
import random
import time
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

MODEL_PATH = r"CHANGE_THIS_TO_THE_LOCATION_OF_YOUR_MODEL"

SEQUENCE_LENGTH = 64
EXPECTED_EXITED_PARTS = 64

# Simulated Annealing settings
SA_MAX_EVALUATIONS = 3000
SA_INITIAL_TEMPERATURE = 100.0
SA_COOLING_RATE = 0.999
SA_MIN_TEMPERATURE = 0.01

# Use an integer for an exactly reproducible run.
# Use None to generate a new seed automatically; the actual seed is logged.
SA_RANDOM_SEED = None

# Relative probabilities of selecting each neighbourhood operator.
# These are normalized automatically and can be adjusted during preliminary calibration.
SA_OPERATOR_WEIGHTS = {
    "swap": 1.0,
    "move": 1.0,
    "reverse": 1.0,
    "reinsert_block": 1.0,
    "swap_blocks": 0.5,
}

# Logging / CSV compatibility for Windows / Excel
CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"
WRITE_LOGS = True

# Polling settings for simulation completion detection
POLL_INTERVAL_SEC = 0.2
STABLE_POLLS_REQUIRED = 3
SIM_TIMEOUT_SEC = 300

# Macro-move settings - kept identical to the LLM-guided optimiser
MIN_REINSERT_BLOCK_LEN = 2
MAX_REINSERT_BLOCK_LEN = 12
MIN_SWAP_BLOCK_LEN = 2
MAX_SWAP_BLOCK_LEN = 8

# Plant Simulation paths
EXITED_PARTS_PATH = ".Models.Model.ExitedParts"

# ============================================================
# PATHS
# ============================================================

def get_base_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


BASE_DIR = get_base_dir()
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = LOG_DIR / f"sa_optimization_log_{RUN_TIMESTAMP}.csv"
GLOBAL_BEST_SEQUENCE_FILE = LOG_DIR / f"sa_global_best_sequence_{RUN_TIMESTAMP}.txt"
RUN_SUMMARY_FILE = LOG_DIR / f"sa_run_summary_{RUN_TIMESTAMP}.txt"

# ============================================================
# PLANT SIMULATION COMMUNICATION
# ============================================================

def set_sequence(ps, sequence: List[str]) -> None:
    """Writes a complete production sequence to Plant Simulation."""
    for i, part_name in enumerate(sequence, start=1):
        ps.ExecuteSimTalk(
            f'.Models.Model.SourceSequence[1, {i}] := .UserObjects.{part_name}'
        )
    for i, part_name in enumerate(sequence, start=1):
        ps.ExecuteSimTalk(
            f'.Models.Model.SourceSequence[3, {i}] := "{part_name}"'
        )


def get_sequence(ps, length: int = SEQUENCE_LENGTH) -> List[str]:
    """Reads the production sequence from Plant Simulation."""
    sequence = []
    for i in range(1, length + 1):
        value = ps.GetValue(f".Models.Model.SourceSequence[1, {i}]")
        sequence.append(str(value).split(".")[-1])
    return sequence


def read_exited_parts(ps) -> Optional[int]:
    """Reads the integer ExitedParts variable from Plant Simulation."""
    try:
        value = ps.GetValue(EXITED_PARTS_PATH)
        return int(float(value))
    except Exception:
        return None


def is_completion_valid(exited_parts: Optional[int]) -> bool:
    return exited_parts == EXPECTED_EXITED_PARTS


def run_simulation_until_finished(
    ps,
    poll_interval_sec: float = POLL_INTERVAL_SEC,
    stable_polls_required: int = STABLE_POLLS_REQUIRED,
    timeout_sec: float = SIM_TIMEOUT_SEC
) -> float:
    """
    Resets and starts Plant Simulation, then waits until SimTime is stable.
    Returns the final simulated time (makespan) in seconds.
    """
    ps.ExecuteSimTalk(".Models.Model.EventController.reset")
    ps.ExecuteSimTalk(".Models.Model.EventController.start")

    start_wall = time.perf_counter()
    last_sim_time = None
    stable_count = 0

    while True:
        if time.perf_counter() - start_wall > timeout_sec:
            raise TimeoutError(
                f"Simulation did not stabilize within {timeout_sec} seconds."
            )

        sim_time = float(
            ps.GetValue(".Models.Model.EventController.SimTime")
        )

        if last_sim_time is not None and abs(sim_time - last_sim_time) < 1e-9:
            stable_count += 1
        else:
            stable_count = 0

        if stable_count >= stable_polls_required:
            return sim_time

        last_sim_time = sim_time
        time.sleep(poll_interval_sec)

# ============================================================
# NEIGHBOURHOOD MOVES
# ============================================================

def format_move(move: Optional[Dict[str, Any]]) -> str:
    if not move:
        return "None"

    action = move.get("action")

    if action == "swap":
        return f'swap({move.get("i")}, {move.get("j")})'
    if action == "move":
        return f'move({move.get("from")} -> {move.get("to")})'
    if action == "reverse":
        return f'reverse({move.get("start")} .. {move.get("end")})'
    if action == "reinsert_block":
        return (
            f'reinsert_block({move.get("start")} .. '
            f'{move.get("end")} -> {move.get("to")})'
        )
    if action == "swap_blocks":
        return (
            f'swap_blocks({move.get("start1")} .. {move.get("end1")}, '
            f'{move.get("start2")} .. {move.get("end2")})'
        )

    return str(move)


def validate_move(
    move: Optional[Dict[str, Any]],
    n: int = SEQUENCE_LENGTH
) -> Tuple[bool, str]:
    """Validation rules kept consistent with the LLM-guided optimiser."""
    if move is None:
        return False, "Move is None"

    action = move.get("action")
    if action not in {
        "swap", "move", "reverse", "reinsert_block", "swap_blocks"
    }:
        return False, f"Unknown action: {action}"

    try:
        if action == "swap":
            i = int(move["i"])
            j = int(move["j"])
            if not (0 <= i < n and 0 <= j < n):
                return False, "swap indices out of range"
            if i == j:
                return False, "swap indices identical"
            return True, "ok"

        if action == "move":
            src = int(move["from"])
            dst = int(move["to"])
            if not (0 <= src < n and 0 <= dst < n):
                return False, "move indices out of range"
            if src == dst:
                return False, "move source and target identical"
            return True, "ok"

        if action == "reverse":
            start = int(move["start"])
            end = int(move["end"])
            if not (0 <= start < n and 0 <= end < n):
                return False, "reverse indices out of range"
            if start >= end:
                return False, "reverse start must be < end"
            return True, "ok"

        if action == "reinsert_block":
            start = int(move["start"])
            end = int(move["end"])
            to = int(move["to"])

            if not (0 <= start < n and 0 <= end < n and 0 <= to < n):
                return False, "reinsert_block indices out of range"
            if start >= end:
                return False, "reinsert_block start must be < end"

            block_len = end - start + 1
            if not (
                MIN_REINSERT_BLOCK_LEN
                <= block_len
                <= MAX_REINSERT_BLOCK_LEN
            ):
                return False, "reinsert_block block length out of allowed range"

            if start <= to <= end:
                return False, "reinsert_block target cannot be inside the block"

            return True, "ok"

        if action == "swap_blocks":
            start1 = int(move["start1"])
            end1 = int(move["end1"])
            start2 = int(move["start2"])
            end2 = int(move["end2"])

            if not all(
                0 <= x < n for x in [start1, end1, start2, end2]
            ):
                return False, "swap_blocks indices out of range"

            if start1 >= end1 or start2 >= end2:
                return False, "swap_blocks each block must have start < end"

            len1 = end1 - start1 + 1
            len2 = end2 - start2 + 1

            if len1 != len2:
                return False, "swap_blocks block lengths must match"

            if not (
                MIN_SWAP_BLOCK_LEN <= len1 <= MAX_SWAP_BLOCK_LEN
            ):
                return False, "swap_blocks block length out of allowed range"

            if not (end1 < start2 or end2 < start1):
                return False, "swap_blocks blocks must be disjoint"

            return True, "ok"

    except KeyError as e:
        return False, f"Missing key: {e}"
    except Exception as e:
        return False, f"Validation error: {e}"

    return False, "Unhandled validation case"


def apply_move(
    sequence: List[str],
    move: Dict[str, Any]
) -> List[str]:
    """Applies one neighbourhood move to a copy of the sequence."""
    seq = sequence.copy()
    action = move["action"]

    if action == "swap":
        i, j = int(move["i"]), int(move["j"])
        seq[i], seq[j] = seq[j], seq[i]
        return seq

    if action == "move":
        src, dst = int(move["from"]), int(move["to"])
        part = seq.pop(src)
        seq.insert(dst, part)
        return seq

    if action == "reverse":
        start, end = int(move["start"]), int(move["end"])
        seq[start:end + 1] = reversed(seq[start:end + 1])
        return seq

    if action == "reinsert_block":
        start = int(move["start"])
        end = int(move["end"])
        to = int(move["to"])

        block = seq[start:end + 1]
        remainder = seq[:start] + seq[end + 1:]

        if to < start:
            insert_pos = to
        else:
            insert_pos = to - len(block) + 1

        insert_pos = max(0, min(insert_pos, len(remainder)))
        return remainder[:insert_pos] + block + remainder[insert_pos:]

    if action == "swap_blocks":
        start1 = int(move["start1"])
        end1 = int(move["end1"])
        start2 = int(move["start2"])
        end2 = int(move["end2"])

        if start2 < start1:
            start1, end1, start2, end2 = (
                start2, end2, start1, end1
            )

        block1 = seq[start1:end1 + 1]
        block2 = seq[start2:end2 + 1]

        new_seq = seq.copy()
        new_seq[start1:end1 + 1] = block2
        new_seq[start2:end2 + 1] = block1
        return new_seq

    return seq


def choose_operator(rng: random.Random) -> str:
    actions = list(SA_OPERATOR_WEIGHTS.keys())
    weights = [SA_OPERATOR_WEIGHTS[a] for a in actions]

    if not actions or any(w < 0 for w in weights) or sum(weights) <= 0:
        raise ValueError(
            "SA_OPERATOR_WEIGHTS must contain non-negative weights "
            "with a positive total."
        )

    return rng.choices(actions, weights=weights, k=1)[0]


def generate_random_move(
    rng: random.Random,
    n: int = SEQUENCE_LENGTH
) -> Dict[str, Any]:
    """
    Generates one valid random neighbourhood move.
    The move set and block-size limits match the LLM-guided optimiser.
    """
    action = choose_operator(rng)

    if action == "swap":
        i, j = rng.sample(range(n), 2)
        return {"action": "swap", "i": i, "j": j}

    if action == "move":
        src, dst = rng.sample(range(n), 2)
        return {"action": "move", "from": src, "to": dst}

    if action == "reverse":
        start, end = sorted(rng.sample(range(n), 2))
        return {"action": "reverse", "start": start, "end": end}

    if action == "reinsert_block":
        max_len = min(MAX_REINSERT_BLOCK_LEN, n - 1)
        block_len = rng.randint(MIN_REINSERT_BLOCK_LEN, max_len)
        start = rng.randint(0, n - block_len)
        end = start + block_len - 1

        valid_targets = [
            idx for idx in range(n)
            if not (start <= idx <= end)
        ]
        to = rng.choice(valid_targets)

        return {
            "action": "reinsert_block",
            "start": start,
            "end": end,
            "to": to,
        }

    if action == "swap_blocks":
        max_len = min(MAX_SWAP_BLOCK_LEN, n // 2)
        block_len = rng.randint(MIN_SWAP_BLOCK_LEN, max_len)

        # Generate until two disjoint equal-length blocks are found.
        for _ in range(1000):
            start1 = rng.randint(0, n - block_len)
            start2 = rng.randint(0, n - block_len)

            end1 = start1 + block_len - 1
            end2 = start2 + block_len - 1

            if end1 < start2 or end2 < start1:
                return {
                    "action": "swap_blocks",
                    "start1": start1,
                    "end1": end1,
                    "start2": start2,
                    "end2": end2,
                }

        raise RuntimeError("Could not generate disjoint swap_blocks move.")

    raise ValueError(f"Unsupported operator: {action}")


def generate_valid_random_move(
    rng: random.Random,
    n: int = SEQUENCE_LENGTH
) -> Dict[str, Any]:
    """
    Defensive wrapper. Random move generation should already be valid,
    but validation is retained so the SA experiment uses exactly the same
    legality rules as the LLM-guided optimiser.
    """
    for _ in range(1000):
        move = generate_random_move(rng, n=n)
        valid, _ = validate_move(move, n=n)
        if valid:
            return move

    raise RuntimeError("Unable to generate a valid neighbourhood move.")

# ============================================================
# SIMULATED ANNEALING
# ============================================================

def acceptance_probability(
    delta: float,
    temperature: float
) -> float:
    """
    For a minimization problem:
    - improvements/equal solutions are accepted with probability 1
    - worse solutions are accepted with exp(-delta / temperature)
    """
    if delta <= 0:
        return 1.0

    if temperature <= 0:
        return 0.0

    return math.exp(-delta / temperature)


def cool_temperature(current_temperature: float) -> float:
    return max(
        SA_MIN_TEMPERATURE,
        current_temperature * SA_COOLING_RATE
    )

# ============================================================
# LOGGING
# ============================================================

FIELDNAMES = [
    "evaluation",
    "random_seed",
    "temperature_before",
    "temperature_after",
    "action",
    "move_text",
    "swap_i",
    "swap_j",
    "move_from",
    "move_to",
    "reverse_start",
    "reverse_end",
    "reinsert_block_start",
    "reinsert_block_end",
    "reinsert_block_to",
    "swap_blocks_start1",
    "swap_blocks_end1",
    "swap_blocks_start2",
    "swap_blocks_end2",
    "current_makespan_before",
    "candidate_makespan",
    "delta_vs_current",
    "acceptance_probability",
    "acceptance_random_draw",
    "accepted",
    "acceptance_reason",
    "current_makespan_after",
    "global_best_makespan_after",
    "new_global_best",
    "completed_simulations",
    "exited_parts",
    "completion_valid",
    "simulation_time_sec",
    "sa_decision_time_sec",
    "inference_time_sec",
    "cumulative_simulation_time_sec",
    "cumulative_sa_decision_time_sec",
    "elapsed_optimization_time_sec",
    "status",
    "error_message",
    "final_best_sequence",
]

log_rows: List[Dict[str, Any]] = []


def clean_csv_text(value: Any) -> Any:
    if isinstance(value, str):
        return value.replace("\r", " ").replace("\n", " ")
    return value


def initialize_log_file() -> None:
    if not WRITE_LOGS:
        return

    with LOG_FILE.open("w", newline="", encoding=CSV_ENCODING) as f:
        writer = csv.DictWriter(
            f,
            fieldnames=FIELDNAMES,
            delimiter=CSV_DELIMITER
        )
        writer.writeheader()


def append_log_row(row: Dict[str, Any]) -> None:
    # Fill any omitted columns with None to keep summary rows simple.
    normalized = {field: row.get(field) for field in FIELDNAMES}
    cleaned = {
        key: clean_csv_text(value)
        for key, value in normalized.items()
    }

    log_rows.append(cleaned)

    if not WRITE_LOGS:
        return

    with LOG_FILE.open("a", newline="", encoding=CSV_ENCODING) as f:
        writer = csv.DictWriter(
            f,
            fieldnames=FIELDNAMES,
            delimiter=CSV_DELIMITER
        )
        writer.writerow(cleaned)


def write_global_best_sequence_file(
    sequence: List[str],
    makespan: float,
    random_seed: int,
    completed_simulations: int,
    total_elapsed_sec: float
) -> None:
    """
    Writes the exact 64-position best sequence and key run information
    to a separate text file for manual verification.
    """
    if not WRITE_LOGS:
        return

    with GLOBAL_BEST_SEQUENCE_FILE.open("w", encoding="utf-8") as f:
        f.write(f"Global best makespan: {makespan}\n")
        f.write(f"Random seed: {random_seed}\n")
        f.write(f"Completed simulations: {completed_simulations}\n")
        f.write(f"Optimization time excluding initial baseline: {total_elapsed_sec:.4f} s\n")
        f.write(f"Sequence length: {len(sequence)}\n\n")
        f.write("Position;Part\n")

        for idx, part in enumerate(sequence, start=1):
            f.write(f"{idx};{part}\n")


def write_run_summary(
    random_seed: int,
    initial_makespan: float,
    global_best_makespan: float,
    completed_simulations: int,
    candidate_evaluations: int,
    initial_simulation_time_sec: float,
    cumulative_candidate_simulation_time_sec: float,
    cumulative_sa_decision_time_sec: float,
    optimization_time_excluding_initial_sec: float,
    total_run_time_including_initial_sec: float,
    final_temperature: float
) -> None:
    if not WRITE_LOGS:
        return

    with RUN_SUMMARY_FILE.open("w", encoding="utf-8") as f:
        f.write("Simulated Annealing run summary\n")
        f.write("===============================\n")
        f.write(f"Random seed: {random_seed}\n")
        f.write(f"Initial makespan: {initial_makespan}\n")
        f.write(f"Global best makespan: {global_best_makespan}\n")
        f.write(f"Candidate evaluations: {candidate_evaluations}\n")
        f.write(f"Completed simulations including initial baseline: {completed_simulations}\n")
        f.write(f"Initial baseline simulation wall time: {initial_simulation_time_sec:.4f} s\n")
        f.write(
            "Cumulative candidate simulation wall time: "
            f"{cumulative_candidate_simulation_time_sec:.4f} s\n"
        )
        f.write(
            "Cumulative SA decision-generation wall time: "
            f"{cumulative_sa_decision_time_sec:.4f} s\n"
        )
        f.write("Inference time: 0.0000 s (not applicable to SA)\n")
        f.write(
            "Optimization time excluding initial baseline: "
            f"{optimization_time_excluding_initial_sec:.4f} s\n"
        )
        f.write(
            "Total run time including initial baseline: "
            f"{total_run_time_including_initial_sec:.4f} s\n"
        )
        f.write(f"Final SA temperature: {final_temperature}\n")
        f.write(f"Initial SA temperature: {SA_INITIAL_TEMPERATURE}\n")
        f.write(f"Cooling rate: {SA_COOLING_RATE}\n")
        f.write(f"Minimum SA temperature: {SA_MIN_TEMPERATURE}\n")
        f.write(f"Maximum candidate evaluations: {SA_MAX_EVALUATIONS}\n")
        f.write(f"Operator weights: {SA_OPERATOR_WEIGHTS}\n")

# ============================================================
# MAIN
# ============================================================

initialize_log_file()
print(f"SA CSV log file will be written to: {LOG_FILE.resolve()}")
print(f"SA best-sequence file will be written to: {GLOBAL_BEST_SEQUENCE_FILE.resolve()}")
print(f"SA run-summary file will be written to: {RUN_SUMMARY_FILE.resolve()}")

try:
    ps = win32com.client.Dispatch(
        "Tecnomatix.PlantSimulation.RemoteControl"
    )
    print("Connected via Tecnomatix.PlantSimulationRemoteControl")
except Exception as e:
    raise RuntimeError(
        f"Could not connect to Plant Simulation RemoteControl: {e}"
    )

overall_run_start = time.perf_counter()

ps.LoadModel(MODEL_PATH)
print("Model opened successfully")

# Initial production sequence - identical to the LLM-guided optimiser.
sequence = [
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH",
    "PartA", "PartB", "PartC", "PartD", "PartE", "PartF", "PartG", "PartH"
]

set_sequence(ps, sequence)
current_sequence = get_sequence(ps)
print("Initial sequence set.")

initial_sim_start = time.perf_counter()
current_makespan = run_simulation_until_finished(ps)
initial_simulation_time_sec = time.perf_counter() - initial_sim_start

initial_exited_parts = read_exited_parts(ps)
if not is_completion_valid(initial_exited_parts):
    raise RuntimeError(
        "Initial simulation did not complete all parts: "
        f"ExitedParts={initial_exited_parts}, "
        f"expected {EXPECTED_EXITED_PARTS}"
    )

initial_makespan = current_makespan
print("Initial makespan:", initial_makespan)
print("Initial ExitedParts:", initial_exited_parts)
print(
    "Initial Python-controlled simulation wall time:",
    round(initial_simulation_time_sec, 4),
    "s"
)

# Random seed management.
if SA_RANDOM_SEED is None:
    actual_random_seed = random.SystemRandom().randrange(0, 2**32)
else:
    actual_random_seed = int(SA_RANDOM_SEED)

rng = random.Random(actual_random_seed)
print("SA random seed:", actual_random_seed)

global_best_sequence = current_sequence.copy()
global_best_makespan = current_makespan

completed_simulations = 1
candidate_evaluations = 0
temperature = float(SA_INITIAL_TEMPERATURE)

cumulative_candidate_simulation_time_sec = 0.0
cumulative_sa_decision_time_sec = 0.0

# Kept consistent with the original LLM optimiser:
# "optimization time" starts after the initial baseline simulation.
optimization_start = time.perf_counter()

for evaluation in range(1, SA_MAX_EVALUATIONS + 1):
    candidate_evaluations += 1

    temperature_before = temperature
    current_makespan_before = current_makespan

    # This is the SA analogue of candidate-generation/inference time:
    # it measures random neighbourhood-move selection and application.
    decision_start = time.perf_counter()
    move = generate_valid_random_move(rng, n=SEQUENCE_LENGTH)
    candidate_sequence = apply_move(current_sequence, move)
    sa_decision_time_sec = time.perf_counter() - decision_start
    cumulative_sa_decision_time_sec += sa_decision_time_sec

    sim_start = time.perf_counter()

    try:
        set_sequence(ps, candidate_sequence)
        candidate_makespan = run_simulation_until_finished(ps)
        simulation_time_sec = time.perf_counter() - sim_start
        cumulative_candidate_simulation_time_sec += simulation_time_sec

        completed_simulations += 1

        exited_parts = read_exited_parts(ps)
        completion_valid = is_completion_valid(exited_parts)

        if not completion_valid:
            raise ValueError(
                "Simulation completed with "
                f"ExitedParts={exited_parts}, "
                f"expected {EXPECTED_EXITED_PARTS}."
            )

        delta = candidate_makespan - current_makespan_before
        accept_prob = acceptance_probability(
            delta,
            temperature_before
        )

        if delta <= 0:
            random_draw = None
            accepted = True
            acceptance_reason = (
                "ImprovedOrEqualCurrentSolution"
            )
        else:
            random_draw = rng.random()
            accepted = random_draw < accept_prob
            acceptance_reason = (
                "AcceptedWorseByTemperature"
                if accepted
                else "RejectedWorseCandidate"
            )

        if accepted:
            current_sequence = candidate_sequence.copy()
            current_makespan = candidate_makespan

        new_global_best = False
        if candidate_makespan < global_best_makespan:
            global_best_sequence = candidate_sequence.copy()
            global_best_makespan = candidate_makespan
            new_global_best = True
            print(
                f"Evaluation {evaluation}: new global best = "
                f"{global_best_makespan:.4f}"
            )

        temperature = cool_temperature(temperature_before)

        append_log_row({
            "evaluation": evaluation,
            "random_seed": actual_random_seed,
            "temperature_before": temperature_before,
            "temperature_after": temperature,
            "action": move["action"],
            "move_text": format_move(move),
            "swap_i": (
                move.get("i")
                if move["action"] == "swap"
                else None
            ),
            "swap_j": (
                move.get("j")
                if move["action"] == "swap"
                else None
            ),
            "move_from": (
                move.get("from")
                if move["action"] == "move"
                else None
            ),
            "move_to": (
                move.get("to")
                if move["action"] == "move"
                else None
            ),
            "reverse_start": (
                move.get("start")
                if move["action"] == "reverse"
                else None
            ),
            "reverse_end": (
                move.get("end")
                if move["action"] == "reverse"
                else None
            ),
            "reinsert_block_start": (
                move.get("start")
                if move["action"] == "reinsert_block"
                else None
            ),
            "reinsert_block_end": (
                move.get("end")
                if move["action"] == "reinsert_block"
                else None
            ),
            "reinsert_block_to": (
                move.get("to")
                if move["action"] == "reinsert_block"
                else None
            ),
            "swap_blocks_start1": (
                move.get("start1")
                if move["action"] == "swap_blocks"
                else None
            ),
            "swap_blocks_end1": (
                move.get("end1")
                if move["action"] == "swap_blocks"
                else None
            ),
            "swap_blocks_start2": (
                move.get("start2")
                if move["action"] == "swap_blocks"
                else None
            ),
            "swap_blocks_end2": (
                move.get("end2")
                if move["action"] == "swap_blocks"
                else None
            ),
            "current_makespan_before": current_makespan_before,
            "candidate_makespan": candidate_makespan,
            "delta_vs_current": delta,
            "acceptance_probability": accept_prob,
            "acceptance_random_draw": random_draw,
            "accepted": accepted,
            "acceptance_reason": acceptance_reason,
            "current_makespan_after": current_makespan,
            "global_best_makespan_after": global_best_makespan,
            "new_global_best": new_global_best,
            "completed_simulations": completed_simulations,
            "exited_parts": exited_parts,
            "completion_valid": completion_valid,
            "simulation_time_sec": round(simulation_time_sec, 6),
            "sa_decision_time_sec": round(sa_decision_time_sec, 6),
            # SA does not use an LLM; zero is written explicitly to simplify
            # later comparison/sensitivity-analysis spreadsheets.
            "inference_time_sec": 0.0,
            "cumulative_simulation_time_sec": round(
                cumulative_candidate_simulation_time_sec, 6
            ),
            "cumulative_sa_decision_time_sec": round(
                cumulative_sa_decision_time_sec, 6
            ),
            "elapsed_optimization_time_sec": round(
                time.perf_counter() - optimization_start, 4
            ),
            "status": "Evaluated",
            "error_message": None,
            "final_best_sequence": None,
        })

    except Exception as e:
        simulation_time_sec = time.perf_counter() - sim_start
        cumulative_candidate_simulation_time_sec += simulation_time_sec

        # Cool after a failed attempted evaluation so the evaluation counter
        # and temperature schedule remain aligned.
        temperature = cool_temperature(temperature_before)

        append_log_row({
            "evaluation": evaluation,
            "random_seed": actual_random_seed,
            "temperature_before": temperature_before,
            "temperature_after": temperature,
            "action": move["action"],
            "move_text": format_move(move),
            "current_makespan_before": current_makespan_before,
            "candidate_makespan": None,
            "delta_vs_current": None,
            "acceptance_probability": None,
            "acceptance_random_draw": None,
            "accepted": False,
            "acceptance_reason": "SimulationFailed",
            "current_makespan_after": current_makespan,
            "global_best_makespan_after": global_best_makespan,
            "new_global_best": False,
            "completed_simulations": completed_simulations,
            "exited_parts": None,
            "completion_valid": False,
            "simulation_time_sec": round(simulation_time_sec, 6),
            "sa_decision_time_sec": round(sa_decision_time_sec, 6),
            "inference_time_sec": 0.0,
            "cumulative_simulation_time_sec": round(
                cumulative_candidate_simulation_time_sec, 6
            ),
            "cumulative_sa_decision_time_sec": round(
                cumulative_sa_decision_time_sec, 6
            ),
            "elapsed_optimization_time_sec": round(
                time.perf_counter() - optimization_start, 4
            ),
            "status": "SimulationFailed",
            "error_message": str(e),
            "final_best_sequence": None,
        })

optimization_time_excluding_initial_sec = (
    time.perf_counter() - optimization_start
)
total_run_time_including_initial_sec = (
    time.perf_counter() - overall_run_start
)

print("\nFinal global best makespan:", global_best_makespan)
print("Initial makespan:", initial_makespan)
print("Candidate evaluations:", candidate_evaluations)
print("Completed simulations including initial baseline:", completed_simulations)
print(
    "Cumulative candidate simulation wall time:",
    round(cumulative_candidate_simulation_time_sec, 4),
    "s"
)
print(
    "Cumulative SA decision time:",
    round(cumulative_sa_decision_time_sec, 6),
    "s"
)
print(
    "Optimization time excluding initial baseline:",
    round(optimization_time_excluding_initial_sec, 4),
    "s"
)
print(
    "Total run time including initial baseline:",
    round(total_run_time_including_initial_sec, 4),
    "s"
)
print("Final temperature:", temperature)

write_global_best_sequence_file(
    global_best_sequence,
    global_best_makespan,
    actual_random_seed,
    completed_simulations,
    optimization_time_excluding_initial_sec
)

write_run_summary(
    random_seed=actual_random_seed,
    initial_makespan=initial_makespan,
    global_best_makespan=global_best_makespan,
    completed_simulations=completed_simulations,
    candidate_evaluations=candidate_evaluations,
    initial_simulation_time_sec=initial_simulation_time_sec,
    cumulative_candidate_simulation_time_sec=(
        cumulative_candidate_simulation_time_sec
    ),
    cumulative_sa_decision_time_sec=cumulative_sa_decision_time_sec,
    optimization_time_excluding_initial_sec=(
        optimization_time_excluding_initial_sec
    ),
    total_run_time_including_initial_sec=(
        total_run_time_including_initial_sec
    ),
    final_temperature=temperature
)

# Final summary row. The exact best sequence is included in the CSV so the
# final result is available both in the Excel-compatible log and in the TXT.
append_log_row({
    "evaluation": SA_MAX_EVALUATIONS,
    "random_seed": actual_random_seed,
    "temperature_before": None,
    "temperature_after": temperature,
    "action": None,
    "move_text": None,
    "current_makespan_before": None,
    "candidate_makespan": None,
    "delta_vs_current": None,
    "acceptance_probability": None,
    "acceptance_random_draw": None,
    "accepted": None,
    "acceptance_reason": None,
    "current_makespan_after": current_makespan,
    "global_best_makespan_after": global_best_makespan,
    "new_global_best": None,
    "completed_simulations": completed_simulations,
    "exited_parts": None,
    "completion_valid": None,
    "simulation_time_sec": None,
    "sa_decision_time_sec": None,
    "inference_time_sec": 0.0,
    "cumulative_simulation_time_sec": round(
        cumulative_candidate_simulation_time_sec, 6
    ),
    "cumulative_sa_decision_time_sec": round(
        cumulative_sa_decision_time_sec, 6
    ),
    "elapsed_optimization_time_sec": round(
        optimization_time_excluding_initial_sec, 4
    ),
    "status": "FinalGlobalBestSequence",
    "error_message": None,
    "final_best_sequence": " | ".join(global_best_sequence),
})

print(f"Final SA CSV log file: {LOG_FILE.resolve()}")
print(f"Global best sequence file: {GLOBAL_BEST_SEQUENCE_FILE.resolve()}")
print(f"SA run summary file: {RUN_SUMMARY_FILE.resolve()}")
