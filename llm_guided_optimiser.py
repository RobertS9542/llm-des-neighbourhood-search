import win32com.client
import json
import requests
import time
import csv
import hashlib
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

MODEL_PATH = r"CHANGE_THIS_TO_THE_LOCATION_OF_YOUR_MODEL"
LLM_MODEL = "mistral-small3.2:24b"
LLM_TEMPERATURE = 0.6

SEQUENCE_LENGTH = 64
EXPECTED_EXITED_PARTS = 64
NUM_ITERATIONS = 100
NUM_STATIONS = 20

# Normal search
BASE_CANDIDATES_PER_ITERATION = 5

# Diversification after stagnation
STAGNATION_THRESHOLD = 3
RESTART_STAGNATION_THRESHOLD = 10
DIVERSIFICATION_CANDIDATES_PER_ITERATION = 10

# Logging / CSV compatibility for Windows / Excel
CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"
WRITE_LOGS = True

# Polling settings for simulation completion detection
POLL_INTERVAL_SEC = 0.2
STABLE_POLLS_REQUIRED = 3
SIM_TIMEOUT_SEC = 300

# Memory settings
RECENT_REJECTED_LIMIT = 40
RECENT_MOVE_TEXT_LIMIT = 12
MAX_LLM_RETRIES_PER_CANDIDATE = 2

# Macro-move settings
MIN_REINSERT_BLOCK_LEN = 2
MAX_REINSERT_BLOCK_LEN = 12
MIN_SWAP_BLOCK_LEN = 2
MAX_SWAP_BLOCK_LEN = 8

# Resource statistics table settings
EXITED_PARTS_PATH = ".Models.Model.ExitedParts"
RESOURCE_STATS_TABLE_PATH = ".Models.Model.ResourceStatistics"
RESOURCE_WORKING_ROW = 1
RESOURCE_SETTINGUP_ROW = 2
RESOURCE_WAITING_ROW = 3
RESOURCE_BLOCKED_ROW = 4
RESOURCE_TOP_K = 3
UTILIZATION_TREND_TOP_K = 5
ACCEPTED_RESOURCE_HISTORY_LIMIT = 5

# Flat process-route table settings
FLAT_PROCESS_TABLE_PATH = ".Models.Model.FlatProcessTable"
FLAT_PROCESS_MAX_ROWS = 200
BOTTLENECK_PROCESS_TOP_K = 5

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
LOG_FILE = LOG_DIR / f"optimization_log_{RUN_TIMESTAMP}.csv"
RAW_LLM_LOG_FILE = LOG_DIR / f"raw_llm_replies_{RUN_TIMESTAMP}.jsonl"
GLOBAL_BEST_SEQUENCE_FILE = LOG_DIR / f"global_best_sequence_{RUN_TIMESTAMP}.txt"

# ============================================================
# CORE HELPERS
# ============================================================

def set_sequence(ps, sequence: List[str]) -> None:
    for i, part_name in enumerate(sequence, start=1):
        ps.ExecuteSimTalk(f'.Models.Model.SourceSequence[1, {i}] := .UserObjects.{part_name}')
    for i, part_name in enumerate(sequence, start=1):
        ps.ExecuteSimTalk(f'.Models.Model.SourceSequence[3, {i}] := "{part_name}"')


def get_sequence(ps, length: int = SEQUENCE_LENGTH) -> List[str]:
    sequence = []
    for i in range(1, length + 1):
        value = ps.GetValue(f".Models.Model.SourceSequence[1, {i}]")
        sequence.append(str(value).split(".")[-1])
    return sequence


def call_llm(prompt: str, model: str = LLM_MODEL, temperature: float = LLM_TEMPERATURE) -> str:
    response = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": temperature,
                "num_predict": 80,
                "num_ctx": 8192
            }
        },
        timeout=120
    )
    response.raise_for_status()
    return response.json()["response"]


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
        return f'reinsert_block({move.get("start")} .. {move.get("end")} -> {move.get("to")})'
    if action == "swap_blocks":
        return f'swap_blocks({move.get("start1")} .. {move.get("end1")}, {move.get("start2")} .. {move.get("end2")})'

    return str(move)


def summarize_recent_history(log_rows: List[Dict[str, Any]], max_items: int = 5) -> str:
    if not log_rows:
        return "No previous moves yet."

    recent = log_rows[-max_items:]
    lines = []
    for idx, row in enumerate(recent, start=1):
        if row.get("action") == "reinsert_block":
            move_dict = {
                "action": "reinsert_block",
                "start": row.get("reinsert_block_start"),
                "end": row.get("reinsert_block_end"),
                "to": row.get("reinsert_block_to"),
            }
        elif row.get("action") == "swap_blocks":
            move_dict = {
                "action": "swap_blocks",
                "start1": row.get("swap_blocks_start1"),
                "end1": row.get("swap_blocks_end1"),
                "start2": row.get("swap_blocks_start2"),
                "end2": row.get("swap_blocks_end2"),
            }
        else:
            move_dict = {
                "action": row.get("action"),
                "i": row.get("swap_i"),
                "j": row.get("swap_j"),
                "from": row.get("move_from"),
                "to": row.get("move_to"),
                "start": row.get("reverse_start"),
                "end": row.get("reverse_end"),
            }

        move_desc = format_move(move_dict)
        result = row.get("status", "Unknown")
        delta = row.get("delta_vs_best_before")
        delta_text = f", delta vs best before: {delta:.3f}" if isinstance(delta, (int, float)) else ""
        lines.append(f"{idx}. {move_desc} -> {result}{delta_text}")

    return "\n".join(lines)


def summarize_recent_rejected_moves(recent_rejected_entries: List[Dict[str, str]], max_items: int = 12) -> str:
    if not recent_rejected_entries:
        return "None"

    items = recent_rejected_entries[-max_items:]
    lines = []
    for idx, item in enumerate(items, start=1):
        lines.append(f"{idx}. {item['move_text']} (from parent seq {item['parent_signature_short']})")
    return "\n".join(lines)


def get_recent_move_texts(log_rows: List[Dict[str, Any]], max_items: int = RECENT_MOVE_TEXT_LIMIT) -> List[str]:
    if not log_rows:
        return []

    recent = log_rows[-max_items:]
    texts = []
    for row in recent:
        if row.get("action") == "reinsert_block":
            move_dict = {
                "action": "reinsert_block",
                "start": row.get("reinsert_block_start"),
                "end": row.get("reinsert_block_end"),
                "to": row.get("reinsert_block_to"),
            }
        elif row.get("action") == "swap_blocks":
            move_dict = {
                "action": "swap_blocks",
                "start1": row.get("swap_blocks_start1"),
                "end1": row.get("swap_blocks_end1"),
                "start2": row.get("swap_blocks_start2"),
                "end2": row.get("swap_blocks_end2"),
            }
        else:
            move_dict = {
                "action": row.get("action"),
                "i": row.get("swap_i"),
                "j": row.get("swap_j"),
                "from": row.get("move_from"),
                "to": row.get("move_to"),
                "start": row.get("reverse_start"),
                "end": row.get("reverse_end"),
            }
        texts.append(format_move(move_dict))
    return texts


def build_recent_move_text_block(log_rows: List[Dict[str, Any]], max_items: int = RECENT_MOVE_TEXT_LIMIT) -> str:
    texts = get_recent_move_texts(log_rows, max_items=max_items)
    if not texts:
        return "None"
    return "\n".join(f"- {t}" for t in texts)


def get_region_instruction(candidate_idx: int, search_mode: str) -> str:
    normal_regions = [
        (0, 15),
        (16, 31),
        (32, 47),
        (48, 63),
        None,
    ]
    diversification_regions = [
        (0, 15),
        (16, 31),
        (32, 47),
        (48, 63),
        (0, 31),
        (32, 63),
        None,
        (8, 23),
        (24, 39),
        (40, 55),
    ]
    if search_mode == "normal":
        regions = normal_regions
    else:
        regions = diversification_regions
    region = regions[candidate_idx - 1] if candidate_idx <= len(regions) else None

    if region is None:
        return (
            "This candidate must focus on a DIFFERENT region of the sequence than the most recent repeated patterns. "
            "Avoid reusing the same from/to indices or block boundaries seen recently."
        )

    lo, hi = region
    return (
        f"This candidate should focus primarily on indices in the region [{lo}..{hi}]. "
        f"Prefer using from/to indices or block boundaries that touch this region. "
        f"If using reinsert_block inside this region, choose a smaller valid sub-block, not the entire region. "
        f"The hard maximum allowed reinsert_block size is 12 positions. Choose the block size based on the intended structural effect. Avoid repeatedly using similar block sizes or the same block boundaries. "
        f"Avoid repeatedly reusing the same indices seen in recent candidates outside this region unless clearly necessary."
    )


def get_candidate_instruction(candidate_idx: int, search_mode: str) -> str:
    """
    Rebalanced action strategy + region targeting:
    - strongly reduce reverse dominance
    - emphasize move / swap / reinsert_block
    - force exploration across different index regions
    """

    if search_mode == "normal":
        instructions = [
            "This candidate must prefer the action type 'move'. Use a targeted move that may relieve blocking or shift load away from a bottleneck.",
            "This candidate must prefer the action type 'swap'. Use a targeted swap if it could locally improve flow or reduce pressure on a problematic station.",
            "This candidate must prefer the action type 'reinsert_block'. Use it only if moving a block clearly improves structure without causing chaos.",
            "This candidate may use 'move' or 'swap', but should avoid reverse unless no better targeted action is plausible.",
            "This candidate may use 'reverse' only as a last resort. Do not propose reverse if earlier candidates in this iteration already suggest a reasonable move, swap, or reinsert_block.",
        ]
    else:
        instructions = [
            "This candidate must prefer the action type 'move'. Use a long-distance move that may relieve blocking, reduce waiting, or shift load away from the most utilized station.",
            "This candidate must prefer the action type 'reinsert_block'. Use it only if relocating a contiguous block could relieve a bottleneck or reduce a queueing pattern.",
            "This candidate must prefer the action type 'swap'. Use a non-local but targeted swap. Avoid reverse here.",
            "This candidate may use 'move' or 'reinsert_block'. Choose the more targeted action and avoid reverse if possible.",
            "This candidate must prefer the least-used action type among move, swap, and reinsert_block. Avoid repeating reverse-dominated patterns.",
            "This candidate may use 'swap_blocks' only if a major structural change seems clearly justified by the utilization summary.",
            "This candidate may use 'reverse' only if there is a clear reason to reorder an entire local region and simpler targeted actions seem weaker.",
            "This candidate should target the most blocked stations using move, swap, or reinsert_block. Avoid reverse unless absolutely necessary.",
            "This candidate should target the most waiting stations using move or swap before considering reverse.",
            "This candidate must produce a valid move different in action type from the dominant recent pattern. Prefer move / swap / reinsert_block over reverse.",
        ]

    base = instructions[candidate_idx - 1] if candidate_idx <= len(instructions) else (
        "Prefer a valid move that is different in style from the previous candidates, and avoid reverse unless it is clearly justified."
    )
    region = get_region_instruction(candidate_idx, search_mode)
    return f"{base} {region}"


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def pct_str(x: float) -> str:
    return f"{round(x * 100, 1)}%"


def read_resource_statistics(ps, num_stations: int = NUM_STATIONS) -> List[Dict[str, Any]]:
    """
    Reads the ResourceStatistics table filled by Plant Simulation.
    Important: from Python/COM, table indexing is [column,row].
    So [1,3] means first column, third row.
    Columns are stations, rows are state categories.
    """
    stats = []

    for station_col in range(1, num_stations + 1):
        working = safe_float(ps.GetValue(f"{RESOURCE_STATS_TABLE_PATH}[{station_col},{RESOURCE_WORKING_ROW}]"))
        setting_up = safe_float(ps.GetValue(f"{RESOURCE_STATS_TABLE_PATH}[{station_col},{RESOURCE_SETTINGUP_ROW}]"))
        waiting = safe_float(ps.GetValue(f"{RESOURCE_STATS_TABLE_PATH}[{station_col},{RESOURCE_WAITING_ROW}]"))
        blocked = safe_float(ps.GetValue(f"{RESOURCE_STATS_TABLE_PATH}[{station_col},{RESOURCE_BLOCKED_ROW}]"))

        station_name = "Station" if station_col == 1 else f"Station{station_col - 1}"

        stats.append({
            "station_index": station_col,
            "station_name": station_name,
            "working": working,
            "setting_up": setting_up,
            "waiting": waiting,
            "blocked": blocked,
        })

    return stats


def normalize_station_name(value: Any) -> str:
    text = str(value).strip()
    if "." in text:
        text = text.split(".")[-1]
    return text


def read_flat_process_table(ps, max_rows: int = FLAT_PROCESS_MAX_ROWS) -> Tuple[Dict[str, List[Dict[str, Any]]], Dict[str, List[Dict[str, Any]]]]:
    """
    Reads .Models.Model.FlatProcessTable once from Plant Simulation.

    Important COM indexing convention:
    [column,row], so [1,3] means first column, third row.

    Expected columns:
    1 = Parts
    2 = Steps
    3 = Stations
    4 = Process time
    """
    part_to_route: Dict[str, List[Dict[str, Any]]] = {}
    station_to_processes: Dict[str, List[Dict[str, Any]]] = {}

    for row in range(1, max_rows + 1):
        try:
            part_raw = ps.GetValue(f"{FLAT_PROCESS_TABLE_PATH}[1,{row}]")
        except Exception:
            break

        if part_raw is None:
            break

        part = str(part_raw).strip()
        if part == "" or part.lower() == "void":
            break

        try:
            step_raw = ps.GetValue(f"{FLAT_PROCESS_TABLE_PATH}[2,{row}]")
            station_raw = ps.GetValue(f"{FLAT_PROCESS_TABLE_PATH}[3,{row}]")
            proc_time_raw = ps.GetValue(f"{FLAT_PROCESS_TABLE_PATH}[4,{row}]")
        except Exception:
            continue

        station = normalize_station_name(station_raw)
        try:
            step = int(float(step_raw))
        except Exception:
            step = None

        proc_time_text = str(proc_time_raw).strip()

        entry = {
            "part": part,
            "step": step,
            "station": station,
            "process_time": proc_time_text,
        }

        part_to_route.setdefault(part, []).append(entry)
        station_to_processes.setdefault(station, []).append(entry)

    for route in part_to_route.values():
        route.sort(key=lambda x: x["step"] if x["step"] is not None else 999)

    for entries in station_to_processes.values():
        entries.sort(key=lambda x: (x["part"], x["step"] if x["step"] is not None else 999))

    return part_to_route, station_to_processes


def build_bottleneck_process_summary(
    resource_stats: List[Dict[str, Any]],
    station_to_processes: Dict[str, List[Dict[str, Any]]],
    top_k: int = BOTTLENECK_PROCESS_TOP_K
) -> str:
    """
    Links current bottleneck stations to the part types that visit them.
    This helps the LLM connect resource symptoms to actionable sequence changes.
    """
    if not resource_stats or not station_to_processes:
        return "Bottleneck-to-process summary:\nNo process-route mapping available."

    selected_station_names = []

    for key in ["blocked", "working", "waiting"]:
        top_items = sorted(resource_stats, key=lambda x: x[key], reverse=True)[:top_k]
        for item in top_items:
            name = item["station_name"]
            if name not in selected_station_names:
                selected_station_names.append(name)

    lines = []
    lines.append("Bottleneck-to-process summary:")
    lines.append("These part types visit the currently important stations. Processing times are shown for the relevant station.")

    for station_name in selected_station_names[:top_k * 2]:
        entries = station_to_processes.get(station_name, [])
        stat = next((s for s in resource_stats if s["station_name"] == station_name), None)
        if stat:
            stat_text = f"working {pct_str(stat['working'])}, waiting {pct_str(stat['waiting'])}, blocked {pct_str(stat['blocked'])}"
        else:
            stat_text = "current stats unavailable"

        if not entries:
            lines.append(f"- {station_name} ({stat_text}): no matching part route found in FlatProcessTable")
            continue

        part_items = []
        for e in entries:
            step_text = f"step {e['step']}" if e["step"] is not None else "step ?"
            part_items.append(f"{e['part']}({step_text}, {e['process_time']})")

        lines.append(f"- {station_name} ({stat_text}): " + ", ".join(part_items))

    lines.append("")
    lines.append("Process-route guidance:")
    lines.append("- When a station has high blocking or waiting, consider sequence moves involving the part types that visit that station.")
    lines.append("- Longer processing times at a bottleneck station can make those part types especially important for sequencing.")
    lines.append("- Use this information as supporting context; makespan remains the primary objective.")

    return "\n".join(lines)


def build_resource_summary(stats: List[Dict[str, Any]], top_k: int = RESOURCE_TOP_K) -> str:
    if not stats:
        return "No resource statistics available."

    top_working = sorted(stats, key=lambda x: x["working"], reverse=True)[:top_k]
    top_blocked = sorted(stats, key=lambda x: x["blocked"], reverse=True)[:top_k]
    top_waiting = sorted(stats, key=lambda x: x["waiting"], reverse=True)[:top_k]

    lines = []
    lines.append("Resource utilization summary:")
    lines.append("")
    lines.append("Top utilized stations:")
    for idx, item in enumerate(top_working, start=1):
        lines.append(f"{idx}. {item['station_name']} - working {pct_str(item['working'])}, waiting {pct_str(item['waiting'])}, blocked {pct_str(item['blocked'])}")

    lines.append("")
    lines.append("Most blocked stations:")
    for idx, item in enumerate(top_blocked, start=1):
        lines.append(f"{idx}. {item['station_name']} - blocked {pct_str(item['blocked'])}, working {pct_str(item['working'])}, waiting {pct_str(item['waiting'])}")

    lines.append("")
    lines.append("Most waiting stations:")
    for idx, item in enumerate(top_waiting, start=1):
        lines.append(f"{idx}. {item['station_name']} - waiting {pct_str(item['waiting'])}, working {pct_str(item['working'])}, blocked {pct_str(item['blocked'])}")

    return "\n".join(lines)


def build_utilization_trend_summary(
    previous_stats: Optional[List[Dict[str, Any]]],
    current_stats: Optional[List[Dict[str, Any]]],
    top_k: int = UTILIZATION_TREND_TOP_K
) -> str:
    """
    Builds a compact trend summary comparing the current accepted state
    with the previous accepted state. This gives the LLM direction-of-change
    information without dumping several full historical summaries.
    """
    if not previous_stats or not current_stats:
        return "Utilization trend since previous accepted state:\nNo previous accepted utilization state available yet."

    prev_by_station = {item["station_name"]: item for item in previous_stats}
    rows = []

    for cur in current_stats:
        name = cur["station_name"]
        prev = prev_by_station.get(name)
        if prev is None:
            continue

        rows.append({
            "station_name": name,
            "working_delta": cur["working"] - prev["working"],
            "waiting_delta": cur["waiting"] - prev["waiting"],
            "blocked_delta": cur["blocked"] - prev["blocked"],
            "current_working": cur["working"],
            "current_waiting": cur["waiting"],
            "current_blocked": cur["blocked"],
        })

    if not rows:
        return "Utilization trend since previous accepted state:\nNo comparable station statistics available."

    def pp(x: float) -> str:
        sign = "+" if x >= 0 else ""
        return f"{sign}{round(x * 100, 1)} pp"

    most_blocking_increase = sorted(rows, key=lambda x: x["blocked_delta"], reverse=True)[:top_k]
    most_waiting_increase = sorted(rows, key=lambda x: x["waiting_delta"], reverse=True)[:top_k]
    most_working_increase = sorted(rows, key=lambda x: x["working_delta"], reverse=True)[:top_k]

    lines = []
    lines.append("Utilization trend since previous accepted state:")
    lines.append("")
    lines.append("Largest blocking increases:")
    for idx, item in enumerate(most_blocking_increase, start=1):
        lines.append(
            f"{idx}. {item['station_name']} - blocked {pp(item['blocked_delta'])}, "
            f"waiting {pp(item['waiting_delta'])}, working {pp(item['working_delta'])}"
        )

    lines.append("")
    lines.append("Largest waiting increases:")
    for idx, item in enumerate(most_waiting_increase, start=1):
        lines.append(
            f"{idx}. {item['station_name']} - waiting {pp(item['waiting_delta'])}, "
            f"blocked {pp(item['blocked_delta'])}, working {pp(item['working_delta'])}"
        )

    lines.append("")
    lines.append("Largest working-utilization increases:")
    for idx, item in enumerate(most_working_increase, start=1):
        lines.append(
            f"{idx}. {item['station_name']} - working {pp(item['working_delta'])}, "
            f"blocked {pp(item['blocked_delta'])}, waiting {pp(item['waiting_delta'])}"
        )

    lines.append("")
    lines.append("Trend guidance:")
    lines.append("- Treat increases in blocking/waiting as warning signals unless makespan improved strongly.")
    lines.append("- Treat increases in working utilization as potentially good, but watch whether they create new bottlenecks.")
    lines.append("- Use this trend only as supporting evidence; makespan remains the primary objective.")

    return "\n".join(lines)


def build_prompt(
    sequence: List[str],
    current_makespan: float,
    best_makespan: float,
    recent_history_text: str,
    rejected_moves_text: str,
    action_counts_text: str,
    no_improvement_count: int,
    search_mode: str,
    candidate_idx: int,
    resource_summary_text: str,
    utilization_trend_text: str,
    bottleneck_process_summary_text: str,
    recent_move_text_block: str
) -> str:
    candidate_instruction = get_candidate_instruction(candidate_idx, search_mode)

    diversification_block = ""
    if search_mode == "diversification":
        diversification_block = f"""
IMPORTANT SEARCH STATE:
- No improvement has been found for {no_improvement_count} iterations.
- The search may be stuck in a local minimum.
- You should now prioritize diversification, but still avoid meaningless disruption.
- In this state, prefer targeted move / swap / reinsert_block actions before using reverse.
"""
    

    return f"""You are helping optimize a manufacturing sequence.

Current sequence:
{sequence}

Current makespan in seconds:
{current_makespan}

Best makespan found so far:
{best_makespan}

Recent move history:
{recent_history_text}

Recently rejected parent-sequence-aware moves to avoid repeating immediately:
{rejected_moves_text}

Recent move texts (use these to avoid index-level repetition):
{recent_move_text_block}

Action usage so far:
{action_counts_text}

Current search mode:
{search_mode}

Candidate-specific instruction:
{candidate_instruction}

{resource_summary_text}

{utilization_trend_text}

{bottleneck_process_summary_text}

IMPORTANT:

The goal is to reduce makespan by proposing exactly ONE valid modification.

You are performing iterative search.
A good move should either:
- improve the sequence locally,
- or introduce a useful structural change without being unnecessarily disruptive.

Interpretation of utilization data:
- Higher working utilization often means a station is being used effectively, but extremely high working utilization may also indicate a bottleneck.
- Higher blocked percentage is generally undesirable and may indicate downstream congestion.
- Higher waiting percentage is generally undesirable and may indicate starvation, poor flow balance, or upstream sequencing problems.
- These are heuristic signals only; the primary objective remains reducing total makespan.

Use the resource utilization summary, utilization trend, and bottleneck-to-process summary to reason about bottlenecks, blocking, waiting, and which part types are likely involved.
Prefer targeted moves that may:
- reduce unnecessary blocking at the most blocked stations,
- reduce excessive waiting/starvation,
- improve flow balance,
- or shift load away from the most heavily utilized bottleneck stations without creating new congestion.

Very important action guidance:
- Do NOT overuse reverse.
- Do NOT overuse reinsert_block.
- Reverse and reinsert_block are not default actions.
- Single-element move is often a strong operator for fine-grained improvement; do not neglect it in favor of block-based operations.
- Prefer move or swap when targeting a specific bottleneck, blocked station, or waiting station.
- Use reinsert_block only when relocating a block has a clear structural reason.
- Use reverse only when there is a clear reason to reorder an entire local region and simpler targeted actions seem weaker.
- If recent candidates already rely on reverse or reinsert_block, strongly prefer move or swap.
- Balance exploration and stability: propose a move that is meaningfully different from recent patterns, but do not make chaotic changes unsupported by the resource data.
- When utilization points to a problematic station, use the bottleneck-to-process summary to identify which part types visit that station, then consider moves affecting those part types in the sequence.

CRITICAL DIVERSITY RULE:
- Do NOT repeatedly use the same indices or positions as in recent candidates.
- Avoid proposing moves that involve the same 'from' index, the same 'to' index, or the same block boundaries (start/end) if similar moves were already proposed recently.
- You must explore different parts of the sequence.
- If recent candidates focused on a specific region, you must propose a move involving a DIFFERENT region unless there is a very strong reason not to.
- Avoid repeating specific patterns like repeatedly moving from the same index or repeatedly manipulating the same short block.
- For reinsert_block, avoid repeatedly using the same block start/end/to combination. If a similar block move appeared recently, choose a different sub-block, a different target, or preferably a different action type such as move or swap.
- Across candidates in the same iteration, avoid generating multiple reinsert_block candidates unless there is a clear and different structural reason for each one.

Avoid random or redundant changes.
Avoid repeatedly using the same action type.
Avoid changes that only reshuffle repeating sub-patterns without creating meaningful structural difference.

---

SEARCH MODE STRATEGY:

Normal mode:
- The search is still improving or exploring effectively.
- Prefer small and medium targeted changes.
- Focus on refinement rather than disruption.

Preferred actions in normal mode:
1. move
2. swap
3. reinsert_block or reverse only when clearly justified

Use reverse only occasionally if there is a clear structural reason.
Avoid swap_blocks in normal mode unless clearly justified.

Guidelines:
- Prefer targeted improvements.
- Avoid overly disruptive block-level changes early.
- Do not default to macro-moves.
- If utilization data indicates a bottleneck, favor a targeted move, swap, or reinsert_block that may relieve that bottleneck.

---

Diversification mode:
- No improvement has been found for {no_improvement_count} iterations.
- The search may be stuck in a local minimum.

Preferred actions in diversification mode:
1. move
2. swap
3. reinsert_block
4. reverse
5. swap_blocks (rare)

Use swap_blocks only if:
- simpler moves repeatedly fail,
- and a major structural change is clearly necessary.

Guidelines:
- Prefer larger changes than in normal mode, but still avoid meaningless reshuffling.
- For move: prefer longer-distance moves.
- For reinsert_block: use it only when moving a block clearly changes the sequence structure. Avoid using reinsert_block repeatedly across candidates in the same iteration.
- For swap: prefer non-local but targeted swaps.
- For move: single-element moves are useful for fine-grained improvements and should remain frequent.
- For reverse: use it sparingly and only with clear justification.
- If utilization data indicates strong blocking or waiting, prefer moves that specifically improve flow through those stations.

{diversification_block}

---

ALLOWED ACTIONS:

1. swap two positions
2. move one element to a new position
3. reverse one contiguous block
4. reinsert one contiguous block at a new position
5. swap two disjoint contiguous blocks of equal length

IMPORTANT VALIDITY CONSTRAINTS FOR BLOCK MOVES:

For reinsert_block:
- The block length must be between 2 and 12 positions.
- Choose the block size based on the intended structural effect.
- Do NOT repeatedly use similar block sizes or the same block boundaries.
- Do NOT move an entire 16-position region such as 0..15, 16..31, 32..47, or 48..63.
- If the candidate instruction focuses on a region such as [16..31], choose a valid sub-block inside that region, but do not always choose the same size or same boundaries.

For swap_blocks:
- The two blocks must be disjoint.
- The two blocks must have equal length.
- Each block length must be between 2 and 8 positions.

---

RETURN FORMAT (STRICT):

Return exactly ONE JSON object, for example:

{{"action":"swap","i":5,"j":19}}
{{"action":"move","from":12,"to":3}}
{{"action":"reverse","start":20,"end":27}}
{{"action":"reinsert_block","start":10,"end":15,"to":40}}
{{"action":"swap_blocks","start1":4,"end1":7,"start2":20,"end2":23}}

---

RULES:

- Use 0-based indices
- All indices must be between 0 and 63
- Return ONLY the JSON object
- Do NOT include explanations
- Do NOT include code
- Do NOT return the full sequence
- Ensure the move is valid
- Avoid repeating recently rejected moves
- Prefer meaningful structural improvements over random edits
"""


def parse_llm_output(output: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    try:
        start = output.find("{")
        end = output.rfind("}") + 1

        if start == -1 or end == 0:
            return None, "No JSON object found"

        dict_str = output[start:end]
        move = json.loads(dict_str)

        if not isinstance(move, dict):
            return None, "Parsed JSON is not a dict"

        return move, None

    except Exception as e:
        return None, f"Parsing error: {e}"


def generate_candidate_with_retries(
    base_prompt: str,
    parent_sequence: List[str],
    seen_moves_by_parent: set,
    recent_rejected_keys: set,
    max_retries: int = MAX_LLM_RETRIES_PER_CANDIDATE
) -> Tuple[str, Optional[Dict[str, Any]], Optional[str], str, bool, int, Optional[str], bool, bool]:
    """
    Calls the LLM with a small retry loop before simulation.

    Retries are used only for:
    - parsing errors / non-JSON output
    - invalid moves
    - exact duplicate moves for the same parent sequence
    - recently rejected exact moves for the same parent sequence

    Diversification compatibility is intentionally NOT handled here.
    That remains part of the later evaluation/filtering logic.

    Returns:
    reply, move, parse_error, validation_message, valid,
    llm_attempts, retry_reason, duplicate_for_same_parent, rejected_for_same_parent
    """
    prompt = base_prompt
    last_reply = ""
    last_move = None
    last_parse_error = None
    last_validation_message = ""
    last_valid = False
    last_retry_reason = None
    last_duplicate = False
    last_recently_rejected = False

    for attempt in range(1, max_retries + 2):
        last_reply = call_llm(prompt)
        last_parse_error = None
        last_validation_message = ""
        last_valid = False
        last_duplicate = False
        last_recently_rejected = False

        move, parse_problem = parse_llm_output(last_reply)
        last_move = move

        if parse_problem:
            last_parse_error = parse_problem
            last_retry_reason = "parse_error"
        else:
            last_valid, last_validation_message = validate_move(move)
            if not last_valid:
                last_retry_reason = "invalid_move"
            else:
                parent_aware_key = build_parent_aware_key(parent_sequence, move)
                last_duplicate = parent_aware_key in seen_moves_by_parent
                last_recently_rejected = parent_aware_key in recent_rejected_keys

                if last_duplicate:
                    last_retry_reason = "duplicate_for_same_parent_sequence"
                elif last_recently_rejected:
                    last_retry_reason = "recently_rejected_for_same_parent_sequence"
                else:
                    return (
                        last_reply,
                        last_move,
                        None,
                        last_validation_message,
                        True,
                        attempt,
                        last_retry_reason,
                        False,
                        False,
                    )

        if attempt <= max_retries:
            if last_retry_reason == "parse_error":
                retry_note = f"""

Your previous response could not be parsed as a valid JSON move.
Reason: {last_parse_error}

Propose a different valid move.
Return ONLY one JSON object.
"""
            elif last_retry_reason == "invalid_move":
                retry_note = f"""

Your previous move was invalid.
Reason: {last_validation_message}
Previous proposal: {last_reply}

Propose a corrected valid move with legal indices and legal block sizes.
Return ONLY one JSON object.
"""
            elif last_retry_reason == "duplicate_for_same_parent_sequence":
                retry_note = f"""

Your previous move duplicated a move already tested from this exact same parent sequence.
Duplicate proposal: {format_move(last_move)}

Do NOT repeat this move.
Propose a different valid move with different indices, and preferably a different action type.
Return ONLY one JSON object.
"""
            elif last_retry_reason == "recently_rejected_for_same_parent_sequence":
                retry_note = f"""

Your previous move matches a recently rejected move from this exact same parent sequence.
Rejected proposal: {format_move(last_move)}

Do NOT repeat this move.
Propose a different valid move with different indices, and preferably a different action type.
Return ONLY one JSON object.
"""
            else:
                retry_note = "\n\nPropose a different valid JSON move. Return ONLY one JSON object."

            prompt = base_prompt + retry_note

    # If the final attempt is still an exact duplicate or recently rejected move,
    # do not simulate it. Exact same parent-aware moves are deterministic and
    # re-running them only wastes simulation time.
    final_valid = last_valid and not last_duplicate and not last_recently_rejected

    return (
        last_reply,
        last_move,
        last_parse_error,
        last_validation_message,
        final_valid,
        max_retries + 1,
        last_retry_reason,
        last_duplicate,
        last_recently_rejected,
    )


def validate_move(move: Optional[Dict[str, Any]], n: int = SEQUENCE_LENGTH) -> Tuple[bool, str]:
    if move is None:
        return False, "Move is None"

    action = move.get("action")
    if action not in {"swap", "move", "reverse", "reinsert_block", "swap_blocks"}:
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
            if not (MIN_REINSERT_BLOCK_LEN <= block_len <= MAX_REINSERT_BLOCK_LEN):
                return False, "reinsert_block block length out of allowed range"
            if start <= to <= end:
                return False, "reinsert_block target cannot be inside the block"
            return True, "ok"

        if action == "swap_blocks":
            start1 = int(move["start1"])
            end1 = int(move["end1"])
            start2 = int(move["start2"])
            end2 = int(move["end2"])

            if not all(0 <= x < n for x in [start1, end1, start2, end2]):
                return False, "swap_blocks indices out of range"
            if start1 >= end1 or start2 >= end2:
                return False, "swap_blocks each block must have start < end"

            len1 = end1 - start1 + 1
            len2 = end2 - start2 + 1
            if len1 != len2:
                return False, "swap_blocks block lengths must match"
            if not (MIN_SWAP_BLOCK_LEN <= len1 <= MAX_SWAP_BLOCK_LEN):
                return False, "swap_blocks block length out of allowed range"
            if not (end1 < start2 or end2 < start1):
                return False, "swap_blocks blocks must be disjoint"

            return True, "ok"

    except KeyError as e:
        return False, f"Missing key: {e}"
    except Exception as e:
        return False, f"Validation error: {e}"

    return False, "Unhandled validation case"


def is_diversification_move(move: Dict[str, Any]) -> bool:
    action = move["action"]

    if action == "swap":
        return abs(int(move["i"]) - int(move["j"])) >= 12
    if action == "move":
        return abs(int(move["from"]) - int(move["to"])) >= 10
    if action == "reverse":
        length = int(move["end"]) - int(move["start"]) + 1
        return 8 <= length <= 20
    if action == "reinsert_block":
        start = int(move["start"])
        end = int(move["end"])
        to = int(move["to"])
        length = end - start + 1
        move_distance = min(abs(to - start), abs(to - end))
        return 3 <= length <= 12 and move_distance >= 8
    if action == "swap_blocks":
        start1 = int(move["start1"])
        end1 = int(move["end1"])
        start2 = int(move["start2"])
        end2 = int(move["end2"])
        length = end1 - start1 + 1
        separation = abs(start2 - start1)
        return 2 <= length <= 8 and separation >= 10 and (end1 < start2 or end2 < start1)

    return False


def canonical_move_key(move: Dict[str, Any]) -> str:
    action = move["action"]

    if action == "swap":
        i, j = sorted([int(move["i"]), int(move["j"])] )
        return json.dumps({"action": "swap", "i": i, "j": j}, sort_keys=True)

    if action == "move":
        return json.dumps({"action": "move", "from": int(move["from"]), "to": int(move["to"])} , sort_keys=True)

    if action == "reverse":
        return json.dumps({"action": "reverse", "start": int(move["start"]), "end": int(move["end"])} , sort_keys=True)

    if action == "reinsert_block":
        return json.dumps({
            "action": "reinsert_block",
            "start": int(move["start"]),
            "end": int(move["end"]),
            "to": int(move["to"])
        }, sort_keys=True)

    if action == "swap_blocks":
        block1 = (int(move["start1"]), int(move["end1"])
        )
        block2 = (int(move["start2"]), int(move["end2"])
        )
        ordered = sorted([block1, block2], key=lambda x: x[0])
        return json.dumps({
            "action": "swap_blocks",
            "start1": ordered[0][0],
            "end1": ordered[0][1],
            "start2": ordered[1][0],
            "end2": ordered[1][1]
        }, sort_keys=True)

    return json.dumps(move, sort_keys=True)


def get_sequence_signature(sequence: List[str]) -> str:
    joined = "|".join(sequence)
    return hashlib.sha1(joined.encode("utf-8")).hexdigest()


def build_parent_aware_key(parent_sequence: List[str], move: Dict[str, Any]) -> str:
    parent_sig = get_sequence_signature(parent_sequence)
    move_key = canonical_move_key(move)
    return f"{parent_sig}::{move_key}"


def apply_move(sequence: List[str], move: Dict[str, Any]) -> List[str]:
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
        start, end, to = int(move["start"]), int(move["end"]), int(move["to"])
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
            start1, end1, start2, end2 = start2, end2, start1, end1

        block1 = seq[start1:end1 + 1]
        block2 = seq[start2:end2 + 1]

        new_seq = seq.copy()
        new_seq[start1:end1 + 1] = block2
        new_seq[start2:end2 + 1] = block1
        return new_seq

    return seq


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
    ps.ExecuteSimTalk(".Models.Model.EventController.reset")
    ps.ExecuteSimTalk(".Models.Model.EventController.start")

    start_wall = time.perf_counter()
    last_sim_time = None
    stable_count = 0

    while True:
        if time.perf_counter() - start_wall > timeout_sec:
            raise TimeoutError(f"Simulation did not stabilize within {timeout_sec} seconds.")

        sim_time = float(ps.GetValue(".Models.Model.EventController.SimTime"))

        if last_sim_time is not None and abs(sim_time - last_sim_time) < 1e-9:
            stable_count += 1
        else:
            stable_count = 0

        if stable_count >= stable_polls_required:
            return sim_time

        last_sim_time = sim_time
        time.sleep(poll_interval_sec)

# ============================================================
# LOGGING
# ============================================================

FIELDNAMES = [
    "iteration",
    "restart_id",
    "iteration_in_restart",
    "candidate_in_iteration",
    "search_mode",
    "no_improvement_count_before_iteration",
    "parent_sequence_signature",
    "action",
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
    "valid",
    "duplicate_for_same_parent_sequence",
    "recently_rejected_for_same_parent_sequence",
    "diversification_compatible",
    "parse_error",
    "validation_message",
    "resulting_makespan",
    "best_makespan_before",
    "best_makespan_after",
    "global_best_makespan_after",
    "completed_simulations",
    "actual_global_best_makespan",
    "elapsed_optimization_time_sec",
    "delta_vs_best_before",
    "status",
    "llm_time_sec",
    "simulation_time_sec",
    "exited_parts",
    "completion_valid",
    "llm_attempts",
    "retry_reason"
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
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES, delimiter=CSV_DELIMITER)
        writer.writeheader()


def append_log_row(row: Dict[str, Any]) -> None:
    cleaned_row = {k: clean_csv_text(v) for k, v in row.items()}
    log_rows.append(cleaned_row)

    if not WRITE_LOGS:
        return

    with LOG_FILE.open("a", newline="", encoding=CSV_ENCODING) as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES, delimiter=CSV_DELIMITER)
        writer.writerow(cleaned_row)


def append_raw_llm_log(entry: Dict[str, Any]) -> None:
    if not WRITE_LOGS:
        return

    with RAW_LLM_LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def write_global_best_sequence_file(sequence: List[str], makespan: float) -> None:
    """
    Writes the exact 64-position part sequence belonging to the global best
    solution into a separate text file for manual verification.
    """
    if not WRITE_LOGS:
        return

    with GLOBAL_BEST_SEQUENCE_FILE.open("w", encoding="utf-8") as f:
        f.write(f"Global best makespan: {makespan}\n")
        f.write(f"Sequence length: {len(sequence)}\n\n")
        f.write("Position;Part\n")
        for idx, part in enumerate(sequence, start=1):
            f.write(f"{idx};{part}\n")

# ============================================================
# MAIN
# ============================================================

initialize_log_file()
print(f"Main CSV log file will be written to: {LOG_FILE.resolve()}")
print(f"Raw LLM reply log file will be written to: {RAW_LLM_LOG_FILE.resolve()}")

try:
    ps = win32com.client.Dispatch("Tecnomatix.PlantSimulation.RemoteControl")
    print("Connected via Tecnomatix.PlantSimulationRemoteControl")
except Exception as e:
    raise RuntimeError(f"Could not connect to Plant Simulation RemoteControl: {e}")

ps.LoadModel(MODEL_PATH)
print("Model opened successfully")

part_to_route, station_to_processes = read_flat_process_table(ps)
print(f"Loaded flat process table: {len(part_to_route)} part types, {len(station_to_processes)} stations.")

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

current_makespan = run_simulation_until_finished(ps)
initial_exited_parts = read_exited_parts(ps)
if not is_completion_valid(initial_exited_parts):
    raise RuntimeError(
        f"Initial simulation did not complete all parts: "
        f"ExitedParts={initial_exited_parts}, expected {EXPECTED_EXITED_PARTS}"
    )
initial_makespan = current_makespan
print("Initial makespan:", current_makespan)
print("Initial ExitedParts:", initial_exited_parts)

resource_stats = read_resource_statistics(ps, num_stations=NUM_STATIONS)
resource_summary_text = build_resource_summary(resource_stats, top_k=RESOURCE_TOP_K)
bottleneck_process_summary_text = build_bottleneck_process_summary(
    resource_stats,
    station_to_processes,
    top_k=BOTTLENECK_PROCESS_TOP_K
)
print(resource_summary_text)
print(bottleneck_process_summary_text)

accepted_resource_stats_history: List[List[Dict[str, Any]]] = [resource_stats]
utilization_trend_text = build_utilization_trend_summary(None, resource_stats, top_k=UTILIZATION_TREND_TOP_K)
print(utilization_trend_text)

original_sequence = current_sequence.copy()

# Best solution inside the current restart trajectory
best_sequence = current_sequence.copy()
best_makespan = current_makespan

# Best solution across all restart trajectories
global_best_sequence = current_sequence.copy()
global_best_makespan = current_makespan

# Counts actual Plant Simulation executions that reached the end of run_simulation_until_finished().
# This includes the initial baseline simulation, candidate simulations, and restart baseline simulations.
completed_simulations = 1

restart_id = 1
iteration_in_restart = 0

seen_moves_by_parent = set()
recent_rejected_keys = set()
recent_rejected_entries: List[Dict[str, str]] = []

action_counts = {
    "swap": 0,
    "move": 0,
    "reverse": 0,
    "reinsert_block": 0,
    "swap_blocks": 0,
}
no_improvement_count = 0

run_start = time.perf_counter()

for iteration in range(1, NUM_ITERATIONS + 1):
    search_mode = "diversification" if no_improvement_count >= STAGNATION_THRESHOLD else "normal"

    candidates_this_iteration = (
        DIVERSIFICATION_CANDIDATES_PER_ITERATION
        if search_mode == "diversification"
        else BASE_CANDIDATES_PER_ITERATION
    )

    iteration_in_restart += 1

    print(f"\n=== Iteration {iteration} | restart: {restart_id} | restart iteration: {iteration_in_restart} | mode: {search_mode} | stagnation: {no_improvement_count} ===")

    parent_sequence_signature = get_sequence_signature(current_sequence)
    recent_history_text = summarize_recent_history(log_rows, max_items=5)
    rejected_moves_text = summarize_recent_rejected_moves(recent_rejected_entries, max_items=12)
    recent_move_text_block = build_recent_move_text_block(log_rows, max_items=RECENT_MOVE_TEXT_LIMIT)
    action_counts_text = ", ".join(f"{k}: {v}" for k, v in action_counts.items())

    resource_stats = read_resource_statistics(ps, num_stations=NUM_STATIONS)
    resource_summary_text = build_resource_summary(resource_stats, top_k=RESOURCE_TOP_K)
    bottleneck_process_summary_text = build_bottleneck_process_summary(
        resource_stats,
        station_to_processes,
        top_k=BOTTLENECK_PROCESS_TOP_K
    )

    previous_accepted_stats = accepted_resource_stats_history[-2] if len(accepted_resource_stats_history) >= 2 else None
    current_accepted_stats = accepted_resource_stats_history[-1] if accepted_resource_stats_history else resource_stats
    utilization_trend_text = build_utilization_trend_summary(
        previous_accepted_stats,
        current_accepted_stats,
        top_k=UTILIZATION_TREND_TOP_K
    )

    best_candidate_data = None

    for candidate_idx in range(1, candidates_this_iteration + 1):
        best_before = best_makespan

        prompt = build_prompt(
            sequence=current_sequence,
            current_makespan=current_makespan,
            best_makespan=best_makespan,
            recent_history_text=recent_history_text,
            rejected_moves_text=rejected_moves_text,
            action_counts_text=action_counts_text,
            no_improvement_count=no_improvement_count,
            search_mode=search_mode,
            candidate_idx=candidate_idx,
            resource_summary_text=resource_summary_text,
            utilization_trend_text=utilization_trend_text,
            bottleneck_process_summary_text=bottleneck_process_summary_text,
            recent_move_text_block=recent_move_text_block
        )

        llm_start = time.perf_counter()
        try:
            (
                reply,
                move,
                parse_error,
                validation_message,
                valid,
                llm_attempts,
                retry_reason,
                duplicate_for_same_parent,
                rejected_for_same_parent,
            ) = generate_candidate_with_retries(
                base_prompt=prompt,
                parent_sequence=current_sequence,
                seen_moves_by_parent=seen_moves_by_parent,
                recent_rejected_keys=recent_rejected_keys,
                max_retries=MAX_LLM_RETRIES_PER_CANDIDATE,
            )
        except Exception as e:
            reply = ""
            move = None
            parse_error = f"LLM call failed: {e}"
            validation_message = ""
            valid = False
            llm_attempts = 1
            retry_reason = "llm_call_failed"
            duplicate_for_same_parent = False
            rejected_for_same_parent = False
        llm_time_sec = time.perf_counter() - llm_start

        candidate_makespan = None
        candidate_exited_parts = None
        candidate_completion_valid = None
        simulation_time_sec = None
        status = "Skipped"
        diversification_compatible = None

        if valid:
            diversification_compatible = is_diversification_move(move)
            parent_aware_key = build_parent_aware_key(current_sequence, move)

            candidate_sequence = apply_move(current_sequence, move)

            sim_start = time.perf_counter()
            try:
                set_sequence(ps, candidate_sequence)
                candidate_makespan = run_simulation_until_finished(ps)
                completed_simulations += 1
                candidate_exited_parts = read_exited_parts(ps)
                candidate_completion_valid = is_completion_valid(candidate_exited_parts)
                simulation_time_sec = time.perf_counter() - sim_start

                if not candidate_completion_valid:
                    status = "SimulationIncomplete"
                    parse_error = (
                        f"Simulation completed with ExitedParts={candidate_exited_parts}, "
                        f"expected {EXPECTED_EXITED_PARTS}. Candidate ignored."
                    )
                    raise ValueError(parse_error)

                would_be_duplicate = parent_aware_key in seen_moves_by_parent
                would_be_recently_rejected = parent_aware_key in recent_rejected_keys
                would_fail_diversification = False
                diversification_soft_warning = (
                    search_mode == "diversification"
                    and not diversification_compatible
                )

                duplicate_for_same_parent = would_be_duplicate
                rejected_for_same_parent = would_be_recently_rejected

                if candidate_makespan < best_makespan:
                    seen_moves_by_parent.add(parent_aware_key)
                    action_counts[move["action"]] += 1

                    if best_candidate_data is None or candidate_makespan < best_candidate_data["candidate_makespan"]:
                        best_candidate_data = {
                            "move": move,
                            "parent_aware_key": parent_aware_key,
                            "candidate_sequence": candidate_sequence,
                            "candidate_makespan": candidate_makespan,
                            "parent_sequence_signature": parent_sequence_signature
                        }

                    if diversification_soft_warning:
                        status = "EvaluatedImprovementDespiteDiversificationMismatch"
                    elif would_be_duplicate:
                        status = "EvaluatedImprovementOverrideDuplicate"
                    elif would_be_recently_rejected:
                        status = "EvaluatedImprovementOverrideRecentRejection"
                    else:
                        if diversification_soft_warning:
                            status = "EvaluatedDiversificationMismatch"
                        else:
                            status = "Evaluated"
                else:
                    if would_be_duplicate:
                        status = "DuplicateForSameParentSequence"
                    elif would_be_recently_rejected:
                        status = "RecentlyRejectedForSameParentSequence"
                    else:
                        seen_moves_by_parent.add(parent_aware_key)
                        action_counts[move["action"]] += 1

                        if best_candidate_data is None or candidate_makespan < best_candidate_data["candidate_makespan"]:
                            best_candidate_data = {
                                "move": move,
                                "parent_aware_key": parent_aware_key,
                                "candidate_sequence": candidate_sequence,
                                "candidate_makespan": candidate_makespan,
                                "parent_sequence_signature": parent_sequence_signature
                            }

                        status = "Evaluated"

            except Exception as e:
                simulation_time_sec = time.perf_counter() - sim_start
                if status == "SimulationIncomplete":
                    parse_error = str(e)
                else:
                    parse_error = f"Simulation failed: {e}"
                    status = "SimulationFailed"
        else:
            if parse_error is None:
                if duplicate_for_same_parent:
                    status = "DuplicateForSameParentSequenceAfterRetries"
                elif rejected_for_same_parent:
                    status = "RecentlyRejectedForSameParentSequenceAfterRetries"
                else:
                    status = "Invalid"

        delta_vs_best_before = (
            candidate_makespan - best_before
            if isinstance(candidate_makespan, (int, float))
            else None
        )

        append_log_row({
            "iteration": iteration,
            "restart_id": restart_id,
            "iteration_in_restart": iteration_in_restart,
            "candidate_in_iteration": candidate_idx,
            "search_mode": search_mode,
            "no_improvement_count_before_iteration": no_improvement_count,
            "parent_sequence_signature": parent_sequence_signature[:12],
            "action": move.get("action") if isinstance(move, dict) else None,
            "swap_i": move.get("i") if isinstance(move, dict) and move.get("action") == "swap" else None,
            "swap_j": move.get("j") if isinstance(move, dict) and move.get("action") == "swap" else None,
            "move_from": move.get("from") if isinstance(move, dict) and move.get("action") == "move" else None,
            "move_to": move.get("to") if isinstance(move, dict) and move.get("action") == "move" else None,
            "reverse_start": move.get("start") if isinstance(move, dict) and move.get("action") == "reverse" else None,
            "reverse_end": move.get("end") if isinstance(move, dict) and move.get("action") == "reverse" else None,
            "reinsert_block_start": move.get("start") if isinstance(move, dict) and move.get("action") == "reinsert_block" else None,
            "reinsert_block_end": move.get("end") if isinstance(move, dict) and move.get("action") == "reinsert_block" else None,
            "reinsert_block_to": move.get("to") if isinstance(move, dict) and move.get("action") == "reinsert_block" else None,
            "swap_blocks_start1": move.get("start1") if isinstance(move, dict) and move.get("action") == "swap_blocks" else None,
            "swap_blocks_end1": move.get("end1") if isinstance(move, dict) and move.get("action") == "swap_blocks" else None,
            "swap_blocks_start2": move.get("start2") if isinstance(move, dict) and move.get("action") == "swap_blocks" else None,
            "swap_blocks_end2": move.get("end2") if isinstance(move, dict) and move.get("action") == "swap_blocks" else None,
            "valid": valid,
            "duplicate_for_same_parent_sequence": duplicate_for_same_parent,
            "recently_rejected_for_same_parent_sequence": rejected_for_same_parent,
            "diversification_compatible": diversification_compatible,
            "parse_error": parse_error,
            "validation_message": validation_message,
            "resulting_makespan": candidate_makespan,
            "best_makespan_before": best_before,
            "best_makespan_after": best_makespan,
            "global_best_makespan_after": global_best_makespan,
            "completed_simulations": completed_simulations,
            "actual_global_best_makespan": global_best_makespan,
            "elapsed_optimization_time_sec": round(time.perf_counter() - run_start, 4),
            "delta_vs_best_before": delta_vs_best_before,
            "status": status,
            "llm_time_sec": round(llm_time_sec, 4),
            "simulation_time_sec": round(simulation_time_sec, 4) if simulation_time_sec is not None else None,
            "exited_parts": candidate_exited_parts,
            "completion_valid": candidate_completion_valid,
            "llm_attempts": llm_attempts,
            "retry_reason": retry_reason,
        })

        append_raw_llm_log({
            "iteration": iteration,
            "candidate_in_iteration": candidate_idx,
            "search_mode": search_mode,
            "parent_sequence_signature": parent_sequence_signature[:12],
            "resource_summary_text": resource_summary_text,
            "utilization_trend_text": utilization_trend_text,
            "bottleneck_process_summary_text": bottleneck_process_summary_text,
            "recent_move_text_block": recent_move_text_block,
            "parse_error": parse_error,
            "llm_attempts": llm_attempts,
            "retry_reason": retry_reason,
            "raw_llm_reply": reply
        })

    if best_candidate_data and best_candidate_data["candidate_makespan"] < best_makespan:
        best_sequence = best_candidate_data["candidate_sequence"].copy()
        best_makespan = best_candidate_data["candidate_makespan"]
        current_sequence = best_sequence.copy()
        current_makespan = best_makespan
        no_improvement_count = 0
        print("Improved in current restart:", best_makespan)

        if best_makespan < global_best_makespan:
            global_best_sequence = best_sequence.copy()
            global_best_makespan = best_makespan
            print("New global best:", global_best_makespan)

        resource_stats = read_resource_statistics(ps, num_stations=NUM_STATIONS)
        accepted_resource_stats_history.append(resource_stats)
        if len(accepted_resource_stats_history) > ACCEPTED_RESOURCE_HISTORY_LIMIT:
            accepted_resource_stats_history.pop(0)

        resource_summary_text = build_resource_summary(resource_stats, top_k=RESOURCE_TOP_K)
        bottleneck_process_summary_text = build_bottleneck_process_summary(
            resource_stats,
            station_to_processes,
            top_k=BOTTLENECK_PROCESS_TOP_K
        )
        previous_accepted_stats = accepted_resource_stats_history[-2] if len(accepted_resource_stats_history) >= 2 else None
        utilization_trend_text = build_utilization_trend_summary(
            previous_accepted_stats,
            resource_stats,
            top_k=UTILIZATION_TREND_TOP_K
        )
        print(resource_summary_text)
        print(utilization_trend_text)
        print(bottleneck_process_summary_text)
    else:
        if best_candidate_data is not None:
            rejected_key = best_candidate_data["parent_aware_key"]
            recent_rejected_keys.add(rejected_key)
            recent_rejected_entries.append({
                "key": rejected_key,
                "move_text": format_move(best_candidate_data["move"]),
                "parent_signature_short": best_candidate_data["parent_sequence_signature"][:12]
            })

            if len(recent_rejected_entries) > RECENT_REJECTED_LIMIT:
                oldest = recent_rejected_entries.pop(0)
                recent_rejected_keys.discard(oldest["key"])

        current_sequence = best_sequence.copy()
        current_makespan = best_makespan
        no_improvement_count += 1
        print("No improvement this iteration. Best in current restart remains:", best_makespan)

        if no_improvement_count >= RESTART_STAGNATION_THRESHOLD and iteration < NUM_ITERATIONS:
            print(f"Restart triggered after {no_improvement_count} stagnant iterations. Restart best: {best_makespan}; global best: {global_best_makespan}")

            restart_id += 1
            iteration_in_restart = 0

            current_sequence = original_sequence.copy()
            set_sequence(ps, current_sequence)
            current_makespan = run_simulation_until_finished(ps)
            completed_simulations += 1
            restart_exited_parts = read_exited_parts(ps)
            if not is_completion_valid(restart_exited_parts):
                raise RuntimeError(
                    f"Restart baseline simulation did not complete all parts: "
                    f"ExitedParts={restart_exited_parts}, expected {EXPECTED_EXITED_PARTS}"
                )

            best_sequence = current_sequence.copy()
            best_makespan = current_makespan
            no_improvement_count = 0

            seen_moves_by_parent = set()
            recent_rejected_keys = set()
            recent_rejected_entries = []
            action_counts = {
                "swap": 0,
                "move": 0,
                "reverse": 0,
                "reinsert_block": 0,
                "swap_blocks": 0,
            }

            resource_stats = read_resource_statistics(ps, num_stations=NUM_STATIONS)
            accepted_resource_stats_history = [resource_stats]
            resource_summary_text = build_resource_summary(resource_stats, top_k=RESOURCE_TOP_K)
            bottleneck_process_summary_text = build_bottleneck_process_summary(
                resource_stats,
                station_to_processes,
                top_k=BOTTLENECK_PROCESS_TOP_K
            )
            utilization_trend_text = build_utilization_trend_summary(None, resource_stats, top_k=UTILIZATION_TREND_TOP_K)
            print(f"Restart {restart_id} initialized. Initial makespan: {current_makespan}")

    # End-of-iteration summary row for easy plotting of convergence curves.
    # Use rows with status == "IterationSummary" to plot:
    # - completed_simulations vs actual_global_best_makespan
    # - elapsed_optimization_time_sec vs actual_global_best_makespan
    append_log_row({
        "iteration": iteration,
        "restart_id": restart_id,
        "iteration_in_restart": iteration_in_restart,
        "search_mode": search_mode,
        "no_improvement_count_before_iteration": no_improvement_count,
        "parent_sequence_signature": None,
        "action": None,
        "swap_i": None,
        "swap_j": None,
        "move_from": None,
        "move_to": None,
        "reverse_start": None,
        "reverse_end": None,
        "reinsert_block_start": None,
        "reinsert_block_end": None,
        "reinsert_block_to": None,
        "swap_blocks_start1": None,
        "swap_blocks_end1": None,
        "swap_blocks_start2": None,
        "swap_blocks_end2": None,
        "valid": None,
        "duplicate_for_same_parent_sequence": None,
        "recently_rejected_for_same_parent_sequence": None,
        "diversification_compatible": None,
        "parse_error": None,
        "validation_message": "End-of-iteration summary row",
        "resulting_makespan": None,
        "best_makespan_before": None,
        "best_makespan_after": best_makespan,
        "global_best_makespan_after": global_best_makespan,
        "completed_simulations": completed_simulations,
        "actual_global_best_makespan": global_best_makespan,
        "elapsed_optimization_time_sec": round(time.perf_counter() - run_start, 4),
        "delta_vs_best_before": None,
        "status": "IterationSummary",
        "llm_time_sec": None,
        "simulation_time_sec": None,
        "exited_parts": None,
        "completion_valid": None,
        "llm_attempts": None,
        "retry_reason": None,
    })

print("\nFinal best result across all restarts:", global_best_makespan)
print("Final best result in last restart:", best_makespan)
total_elapsed = time.perf_counter() - run_start
print(f"Total optimization time: {total_elapsed:.1f} s")
print(f"Final main CSV log file: {LOG_FILE.resolve()}")
print(f"Final raw LLM reply log file: {RAW_LLM_LOG_FILE.resolve()}")

write_global_best_sequence_file(global_best_sequence, global_best_makespan)

append_log_row({
    "iteration": NUM_ITERATIONS,
    "restart_id": restart_id,
    "iteration_in_restart": iteration_in_restart,
    "candidate_in_iteration": None,
    "search_mode": None,
    "no_improvement_count_before_iteration": no_improvement_count,
    "parent_sequence_signature": None,
    "action": None,
    "swap_i": None,
    "swap_j": None,
    "move_from": None,
    "move_to": None,
    "reverse_start": None,
    "reverse_end": None,
    "reinsert_block_start": None,
    "reinsert_block_end": None,
    "reinsert_block_to": None,
    "swap_blocks_start1": None,
    "swap_blocks_end1": None,
    "swap_blocks_start2": None,
    "swap_blocks_end2": None,
    "valid": None,
    "duplicate_for_same_parent_sequence": None,
    "recently_rejected_for_same_parent_sequence": None,
    "diversification_compatible": None,
    "parse_error": None,
    "validation_message": " | ".join(global_best_sequence),
    "resulting_makespan": None,
    "best_makespan_before": None,
    "best_makespan_after": best_makespan,
    "global_best_makespan_after": global_best_makespan,
    "completed_simulations": completed_simulations,
    "actual_global_best_makespan": global_best_makespan,
    "elapsed_optimization_time_sec": round(time.perf_counter() - run_start, 4),
    "delta_vs_best_before": None,
    "status": "FinalGlobalBestSequence",
    "llm_time_sec": None,
    "simulation_time_sec": None,
    "exited_parts": None,
    "completion_valid": None,
    "llm_attempts": None,
    "retry_reason": None,
})

print(f"Global best sequence file: {GLOBAL_BEST_SEQUENCE_FILE.resolve()}")
