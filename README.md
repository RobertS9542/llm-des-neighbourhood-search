README

Repository contents

This repository contains the materials associated with the study
“Context-Aware LLM-Guided Neighbourhood Search for Simulation-Based
Production Scheduling”.

The repository includes: - the Python implementation of the LLM-guided
optimiser; - the Python implementation of the simulated annealing (SA)
benchmark; - the raw log files generated during the optimisation
experiments; - the spreadsheet files used for the evaluation and
comparison of the LLM-guided, genetic algorithm (GA), and SA
experiments; - the spreadsheet containing the timing measurements and
sensitivity/break-even analysis; and - this README file.

The Siemens Tecnomatix Plant Simulation model is not included in the
public repository due to licensing considerations associated with the
commercial simulation software. Consequently, the optimisation scripts
cannot reproduce the reported experiments without access to the
corresponding Plant Simulation model and a compatible installation of
Siemens Tecnomatix Plant Simulation.

LLM-guided optimiser

To use the LLM-guided optimiser, Ollama must be installed and the
required model must be pulled. The runtime model used in the study was
Mistral Small 3.2 24B through Ollama.

The config section of the optimiser contains the following adjustable
parameters. MODEL_PATH, together with the Plant Simulation object paths,
must be defined according to the location and internal structure of the
simulation model being used.

MODEL_PATH: Defines the access path to the Plant Simulation model to be
optimised.

LLM_MODEL: Defines the exact name of the LLM to be used in Ollama (it
must exactly match the model name in Ollama).

LLM_TEMPERATURE: Defines the temperature setting used by the LLM.

SEQUENCE_LENGTH: Defines the length of the production sequence.

EXPECTED_EXITED_PARTS: Defines the expected number of completed parts
used to verify that a simulation run finished successfully.

NUM_ITERATIONS: Defines the number of optimisation iterations.

NUM_STATIONS: Defines the number of workstations in the simulation
model.

BASE_CANDIDATES_PER_ITERATION: Defines the number of candidate
production sequences requested from the LLM during normal search mode.

STAGNATION_THRESHOLD: Defines the number of consecutive iterations
without improvement required to activate the diversification search
mode.

RESTART_STAGNATION_THRESHOLD: Defines the number of consecutive
iterations without improvement required to restart the optimisation
process from the initial production sequence.

CSV_DELIMITER: Defines the delimiter used in the generated CSV log
files.

CSV_ENCODING: Defines the character encoding used when writing the CSV
log files.

WRITE_LOGS: Enables or disables the generation of optimisation log
files.

POLL_INTERVAL_SEC: Defines the time interval between consecutive checks
of the simulation completion status.

STABLE_POLLS_REQUIRED: Defines the number of consecutive successful
completion checks required before a simulation run is considered
finished.

SIM_TIMEOUT_SEC: Defines the maximum allowed duration of a single
simulation run before a timeout occurs.

RECENT_REJECTED_LIMIT: Defines the maximum number of recently rejected
candidate sequences stored by the optimiser.

RECENT_MOVE_TEXT_LIMIT: Defines the maximum number of recently applied
move descriptions stored for prompt generation.

MAX_LLM_RETRIES_PER_CANDIDATE: Defines the maximum number of retries
allowed for generating a valid candidate from the LLM.

MIN_REINSERT_BLOCK_LEN: Defines the minimum block length used in
reinsert block moves.

MAX_REINSERT_BLOCK_LEN: Defines the maximum block length used in
reinsert block moves.

MIN_SWAP_BLOCK_LEN: Defines the minimum block length used in swap block
moves.

MAX_SWAP_BLOCK_LEN: Defines the maximum block length used in swap block
moves.

EXITED_PARTS_PATH: Defines the access path of the exited-parts counter
in the Plant Simulation model.

RESOURCE_STATS_TABLE_PATH: Defines the access path of the resource
statistics table in the Plant Simulation model.

RESOURCE_WORKING_ROW: Defines the row containing the working-time
statistics in the resource statistics table.

RESOURCE_SETTINGUP_ROW: Defines the row containing the setup-time
statistics in the resource statistics table.

RESOURCE_WAITING_ROW: Defines the row containing the waiting-time
statistics in the resource statistics table.

RESOURCE_BLOCKED_ROW: Defines the row containing the blocked-time
statistics in the resource statistics table.

RESOURCE_TOP_K: Defines the number of resources with the highest
utilisation included in the extracted statistics.

UTILIZATION_TREND_TOP_K: Defines the number of resource utilisation
trends included in the extracted statistics.

ACCEPTED_RESOURCE_HISTORY_LIMIT: Defines the number of previously
accepted resource statistic snapshots retained by the optimiser.

FLAT_PROCESS_TABLE_PATH: Defines the access path of the flat
process-route table in the Plant Simulation model.

FLAT_PROCESS_MAX_ROWS: Defines the maximum number of rows read from the
flat process-route table.

BOTTLENECK_PROCESS_TOP_K: Defines the number of bottleneck processes
included in the extracted statistics.

Simulated annealing benchmark

The repository also contains the Python implementation of the simulated
annealing benchmark introduced in the revised study. The SA optimiser
uses the same Plant Simulation communication mechanism and the same
principal neighbourhood-move types as the LLM-guided optimiser, but
candidate moves are selected without an LLM or the contextual
information supplied to the LLM-guided method.

The SA experiments used the following main settings: - candidate
evaluations per run: 3000; - initial temperature: 100; - cooling factor:
0.999; - minimum temperature: 0.01; - neighbourhood-selection weights:
move = 1.0, swap = 1.0, reverse = 1.0, reinsert_block = 1.0, and
swap_blocks = 0.5.

The parameter values were selected to provide gradual cooling and a
sufficiently extensive exploration of the solution space within the
predefined evaluation budget. The lower weight assigned to swap_blocks
reduces the frequency of this comparatively disruptive move while
retaining it as part of the available neighbourhood.

The implementation follows the standard SA acceptance principle:
improving candidates are accepted, while worsening candidates can be
accepted according to the temperature-dependent acceptance probability.
Ten independent SA runs were conducted for each of the three worker
configurations investigated in the study.

Timing-data processing note

The original CSV log files are retained unchanged and should be regarded
as the authoritative source for the recorded optimisation and timing
data.

During spreadsheet-based processing of the logged timing data, isolated
formatting artefacts were introduced because spreadsheet software
automatically interpreted some decimal numeric values as dates under the
applicable locale settings. For example, a raw timing value written as
2.23 seconds could be displayed or imported as a date-like value such as
“23-Feb”. These artefacts were introduced only during spreadsheet
interpretation/processing and were not present in the original CSV logs.

Where such cases occurred, the affected values in the derived analysis
spreadsheets were corrected using the corresponding numeric values from
the original CSV files. The corrections therefore concern spreadsheet
formatting/type interpretation only; they do not represent changes to
the original experimental logs and do not affect the reported
optimisation results or the underlying timing measurements. The original
CSV files have been preserved unchanged to allow the processed values to
be verified against the source data.

Reproducibility note

The simulation model used in the study is deterministic. Nevertheless,
reproduction of the complete experiments requires the corresponding
Plant Simulation model, the same or a compatible Plant Simulation
environment, and the required Python dependencies. The LLM-guided
experiments additionally require Ollama and the specified runtime LLM.
Exact LLM outputs may also depend on the runtime software/model
environment.# llm-des-neighbourhood-search
