# Resume Reviewer

[![Latest release](https://img.shields.io/github/v/release/ckovacs_boxboat/screener)](https://github.com/ckovacs_boxboat/screener/releases/latest)

AI workflow that evaluates a directory of resumes against a skill list and a
project description using the Claude API, and exports the results as an Excel
spreadsheet.

## Quickstart

```sh
# 1. Install uv (skip if you have it)
curl -LsSf https://astral.sh/uv/install.sh | sh        # macOS/Linux
# or: brew install uv

# 2. Install the latest revision of resume-reviewer as a tool
uv tool install --force git+ssh://git@github.com/ckovacs_boxboat/screener.git

# 3. Configure the backend (ICA is the default)
export ICA_ANTHROPIC_API_KEY=<your-key>

# 4. Put your inputs in the current directory and run
#      skills.txt   one skill per line          (default skills list)
#      project.md   the project description     (default project description)
resume-reviewer ./resumes --output evaluations.xlsx
```

Re-run step 2 any time to pick up the newest revision. To work from a clone
instead, use `uv sync` and `uv run resume-reviewer ...` as described below.

## What it does

For each resume (`.pdf`, `.docx`, `.txt`, `.md`) in the input directory, Claude
produces a structured evaluation: candidate name, an overall fit score (1-5), a
hire recommendation, per-skill scores with supporting evidence, strengths, and
gaps. Results land in an `.xlsx` workbook:

- **Summary** — one row per candidate, sorted by overall score, with one column
  per skill and color-coded recommendations
- **Skill Details** — per-skill evidence for every candidate
- **Errors** — any resumes that failed to process (only present if something failed)

PDFs are sent to the API natively (no local text extraction); DOCX/TXT/MD are
read as text. Evaluations run concurrently, and the shared evaluation context
(prompt + project description + skills) is prompt-cached across calls.

## Setup

Managed with [uv](https://docs.astral.sh/uv/):

```sh
uv sync          # creates .venv and installs locked dependencies
```

## Backends

**IBM Consulting Advantage (`ica`) is the default backend.** ICA exposes an
Anthropic-native Messages endpoint (`/ica/v1/messages`) fronting
Anthropic-on-Bedrock, so the stock `anthropic` SDK works against it unchanged
with a base URL override — the same approach as
`uft-to-flaui-agentic-framework`. Structured results use tool-forced JSON
(a forced `tool_choice` on a schema-carrying tool), which is verified to work
through ICA.

| Provider | Select with | Auth | Model env | Host env |
|---|---|---|---|---|
| `ica` (default) | `--provider ica` or `PROVIDER=ica` | `ICA_ANTHROPIC_API_KEY` (or `ICA_API_KEY`) | `ICA_MODEL` | `ICA_HOST` (the `/ica` root, not `/ica/v1`) |
| `anthropic` | `--provider anthropic` | `ANTHROPIC_API_KEY` | `CLAUDE_MODEL` | — |

```sh
export ICA_ANTHROPIC_API_KEY=<your-key>       # default (ICA)
# or:
uv run resume-reviewer ... --provider anthropic   # direct Claude API
```

## Usage

```sh
export ICA_ANTHROPIC_API_KEY=<your-key>
uv run resume-reviewer ./resumes \
  --skills "Python, Kubernetes, Terraform, CI/CD" \
  --project-file project.md \
  --prompt "Weight hands-on infrastructure experience heavily." \
  --output evaluations.xlsx
```

| Flag | Description |
|---|---|
| `resumes_dir` | Directory containing resumes (`.pdf`, `.docx`, `.txt`, `.md`) |
| `--skills` / `--skills-file` | Comma-separated list, or a file with one skill per line (default: `skills.txt` in the current directory) |
| `--project` / `--project-file` | Project description text or file (default: `project.md` in the current directory) |
| `--prompt` / `--prompt-file` | Additional evaluation instructions (optional) |
| `--output` | Output `.xlsx` path (default: `evaluations.xlsx`) |
| `--provider` | `ica` (default) or `anthropic` |
| `--host` | Base URL override for the ICA backend |
| `--model` | Claude model for full evaluations (default: `ICA_MODEL`/`CLAUDE_MODEL` env, else `claude-opus-4-8`) |
| `--workers` | Concurrent evaluations (default: 4) |
| `--screen` | Enable the two-tier low-cost screening mode (see below) |
| `--screen-model` | Screening model (default: `claude-haiku-4-5`) |
| `--screen-cutoff` | Minimum screening score to advance to full evaluation (default: 3) |

## Cost control: two-tier screening

With `--screen`, every resume first gets a cheap screening pass with a
lower-cost model (Haiku 4.5 by default — ~5x cheaper per token than Opus). Only
candidates scoring at or above `--screen-cutoff` advance to the full per-skill
evaluation with the main model. Screened-out candidates still appear on the
Summary sheet (marked `screened out` in the Stage column) with their screening
score and summary, but without per-skill detail.

```sh
uv run resume-reviewer ./resumes \
  --skills-file skills.txt --project-file project.md \
  --screen --screen-cutoff 3
```

Use this for large batches; skip it when you have a handful of resumes and want
full detail on everyone.
