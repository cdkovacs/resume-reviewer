# Resume Reviewer

[![Latest release](https://img.shields.io/github/v/release/cdkovacs/resume-reviewer)](https://github.com/cdkovacs/resume-reviewer/releases/latest)

AI workflow that evaluates a directory of resumes against a skill list and a
project description using the ICA/Claude API, and exports the results as an Excel
spreadsheet.

## Quickstart

**macOS / Linux:**

```sh
# 1. Install uv (skip if you have it)
curl -LsSf https://astral.sh/uv/install.sh | sh        # or: brew install uv

# 2. Install the latest revision of resume-reviewer as a tool
uv tool install --force git+https://github.com/cdkovacs/resume-reviewer.git

# 3. Configure the backend (ICA is the default)
export ICA_ANTHROPIC_API_KEY=<your-key>

# 4. Put your inputs in the current directory and run
#      skills.csv   skill[,weight] per line     (default skills list; weight defaults to 1.0)
#      project.md   the project description     (default project description)
resume-reviewer ./resumes --output evaluations.xlsx
```

**Windows (PowerShell):**

```powershell
# 1. Install uv (skip if you have it)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
# or: winget install --id=astral-sh.uv -e

# 2. Install the latest revision of resume-reviewer as a tool
uv tool install --force git+https://github.com/cdkovacs/resume-reviewer.git

# 3. Configure the backend (ICA is the default)
$env:ICA_ANTHROPIC_API_KEY = "<your-key>"

# 4. Put your inputs in the current directory and run
#      skills.csv   skill[,weight] per line     (default skills list; weight defaults to 1.0)
#      project.md   the project description     (default project description)
resume-reviewer .\resumes --output evaluations.xlsx
```

Re-run step 2 any time to pick up the newest revision. To work from a clone
instead, use `uv sync` and `uv run resume-reviewer ...` as described below.

### Install from a wheel

Every release has the built wheel attached as an asset (immutable, with a
signed attestation). Useful when you want a pinned version, or on machines
without git/SSH access to the repo:

```sh
# Download the wheel from the latest release (or pass a tag instead of omitting it)
gh release download --repo cdkovacs/resume-reviewer --pattern '*.whl' --dir dist

# Install it as a tool (or: uv pip install ./dist/*.whl into an active venv)
uv tool install ./dist/resume_reviewer-*.whl

# Optional: verify the asset came from an immutable release
gh release verify-asset --repo cdkovacs/resume-reviewer <tag> ./dist/resume_reviewer-*.whl
```

Alternatively download the `.whl` from the Releases page in a browser and run
`uv tool install ./resume_reviewer-<version>-py3-none-any.whl`.

## What it does

For each resume (`.pdf`, `.docx`, `.txt`, `.md`) in the input directory, ICA/Claude
produces a structured evaluation: candidate name, an overall fit score (0-4), a
hire recommendation, per-skill scores with supporting evidence, strengths, and
gaps. Results land in an `.xlsx` workbook:

- **Summary** — one row per candidate, sorted by overall score, with one column
  per skill and color-coded recommendations
- **Skill Details** — per-skill evidence for every candidate
- **Ranking** — force-ranked list and horizontal bar chart of each candidate's
  weighted average (0-4), sorted most qualified first
- **Candidate Feedback** — one row per fully-evaluated candidate (name, email
  from the resume, and two feedback texts: a professional voice, plus a warmer
  personal voice for close, well-known colleagues — pick the column that fits
  how well you know the candidate): a respectful, positive message
  telling the candidate they are being considered for a seat on a project,
  describing what we're looking for, highlighting where their background
  matched, and asking about experience in areas the resume didn't show — with
  no scores, rankings, or confidential client details
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
| `--skills` / `--skills-file` | Comma-separated list (all weight 1.0), or a CSV file with `skill[,weight]` per line (default: `skills.csv`, or legacy `skills.txt`, in the current directory) |
| `--project` / `--project-file` | Project description text or file (default: `project.md` in the current directory) |
| `--prompt` / `--prompt-file` | Additional evaluation instructions (optional) |
| `--output` | Output `.xlsx` path (default: `evaluations.xlsx`) |
| `--provider` | `ica` (default) or `anthropic` |
| `--host` | Base URL override for the ICA backend |
| `--model` | Claude model for full evaluations (default: `ICA_MODEL`/`CLAUDE_MODEL` env, else `claude-opus-4-8`) |
| `--workers` | Concurrent evaluations (default: 4) |
| `--team-size` | Target team size: adds a **Team** tab with the strongest mix of N candidates and per-skill score totals — or four cost-shaped teams when rates are available (see Staffing rates) |
| `--rates` | Rates table `.xlsx` with columns Country, Line, Band, Rate, Currency (default: `rates.xlsx`, then `~/staffing/rates.xlsx`) |
| `--staff-rates` | Staff rates workbook to create/update (default: `staff-rates.xlsx`; pass `''` to disable) |
| `--screen` | Enable the two-tier low-cost screening mode (see below) |
| `--screen-model` | Screening model (default: `claude-haiku-4-5`) |
| `--screen-cutoff` | Minimum screening score to advance to full evaluation (default: 2) |

## Skill weights

`skills.csv` takes an optional second column: the skill's weight (default
`1.0`). Weights express relative importance to the project — they are passed
to the model (informing the overall fit score, recommendation, and team
selection), shown in the skill column headers when not 1.0, and used to
compute the **Weighted avg (0-4)** column on the Summary sheet
(`Σ score×weight / Σ weight`). Individual skill scores stay unweighted,
evidence-based 0-4. Quote skill names that contain commas:

```csv
# skill,weight   (weight is optional and defaults to 1.0)
Powershell scripting,2.0
Terraform
"Communication, stakeholder management",0.5
```

## Team selection

With `--team-size N`, a **Team** tab is added (second tab, after Summary)
showing the mix of N candidates that forms the strongest overall team. The
selection is made by the model from the full evaluations, optimizing for
complementary skill coverage — every skill area covered by at least one strong
member, weighted toward the project's critical initiatives — rather than just
the top N individual scores. The bottom of the team list has a **Team skill
total** row summing each skill column, indicating the team's strength in that
area, followed by the selection rationale.

```sh
uv run resume-reviewer ./resumes --team-size 5
```

Only fully-evaluated candidates are considered (with `--screen`, screened-out
candidates are excluded). If the selection call fails, the tab falls back to
the top N by overall score and says so in the rationale.

**Team characteristics from the project description are hard requirements.**
If `project.md` states composition needs — a project manager, a minimum number
of scripting specialists, a seniority or geographic mix — every selected team
(the single Team tab and all four cost shapes alike; cost posture never
overrides them) must satisfy them where the pool allows. Each team's rationale
includes a Composition line stating which members satisfy each requirement, or
why one couldn't be met with the available candidates.

## Staffing rates and cost-shaped teams

After each run, every evaluated candidate is synced into **`staff-rates.xlsx`**
(columns: Name, Geo, Line, Band, Rate, Currency). Candidates already present
are skipped — their rows are never modified — and new candidates are appended
with blank Geo/Line/Band. The workflow is two-phase:

1. **First run** scaffolds `staff-rates.xlsx`. If the file doesn't exist yet,
   you're offered a fast bootstrap: populate it with candidate names only —
   extracted with the cheap screening model — skipping skill evaluations
   entirely. The staffing team then fills in each person's Geo, Line, and Band
   using the **cascading dropdowns** on those columns: pick a Geo and the
   Line list filters to that geo's lines; pick a Line and the Band list
   filters to the bands available for that Geo+Line (Currency is a flat
   list). Sourced from the rates table via a hidden Lookups sheet. Note an
   Excel limitation: changing a parent (e.g. Geo) does not clear an
   already-selected child value — reselect Line/Band after changing Geo.
2. The rates table (`--rates`, defaulting to `rates.xlsx` then
   `~/staffing/rates.xlsx`; columns Country, Line, Band, Rate, Currency) is
   copied into a **Rates** tab of the workbook, and the Rate/Currency columns
   are live `VLOOKUP` formulas against it — they populate in Excel the moment
   Geo/Line/Band are chosen (or show `no match` for combinations not in the
   table). Re-running the tool refreshes the Rates tab, reports unmatched
   combinations, and uses the looked-up rates for team shapes.

Once **every** evaluated candidate has a rate and `--team-size N` is given,
the single Team tab is replaced by a **Team Shapes** tab with four teams of N,
each with a decent skill fit but a different cost posture:

| Shape | Objective |
|---|---|
| low cost | Minimize average cost rate while keeping an acceptable skill fit |
| medium cost | Balance cost and skill fit |
| high cost | Favor premium, senior talent where it buys real skill |
| cost not considered | The strongest possible skill fit, ignoring cost |

The tab opens with a comparison table — each shape's **skill fit score**
(weighted coverage on the 0–4 scale: each skill counts as the team's best
member on that skill) and **average cost rate** — so the staffing team can see
what a rate difference buys in skill. Below it, each team is broken out with
member rates, per-skill scores, skill totals, and the selection rationale.
Membership is chosen by the model; fit scores and averages are computed
deterministically so the four shapes are directly comparable.

`staff-rates.xlsx` and rate tables are `.gitignore`d — they contain names and
rates and should not be committed.

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
  --screen --screen-cutoff 2
```

Use this for large batches; skip it when you have a handful of resumes and want
full detail on everyone.
