#!/usr/bin/env python3
"""Evaluate a directory of resumes against a skill list and project description
using the Claude API, and export the results as an Excel spreadsheet.

Backends (following the uft-to-flaui-agentic-framework pattern):
    ica        IBM Consulting Advantage (default) — Anthropic-native Messages
               endpoint fronting Anthropic-on-Bedrock. Uses the stock anthropic
               SDK pointed at the /ica base URL.
               Env: ICA_ANTHROPIC_API_KEY (or ICA_API_KEY), ICA_MODEL, ICA_HOST
    anthropic  Direct Claude API. Env: ANTHROPIC_API_KEY, CLAUDE_MODEL

Usage:
    export ICA_ANTHROPIC_API_KEY=<your-key>
    python resume_reviewer.py ./resumes \
        --skills "Python, Kubernetes, Terraform, CI/CD" \
        --project-file project.md \
        --prompt "Weight hands-on infrastructure experience heavily." \
        --output evaluations.xlsx
"""

from __future__ import annotations

import argparse
import base64
import os
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import List, Literal, Optional, Type, TypeVar

import anthropic
from pydantic import BaseModel, Field, ValidationError

RESUME_EXTENSIONS = {".pdf", ".docx", ".txt", ".md"}
DEFAULT_MODEL = "claude-opus-4-8"
DEFAULT_SCREEN_MODEL = "claude-haiku-4-5"
# The /ica root, NOT /ica/v1 — the SDK appends /v1/messages itself.
ICA_DEFAULT_BASE_URL = "https://api.nextgen-beta.ica.ibm.com/ica"


# --------------------------------------------------------------------------
# Structured output schema
# --------------------------------------------------------------------------

class SkillEvaluation(BaseModel):
    skill: str = Field(description="The skill being evaluated, exactly as given in the skills list")
    score: int = Field(description="Score from 1 (no evidence) to 5 (expert, strong direct evidence)")
    evidence: str = Field(description="Specific evidence from the resume supporting the score, or why evidence is lacking")


class ResumeEvaluation(BaseModel):
    candidate_name: str = Field(description="Candidate's full name as it appears on the resume")
    overall_score: int = Field(description="Overall fit for the project, 1 (poor fit) to 5 (excellent fit)")
    recommendation: Literal["strong yes", "yes", "maybe", "no"]
    summary: str = Field(description="2-4 sentence assessment of the candidate's fit for the project")
    strengths: List[str] = Field(description="Top strengths relevant to the project")
    gaps: List[str] = Field(description="Notable gaps or risks relative to the project needs")
    skill_evaluations: List[SkillEvaluation] = Field(
        description="One entry per skill in the skills list, in the same order"
    )


class ScreenResult(BaseModel):
    candidate_name: str = Field(description="Candidate's full name as it appears on the resume")
    overall_score: int = Field(description="Overall fit for the project, 1 (poor fit) to 5 (excellent fit)")
    recommendation: Literal["strong yes", "yes", "maybe", "no"]
    summary: str = Field(description="1-2 sentence assessment of the candidate's fit")


@dataclass
class Result:
    file: Path
    evaluation: Optional[ResumeEvaluation] = None
    error: Optional[str] = None
    stage: str = "full"  # "full" or "screened out"


# --------------------------------------------------------------------------
# Resume loading
# --------------------------------------------------------------------------

def extract_docx_text(path: Path) -> str:
    import docx  # python-docx

    document = docx.Document(str(path))
    parts = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def resume_content_block(path: Path) -> dict:
    """Build the message content block carrying the resume itself."""
    if path.suffix.lower() == ".pdf":
        data = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": data},
        }
    if path.suffix.lower() == ".docx":
        text = extract_docx_text(path)
    else:
        text = path.read_text(encoding="utf-8", errors="replace")
    if not text.strip():
        raise ValueError("resume file is empty after text extraction")
    return {"type": "text", "text": f"<resume>\n{text}\n</resume>"}


# --------------------------------------------------------------------------
# Backends
# --------------------------------------------------------------------------

def make_client(provider: str, host: Optional[str] = None) -> anthropic.Anthropic:
    """Build an Anthropic-SDK client for the chosen provider.

    ICA exposes an Anthropic-native Messages endpoint (/ica/v1/messages), so the
    stock SDK works unchanged with a base_url override — same approach as
    uft-to-flaui-agentic-framework's ICABackend.
    """
    if provider == "ica":
        key = os.environ.get("ICA_ANTHROPIC_API_KEY") or os.environ.get("ICA_API_KEY")
        if not key:
            sys.exit("error: ICA_ANTHROPIC_API_KEY (or ICA_API_KEY) is not set")
        base_url = (host or os.environ.get("ICA_HOST", ICA_DEFAULT_BASE_URL)).rstrip("/")
        return anthropic.Anthropic(api_key=key, base_url=base_url)
    if provider == "anthropic":
        if not os.environ.get("ANTHROPIC_API_KEY"):
            sys.exit("error: ANTHROPIC_API_KEY is not set")
        return anthropic.Anthropic()
    sys.exit(f"error: unknown provider {provider!r} (choose: ica, anthropic)")


def resolve_model(provider: str, explicit: Optional[str]) -> str:
    """--model flag > provider-specific env var > default."""
    env_key = "ICA_MODEL" if provider == "ica" else "CLAUDE_MODEL"
    return explicit or os.environ.get(env_key) or DEFAULT_MODEL


# --------------------------------------------------------------------------
# Evaluation
# --------------------------------------------------------------------------

def build_system_prompt(prompt: str, project: str, skills: List[str]) -> str:
    skills_list = "\n".join(f"- {s}" for s in skills)
    sections = [
        "You are an experienced technical recruiter and hiring manager. "
        "Evaluate the resume you are given against the project description and skills below. "
        "Base every score strictly on evidence in the resume; do not give the benefit of the doubt "
        "for skills that are not demonstrated. Score each skill from 1 (no evidence) to 5 "
        "(expert with strong direct evidence). Provide one skill_evaluations entry per listed "
        "skill, using the exact skill names given, in the same order.",
        f"<project_description>\n{project}\n</project_description>",
        f"<skills_to_evaluate>\n{skills_list}\n</skills_to_evaluate>",
    ]
    if prompt:
        sections.append(f"<additional_evaluation_instructions>\n{prompt}\n</additional_evaluation_instructions>")
    return "\n\n".join(sections)


T = TypeVar("T", bound=BaseModel)


def call_structured(
    client: anthropic.Anthropic,
    model: str,
    system_prompt: str,
    user_content: list,
    schema_model: Type[T],
    tool_name: str,
    max_tokens: int,
) -> T:
    """Tool-forced structured output, validated with Pydantic.

    Forcing a single tool whose input_schema is the result schema works on both
    the direct Claude API and ICA's Bedrock-fronted Messages endpoint (the
    surface uft-to-flaui-agentic-framework verified live), unlike
    output_config.format which is unverified through the proxy. Forced
    tool_choice also means no thinking parameter. One retry on validation
    failure, feeding the error back as a tool_result.
    """
    tool = {
        "name": tool_name,
        "description": "Record the evaluation result, strictly conforming to the schema.",
        "input_schema": schema_model.model_json_schema(),
    }
    messages: list = [{"role": "user", "content": user_content}]
    last_err: Optional[ValidationError] = None
    for attempt in range(2):
        response = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=[
                {
                    "type": "text",
                    "text": system_prompt,
                    # Shared across every resume in the run — cache it.
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            tools=[tool],
            tool_choice={"type": "tool", "name": tool_name},
            messages=messages,
        )
        if response.stop_reason == "refusal":
            raise RuntimeError("the model declined to evaluate this resume (stop_reason: refusal)")
        if response.stop_reason == "max_tokens":
            raise RuntimeError("evaluation was truncated (stop_reason: max_tokens)")
        block = next(
            (b for b in response.content if b.type == "tool_use" and b.name == tool_name),
            None,
        )
        if block is None:
            raise RuntimeError("model did not return the expected tool_use block")
        try:
            return schema_model.model_validate(block.input)
        except ValidationError as exc:
            last_err = exc
            if attempt == 0:
                messages = messages + [
                    {"role": "assistant", "content": response.content},
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "tool_result",
                                "tool_use_id": block.id,
                                "content": f"Arguments failed validation: {exc}. "
                                f"Call {tool_name} again with corrected arguments.",
                                "is_error": True,
                            }
                        ],
                    },
                ]
    raise RuntimeError(f"model output failed schema validation after retry: {last_err}")


def evaluate_resume(
    client: anthropic.Anthropic,
    model: str,
    system_prompt: str,
    path: Path,
) -> ResumeEvaluation:
    return call_structured(
        client,
        model,
        system_prompt,
        [resume_content_block(path), {"type": "text", "text": "Evaluate this resume."}],
        ResumeEvaluation,
        tool_name="record_evaluation",
        max_tokens=16000,
    )


def screen_resume(
    client: anthropic.Anthropic,
    model: str,
    system_prompt: str,
    path: Path,
) -> ScreenResult:
    """Cheap first-pass screen with a lower-cost model (no per-skill breakdown)."""
    return call_structured(
        client,
        model,
        system_prompt,
        [
            resume_content_block(path),
            {
                "type": "text",
                "text": "Quickly screen this resume: give only the candidate name, "
                "an overall fit score, a recommendation, and a 1-2 sentence summary.",
            },
        ],
        ScreenResult,
        tool_name="record_screen_result",
        max_tokens=2000,
    )


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------

RECOMMENDATION_FILLS = {
    "strong yes": "C6EFCE",
    "yes": "E2EFDA",
    "maybe": "FFF2CC",
    "no": "F8CBAD",
}


def write_workbook(results: List[Result], skills: List[str], output: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    wrap = Alignment(wrap_text=True, vertical="top")

    def style_header(ws) -> None:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        ws.freeze_panes = "A2"

    # --- Summary sheet -----------------------------------------------------
    ws = wb.active
    assert ws is not None
    ws.title = "Summary"
    headers = ["Candidate", "File", "Overall (1-5)", "Recommendation", "Stage"]
    headers += [f"{s} (1-5)" for s in skills]
    headers += ["Strengths", "Gaps", "Summary"]
    ws.append(headers)

    ok = [(r, r.evaluation) for r in results if r.evaluation is not None]
    ok.sort(key=lambda pair: pair[1].overall_score, reverse=True)

    for r, ev in ok:
        scores_by_skill = {se.skill.strip().lower(): se.score for se in ev.skill_evaluations}
        row = [ev.candidate_name, r.file.name, ev.overall_score, ev.recommendation, r.stage]
        row += [scores_by_skill.get(s.strip().lower(), "") for s in skills]
        row += ["\n".join(ev.strengths), "\n".join(ev.gaps), ev.summary]
        ws.append(row)
        rec_cell = ws.cell(row=ws.max_row, column=4)
        fill = RECOMMENDATION_FILLS.get(ev.recommendation)
        if fill:
            rec_cell.fill = PatternFill("solid", fgColor=fill)
        for col in (len(headers) - 2, len(headers) - 1, len(headers)):
            ws.cell(row=ws.max_row, column=col).alignment = wrap

    style_header(ws)
    widths = [24, 24, 12, 15, 13] + [14] * len(skills) + [45, 45, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # --- Details sheet -----------------------------------------------------
    ws_d = wb.create_sheet("Skill Details")
    ws_d.append(["Candidate", "Skill", "Score (1-5)", "Evidence"])
    for r, ev in ok:
        for se in ev.skill_evaluations:
            ws_d.append([ev.candidate_name, se.skill, se.score, se.evidence])
            ws_d.cell(row=ws_d.max_row, column=4).alignment = wrap
    style_header(ws_d)
    for col, w in zip("ABCD", [24, 24, 12, 90]):
        ws_d.column_dimensions[col].width = w

    # --- Errors sheet ------------------------------------------------------
    failed = [r for r in results if r.error is not None]
    if failed:
        ws_e = wb.create_sheet("Errors")
        ws_e.append(["File", "Error"])
        for r in failed:
            ws_e.append([r.file.name, r.error])
        style_header(ws_e)
        ws_e.column_dimensions["A"].width = 30
        ws_e.column_dimensions["B"].width = 100

    wb.save(output)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def read_arg_or_file(value: Optional[str], file_value: Optional[str], name: str, required: bool) -> str:
    if value and file_value:
        sys.exit(f"error: pass --{name} or --{name}-file, not both")
    if file_value:
        return Path(file_value).read_text(encoding="utf-8")
    if value:
        return value
    if required:
        sys.exit(f"error: --{name} or --{name}-file is required")
    return ""


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Evaluate resumes against skills and a project description; export to Excel."
    )
    parser.add_argument("resumes_dir", help="Directory containing resumes (.pdf, .docx, .txt, .md)")
    parser.add_argument("--prompt", help="Additional evaluation instructions")
    parser.add_argument("--prompt-file", help="File containing additional evaluation instructions")
    parser.add_argument("--project", help="Project description text")
    parser.add_argument("--project-file", help="File containing the project description")
    parser.add_argument("--skills", help="Comma-separated list of skills to evaluate")
    parser.add_argument("--skills-file", help="File with one skill per line")
    parser.add_argument("--output", default="evaluations.xlsx", help="Output .xlsx path (default: evaluations.xlsx)")
    parser.add_argument(
        "--provider",
        default=os.environ.get("PROVIDER", "ica"),
        choices=["ica", "anthropic"],
        help="LLM backend (default: ica — IBM Consulting Advantage)",
    )
    parser.add_argument("--host", help="Base URL override for the ICA backend (the /ica root)")
    parser.add_argument(
        "--model",
        default=None,
        help=f"Claude model ID (default: env ICA_MODEL / CLAUDE_MODEL, else {DEFAULT_MODEL})",
    )
    parser.add_argument("--workers", type=int, default=4, help="Concurrent evaluations (default: 4)")
    parser.add_argument(
        "--screen",
        action="store_true",
        help="Two-tier mode: screen every resume with a cheaper model first; only candidates "
        "at or above --screen-cutoff get the full evaluation with --model",
    )
    parser.add_argument(
        "--screen-model",
        default=DEFAULT_SCREEN_MODEL,
        help=f"Model for the screening pass (default: {DEFAULT_SCREEN_MODEL})",
    )
    parser.add_argument(
        "--screen-cutoff",
        type=int,
        default=3,
        help="Minimum screening score (1-5) to advance to full evaluation (default: 3)",
    )
    args = parser.parse_args()

    resumes_dir = Path(args.resumes_dir)
    if not resumes_dir.is_dir():
        sys.exit(f"error: {resumes_dir} is not a directory")

    project = read_arg_or_file(args.project, args.project_file, "project", required=True)

    skills_text = read_arg_or_file(args.skills, args.skills_file, "skills", required=True)
    if args.skills_file:
        skills = [line.strip() for line in skills_text.splitlines() if line.strip()]
    else:
        skills = [s.strip() for s in skills_text.split(",") if s.strip()]
    if not skills:
        sys.exit("error: no skills provided")

    prompt = read_arg_or_file(args.prompt, args.prompt_file, "prompt", required=False)

    files = sorted(
        p for p in resumes_dir.iterdir()
        if p.is_file() and p.suffix.lower() in RESUME_EXTENSIONS and not p.name.startswith(".")
    )
    if not files:
        sys.exit(f"error: no resumes ({', '.join(sorted(RESUME_EXTENSIONS))}) found in {resumes_dir}")

    model = resolve_model(args.provider, args.model)
    print(
        f"Evaluating {len(files)} resume(s) against {len(skills)} skill(s) "
        f"with {model} via {args.provider}..."
    )

    client = make_client(args.provider, args.host)
    system_prompt = build_system_prompt(prompt, project, skills)

    lock = threading.Lock()

    def run_all(paths: List[Path], worker, label: str) -> List[Result]:
        out: List[Result] = []
        done = 0
        with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
            futures = {pool.submit(worker, p): p for p in paths}
            for future in as_completed(futures):
                result = future.result()
                with lock:
                    out.append(result)
                    done += 1
                    status = "ok" if result.evaluation else f"FAILED ({result.error})"
                    print(f"  [{label} {done}/{len(paths)}] {result.file.name}: {status}")
        return out

    def full_eval(path: Path) -> Result:
        try:
            ev = evaluate_resume(client, model, system_prompt, path)
            return Result(file=path, evaluation=ev)
        except Exception as exc:  # keep going; failures land on the Errors sheet
            return Result(file=path, error=f"{type(exc).__name__}: {exc}")

    if args.screen:
        def screen_one(path: Path) -> Result:
            try:
                sr = screen_resume(client, args.screen_model, system_prompt, path)
                ev = ResumeEvaluation(
                    candidate_name=sr.candidate_name,
                    overall_score=sr.overall_score,
                    recommendation=sr.recommendation,
                    summary=f"[Screened by {args.screen_model}] {sr.summary}",
                    strengths=[],
                    gaps=[],
                    skill_evaluations=[],
                )
                return Result(file=path, evaluation=ev, stage="screened out")
            except Exception as exc:
                return Result(file=path, error=f"{type(exc).__name__}: {exc}")

        print(f"Screening pass with {args.screen_model} (cutoff: {args.screen_cutoff})...")
        screened = run_all(files, screen_one, "screen")
        advancing = [
            r.file for r in screened
            if r.evaluation and r.evaluation.overall_score >= args.screen_cutoff
        ]
        results = [
            r for r in screened
            if r.error or (r.evaluation and r.evaluation.overall_score < args.screen_cutoff)
        ]
        print(
            f"Screening done: {len(advancing)} of {len(files)} advance to full evaluation "
            f"({len(files) - len(advancing)} screened out or failed)."
        )
        if advancing:
            print(f"Full evaluation with {model}...")
            results += run_all(advancing, full_eval, "eval")
    else:
        results = run_all(files, full_eval, "eval")

    output = Path(args.output)
    write_workbook(results, skills, output)

    succeeded = sum(1 for r in results if r.evaluation)
    print(f"\nWrote {output} — {succeeded} evaluated, {len(results) - succeeded} failed.")
    return 0 if succeeded else 1


if __name__ == "__main__":
    sys.exit(main())
