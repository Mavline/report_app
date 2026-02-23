# AGENTS.md

## Purpose
This project uses a local memory bank (Cline-style) to preserve context across sessions, especially for recurring date-format issues in Excel parsing/merge logic.

## Memory Bank Location
Use `memory-bank/` in the project root.

## Files (7 total)
1. `memory-bank/projectbrief.md`
2. `memory-bank/productContext.md`
3. `memory-bank/systemPatterns.md`
4. `memory-bank/techContext.md`
5. `memory-bank/activeContext.md`
6. `memory-bank/progress.md`
7. `memory-bank/change-log.md`

## Read Order (start of each work session)
1. `memory-bank/activeContext.md`
2. `memory-bank/progress.md`
3. `memory-bank/change-log.md`
4. `memory-bank/systemPatterns.md`
5. `memory-bank/techContext.md`
6. `memory-bank/productContext.md`
7. `memory-bank/projectbrief.md`

## Update Rules
- Update `memory-bank/activeContext.md` at the start and end of meaningful work.
- Append to `memory-bank/change-log.md` after each code change, `git pull`, deployment-impacting fix, or discovered root cause.
- Update `memory-bank/progress.md` with completed/pending items after each session.
- Update `memory-bank/systemPatterns.md` only when architecture, parsing strategy, or invariants change.
- Use exact dates in `YYYY-MM-DD` format and include commit hashes when known.
- Record what was verified vs what was inferred.

## Date-Handling Rules (Important for this project)
- Treat date parsing/header matching as a high-risk area.
- When debugging merge mismatches, check both:
  - displayed drag-and-drop header labels
  - actual object keys returned by `XLSX.utils.sheet_to_json`
- Prefer strategies that keep header display values and row keys generated from the same source (`cell.w` when available).
- If adding date normalization, document supported variants in `memory-bank/systemPatterns.md`.
- Before concluding a fix works, verify merge with real-world Excel samples (especially different month abbreviations and separators).

## Git / History Rules
- Before changing date logic, inspect prior date-related commits and summarize the reason previous fixes were partial.
- Record `git pull` fast-forward updates in `memory-bank/change-log.md`.
- Do not delete historical notes; append corrections with dates.

## Entry Style
- Keep notes concise and factual.
- Separate facts, assumptions, and open questions.
- Prefer actionable language (what to verify next, what can break).

