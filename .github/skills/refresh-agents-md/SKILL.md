---
name: refresh-agents-md
description: 'Update agents.md from the latest coding session without bloat. Use when asked to refresh project context, architecture notes, commands, or key file roles while keeping the guide concise.'
argument-hint: 'What changed this session that should be reflected in agents.md?'
user-invocable: true
disable-model-invocation: false
---

# Refresh agents.md (Concise)

Update `agents.md` so it reflects the newest session-level reality while staying short, factual, and durable.

## When To Use
- User asks to "update" or "refresh" `agents.md`.
- Session changed architecture, key workflows, paths, commands, or conventions.
- Existing `agents.md` contains stale claims or drift.

## Inputs To Gather
1. Session deltas: what changed in code, structure, deploy flow, and config.
2. Current `agents.md` sections impacted by those deltas.
3. Evidence from workspace files (not assumptions).

## Procedure
1. Read current `agents.md` and identify stale, missing, or verbose areas.
2. Collect only durable session facts:
- Stable endpoints, key files, run/test commands, deployment notes, conventions.
- Exclude transient debugging details, one-off logs, timestamps, and personal TODOs.
3. Decide update scope:
- Minor drift: patch only affected bullets/rows.
- Major drift: rewrite only impacted sections, keep structure recognizable.
4. Apply anti-bloat rules during edit:
- Keep each section to concise bullets or compact tables.
- Prefer replacing redundant text over appending new prose.
- Remove superseded lines instead of preserving history.
5. Validate facts against source files.
6. Perform a final concision pass (delete repetition, shorten wording, keep high signal).

## Branching Logic
- If a fact is unverifiable from workspace files: omit it and flag as open question.
- If two lines express the same idea: keep the clearer one, delete the other.
- If a section grows beyond usefulness: collapse to 3-6 bullets and link key files.
- If no durable project knowledge changed: make no file edit and report "no update needed".

## Quality Checks
- Accuracy: every new statement is traceable to workspace code/docs.
- Brevity: no section contains filler or duplicate guidance.
- Durability: content helps future sessions, not just this one moment.
- Coverage: reflects latest architecture/workflow changes from this session.

## Output Contract
- Update `agents.md` with minimal diff.
- Summarize what changed in 3-6 bullets.
- List any omitted uncertain items as open questions.

## Example Prompts
- `/refresh-agents-md Capture this session's backend endpoint changes in agents.md.`
- `/refresh-agents-md Update agents.md from today's sidebar sync workflow edits.`
- `/refresh-agents-md Reconcile agents.md with current folder structure and commands, keep it concise.`
