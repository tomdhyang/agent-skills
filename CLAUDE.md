# agent-skills

Personal skill repo. See `README.md` for layout and workflows.

- `skills/mine/` is mine. `skills/vendor/*` is other people's, pulled in with `git subtree`: do not edit anything under `skills/vendor/`.
- Every skill in `skills/mine/` carries `SKILL.md` and `agents/openai.yaml`. Keep `disable-model-invocation` (Claude Code) and `policy.allow_implicit_invocation` (Codex) in agreement.
- Frontmatter in `skills/mine/*/SKILL.md` stays within the Agent Skills spec plus `disable-model-invocation`: `name`, `description`, `license`, `compatibility`, `metadata`, `allowed-tools`. Claude-only keys (`context`, `hooks`, `paths`, `argument-hint`) break claude.ai and API uploads.
- After adding, removing, or renaming a skill, run `scripts/link.sh`.
- When editing a skill, call the Skill tool with "writing-for-agents" first.
