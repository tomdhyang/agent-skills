# agent-skills

One repo is the source of truth for my agent skills. Claude Code reads
`~/.claude/skills`, Codex reads `~/.agents/skills`; `scripts/link.sh` fills both
with per-skill symlinks into this clone. The clone lives at `~/.agents/skills-repo`.

```
skills/mine/<name>/            my skills
skills/vendor/<source>/        other people's repos, each a git subtree, read-only
```

Each recipe below ends when its last line holds. `link.sh` is idempotent: run it
whenever in doubt.

## Install on a new machine

```bash
git clone https://github.com/tomdhyang/agent-skills.git ~/.agents/skills-repo
~/.agents/skills-repo/scripts/link.sh
```

Done when `link.sh` reports the same skill count for both directories and prints
no `warn` lines.

## Sync changes I pushed from another machine

```bash
cd ~/.agents/skills-repo && git pull && ./scripts/link.sh
```

Edited skill bodies take effect at the next session with no further step; a
skill added, removed, or renamed appears only after `link.sh`, which is why the
command always runs it.

## Update a vendored source

```bash
cd ~/.agents/skills-repo
git subtree pull --prefix=skills/vendor/mattpocock https://github.com/mattpocock/skills.git main --squash
./scripts/link.sh
git push
```

Done when `git push` succeeds; until then other machines keep the old version.
`skills/vendor/mattpocock/CHANGELOG.md` says what changed.

## Add a skill of my own

1. Create `skills/mine/<name>/SKILL.md` and `skills/mine/<name>/agents/openai.yaml`.
2. Decide invocation, once, for both harnesses:
   - **Model-invoked** (default): omit `disable-model-invocation`, omit `policy` in
     `openai.yaml`. Write `description` for the model: what it does, then the
     triggers.
   - **User-invoked** (slash / `$` only): `disable-model-invocation: true` in
     `SKILL.md` and `policy.allow_implicit_invocation: false` in `openai.yaml`.
     Write `description` as one line for a human browsing the menu.
3. `./scripts/link.sh`, then commit and push.

Done when `link.sh` prints no `warn` line for the new skill and it appears in
both skill directories.

## Add another person's skills

```bash
git subtree add --prefix=skills/vendor/<name> <git-url> main --squash
```

Then append each directory whose children are skills to `BUCKETS` in
`scripts/link.sh`, run it, commit, push. Done when the vendor's skills appear in
both skill directories.

## Rules

- `skills/vendor/` is read-only. Standing behaviour that should override a
  vendor skill goes in `~/.claude/CLAUDE.md` or `~/.codex/AGENTS.md`, where a
  subtree pull cannot conflict with it.
- A skill is retired by leaving its bucket out of `BUCKETS`; the folder stays.
- `SKILL.md` frontmatter stays within the Agent Skills spec
  (`name`, `description`, `license`, `compatibility`, `metadata`,
  `allowed-tools`) plus `disable-model-invocation`. Claude-only keys such as
  `context`, `hooks`, `paths` make claude.ai and API uploads fail.
- Skill folder names `synced` and `anthropic-skills*` are reserved by Claude Code.
