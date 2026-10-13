# agent-skills

My agent skills, shared between Claude Code and Codex. One repo is the source of
truth; `scripts/link.sh` symlinks the skills I use into each harness's directory.

```
skills/
├── mine/                 my own skills
└── vendor/
    └── mattpocock/       git subtree of github.com/mattpocock/skills (never edit)
```

## Install on a new machine

```bash
git clone https://github.com/tomdhyang/agent-skills.git ~/.agents/skills-repo
~/.agents/skills-repo/scripts/link.sh
```

`link.sh` links into `~/.claude/skills` (Claude Code) and `~/.agents/skills`
(Codex). Re-run it after any change below.

## Add my own skill

```
skills/mine/<name>/
├── SKILL.md              frontmatter: name, description[, disable-model-invocation: true]
└── agents/openai.yaml    Codex metadata[, policy.allow_implicit_invocation: false]
```

A skill is user-invoked (slash / `$` only) in **both** harnesses or neither.
`link.sh` warns when the two files disagree.

- Model-invoked (default): `description` is for the model, keep trigger phrases.
- User-invoked: set both flags; `description` is a one-line summary for a human.

## Add another person's skills

```bash
git subtree add --prefix=skills/vendor/<name> <git-url> main --squash
```

Then add the bucket path(s) to `BUCKETS` in `scripts/link.sh`.

## Update a vendored source

```bash
git subtree pull --prefix=skills/vendor/mattpocock https://github.com/mattpocock/skills.git main --squash
./scripts/link.sh
```

Read `skills/vendor/mattpocock/CHANGELOG.md` to see what changed.

## Rules

- Never edit files under `skills/vendor/`. Standing behaviour goes in `~/.claude/CLAUDE.md`
  or `~/.codex/AGENTS.md`; it survives every subtree pull with zero conflicts.
- "Uninstall" a skill by leaving it out of a linked bucket, not by deleting it.
- Don't name a skill folder `synced` or anything starting with `anthropic-skills`
  (reserved by Claude Code).
