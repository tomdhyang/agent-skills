# agent-skills

`README.md` holds every recipe (install, sync, update a vendor, add a skill, add
a vendor) and the rules. Follow it for any change here.

- Edit only `skills/mine/` and `scripts/link.sh`; `skills/vendor/` is read-only.
- Before writing or editing a `SKILL.md`, call the Skill tool with
  "writing-for-agents".
- After any change, finish by running `scripts/link.sh` and reading its output:
  a `warn` line means the two harness flags disagree and the change is not done.
