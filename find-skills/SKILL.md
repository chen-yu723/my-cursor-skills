---
name: find-skills
description: Discover, search, and install Cursor Agent Skills from GitHub repositories or locally installed skill directories. Use when the user wants to find available skills, install a skill from GitHub using `npx skills add`, browse what skills are already installed, or asks about skill discovery and management.
---

# Find Skills

## Overview

Two main workflows:
1. **Local search** — browse already-installed skills under `~/.cursor/skills/` and `.cursor/skills/`
2. **Remote install** — pull a skill from a GitHub repository using the `npx skills` CLI

---

## 1. Search locally installed skills

Personal skills (available in all projects):
```
~/.cursor/skills/
```

Project skills (current repo only):
```
.cursor/skills/
```

To list what's installed, read the `SKILL.md` frontmatter in each subdirectory. The `description` field explains what the skill does and when to trigger it.

---

## 2. Install a skill from GitHub

### Syntax
```bash
npx skills add <github-repo-url> --skill <skill-name>
```

### Example
```bash
npx skills add https://github.com/vercel-labs/skills --skill find-skills
```

### Steps

1. Identify the GitHub repository that hosts the skill collection.
2. Identify the skill name (subdirectory name inside the repo's `skills/` or root folder).
3. Run the `npx skills add` command above in the terminal.
4. The CLI installs the skill into `~/.cursor/skills/<skill-name>/` by default.
5. Verify installation by checking that `~/.cursor/skills/<skill-name>/SKILL.md` exists.

### Flags

| Flag | Description |
|------|-------------|
| `--skill <name>` | Name of the specific skill to install |
| `--dir <path>` | Override install directory (default: `~/.cursor/skills/`) |
| `--project` | Install into `.cursor/skills/` (project-scoped) instead |

---

## 3. Browse a GitHub skills repository

To see all available skills in a repo before installing:
```bash
npx skills list https://github.com/<owner>/<repo>
```

This lists all available skill names with their descriptions so you can pick the right one.

---

## Tips

- After installing a new skill, restart Cursor or reload the window so the agent picks it up.
- Skill names are case-sensitive and use lowercase-with-hyphens convention.
- If a skill already exists locally, `npx skills add` will prompt before overwriting.
