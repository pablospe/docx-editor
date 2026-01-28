<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

# Claude Instructions for docx-editor

## Quick Reference

```bash
# Development
uv sync --dev              # Install dependencies
uv run pytest              # Run tests
uv run ruff check .        # Lint
uv run ruff format .       # Format
uv run mkdocs serve        # Preview docs
```

## Code Quality Principles (CRITICAL)

**MANDATORY: These principles are CRITICAL for all code changes.**

### Core Principles

1. **KISS (Keep It Simple, Stupid)** - Choose the simplest solution that works.
2. **DRY (Don't Repeat Yourself)** - Search for and reuse existing code and patterns.
3. **YAGNI (You Aren't Gonna Need It)** - Only build what you need now.
4. **Leverage Existing Libraries** - Use defusedxml for XML parsing. Don't reinvent wheels.
5. **No Magic** - Explicit configuration only. No guessing, parsing, or fallbacks.
6. **Clear Separation of Concerns** - One purpose per file/class.
7. **Named Objects Over Tuples** - For 3+ return values, use dataclasses with named fields.
8. **Small, Focused Interfaces** - Minimal abstract base classes. Easy to extend.

### Before Writing Code

- [ ] Searched codebase for existing solutions
- [ ] Verified no library can solve this
- [ ] Kept it simple (KISS)

### Common Mistakes to Avoid

- ❌ **DON'T** create new files unless necessary → ✅ Edit existing files
- ❌ **DON'T** use `git add -A` → ✅ Use `git add <specific-file>`
- ❌ **DON'T** reinvent existing functionality → ✅ Reuse patterns
- ❌ **DON'T** skip tests → ✅ Write tests for changes
