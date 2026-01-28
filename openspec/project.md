# Project Context

## Purpose

A Python library for editing Microsoft Word (.docx) files with tracked changes support. Enables programmatic document manipulation while preserving revision history - essential for legal, enterprise, and collaborative workflows.

**Key differentiator:** Native track changes support (insertions, deletions, comments) that Word can display and accept/reject.

## Tech Stack

- **Language:** Python 3.10+
- **Package manager:** uv
- **XML parsing:** defusedxml (secure parsing)
- **Testing:** pytest, pytest-cov
- **Linting/Formatting:** ruff
- **Documentation:** mkdocs with material theme

## Project Conventions

### Code Style

- Follow ruff defaults (based on Black formatting)
- Type hints required for public APIs
- Docstrings for public functions (Google style)
- No magic - explicit configuration only

### Architecture Patterns

```
docx_editor/
├── document.py       # Public API facade (Document class)
├── xml_editor.py     # Core XML manipulation (XMLEditor, DocxXMLEditor)
├── track_changes.py  # RevisionManager - track changes operations
├── comments.py       # CommentManager - comment operations
├── workspace.py      # Workspace management (.docx/{stem}/ folders)
├── exceptions.py     # Custom exception hierarchy
└── ooxml/           # OOXML utilities (pack/unpack)
```

**Key patterns:**
- Facade pattern: `Document` wraps `RevisionManager`/`CommentManager`
- Workspace pattern: Unpacks .docx to temp folder, edits in place, repacks on save
- Attribute injection: Auto-adds tracking metadata (author, date, IDs)

### Testing Strategy

- pytest with fixtures for test documents
- 98% code coverage target
- Test documents in `tests/fixtures/`
- Run: `uv run pytest` or `uv run pytest -n auto` for parallel

### Git Workflow

- Main branch: `main`
- Feature branches from `main`
- Commits without Co-Authored-By attribution
- PR-based workflow

## Domain Context

### OOXML Structure

Word documents are ZIP archives containing XML files:
- `word/document.xml` - Main content
- `word/comments.xml` - Comments
- `word/people.xml` - Authors

### Key XML Elements

| Element | Purpose |
|---------|---------|
| `<w:p>` | Paragraph |
| `<w:r>` | Run (text with formatting) |
| `<w:t>` | Text content |
| `<w:ins>` | Tracked insertion |
| `<w:del>` | Tracked deletion |
| `<w:delText>` | Deleted text (inside `<w:del>`) |
| `<w:commentRangeStart/End>` | Comment anchors |

### Revision Tracking

- Insertions: Wrap content in `<w:ins>` with author/date/id
- Deletions: Convert `<w:t>` to `<w:delText>`, wrap in `<w:del>`
- Each revision gets unique `w:id` and `w:rsidR` (session ID)

## Important Constraints

1. **Security:** Use defusedxml for all XML parsing (prevents XXE attacks)
2. **Compatibility:** Output must open in Microsoft Word without warnings
3. **Non-destructive:** Operations must preserve existing formatting and metadata
4. **Namespace handling:** OOXML uses many namespaces (w, w14, w16du, etc.)

## External Dependencies

- **python-docx:** Not used (we handle OOXML directly for track changes)
- **defusedxml:** Required for secure XML parsing
- No external services or APIs
