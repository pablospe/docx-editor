# docx-editor

[![Release](https://img.shields.io/github/v/release/pablospe/docx-editor)](https://img.shields.io/github/v/release/pablospe/docx-editor)
[![Build status](https://img.shields.io/github/actions/workflow/status/pablospe/docx-editor/main.yml?branch=main)](https://github.com/pablospe/docx-editor/actions/workflows/main.yml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/pablospe/docx-editor/branch/main/graph/badge.svg)](https://codecov.io/gh/pablospe/docx-editor)
[![Commit activity](https://img.shields.io/github/commit-activity/m/pablospe/docx-editor)](https://img.shields.io/github/commit-activity/m/pablospe/docx-editor)
[![License](https://img.shields.io/github/license/pablospe/docx-editor)](https://img.shields.io/github/license/pablospe/docx-editor)

Pure Python library for Word document track changes and comments, without requiring Microsoft Word.

> **Note:** The PyPI package is named `docx-editor` because `docx-edit` was too similar to an existing package.

- **Github repository**: <https://github.com/pablospe/docx-editor/>
- **Documentation**: <https://pablospe.github.io/docx-editor/>

## Features

- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-editor
```

## Quick Start

```python
from docx_editor import Document

with Document.open("contract.docx") as doc:
    # Track changes
    doc.replace("30 days", "60 days")
    doc.insert_after("Section 5", "New clause")
    doc.delete("obsolete text")

    # Comments
    doc.add_comment("Section 5", "Please review")

    # Revision management
    revisions = doc.list_revisions()
    doc.accept_revision(revision_id=1)

    doc.save()
```

---

Repository initiated with [fpgmaas/cookiecutter-uv](https://github.com/fpgmaas/cookiecutter-uv).
