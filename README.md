# docx-edit

[![Release](https://img.shields.io/github/v/release/pablospe/docx-edit)](https://img.shields.io/github/v/release/pablospe/docx-edit)
[![Build status](https://img.shields.io/github/actions/workflow/status/pablospe/docx-edit/main.yml?branch=main)](https://github.com/pablospe/docx-edit/actions/workflows/main.yml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/pablospe/docx-edit/branch/main/graph/badge.svg)](https://codecov.io/gh/pablospe/docx-edit)
[![Commit activity](https://img.shields.io/github/commit-activity/m/pablospe/docx-edit)](https://img.shields.io/github/commit-activity/m/pablospe/docx-edit)
[![License](https://img.shields.io/github/license/pablospe/docx-edit)](https://img.shields.io/github/license/pablospe/docx-edit)

Pure Python library for Word document track changes and comments, without requiring Microsoft Word.

- **Github repository**: <https://github.com/pablospe/docx-edit/>
- **Documentation**: <https://pablospe.github.io/docx-edit/>

## Features

- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-edit
```

## Quick Start

```python
from docx_edit import Document

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
