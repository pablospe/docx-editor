---
name: docx
description: "Comprehensive document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction. When Claude needs to work with professional documents (.docx files) for: (1) Creating new documents, (2) Modifying or editing content, (3) Working with tracked changes, (4) Adding comments, or any other document tasks"
---

# DOCX creation, editing, and analysis

## Overview

A user may ask you to create, edit, or analyze the contents of a .docx file. A .docx file is essentially a ZIP archive containing XML files and other resources. You have different tools and workflows available for different tasks.

## Workflow Decision Tree

```
What do you need to do?
|
+-- Read/Analyze Content
|   Use pandoc for text extraction (see "Reading and analyzing content")
|
+-- Create New Document
|   Use docx-js (see "Creating a new Word document")
|
+-- Edit Existing Document
    Use docx_edit Python library (see "Editing an existing Word document")
    - Tracked changes (redlining)
    - Comments (add, reply, resolve)
    - Accept/reject revisions
```

## Reading and analyzing content

### Text extraction

Convert the document to markdown using pandoc. Pandoc provides excellent support for preserving document structure and can show tracked changes:

```bash
# Convert document to markdown with tracked changes
pandoc --track-changes=all path-to-file.docx -o output.md

# Options: --track-changes=accept/reject/all
```

### Raw XML access

For comments, complex formatting, document structure, embedded media, and metadata, unpack the document:

```bash
unzip document.docx -d unpacked/
```

Key file structures:
* `word/document.xml` - Main document contents
* `word/comments.xml` - Comments referenced in document.xml
* `word/media/` - Embedded images and media files
* Tracked changes use `<w:ins>` (insertions) and `<w:del>` (deletions) tags

## Creating a new Word document

When creating a new Word document from scratch, use **docx-js**, which allows you to create Word documents using JavaScript/TypeScript.

### Workflow

1. **MANDATORY - READ ENTIRE FILE**: Read [`docx-js.md`](docx-js.md) (~350 lines) completely from start to finish. Read the full file content for detailed syntax, critical formatting rules, and best practices before proceeding with document creation.
2. Create a JavaScript/TypeScript file using Document, Paragraph, TextRun components
3. Export as .docx using Packer.toBuffer()

## Editing an existing Word document

Use the **docx_edit** Python library for all editing operations. It handles tracked changes, comments, and revisions with a simple API.

### Installation

```bash
pip install docx-edit
```

### Basic Usage

```python
from docx_edit import Document

# Open document with author name for tracked changes
doc = Document.open("contract.docx", author="Reviewer Name")

# Make changes (automatically tracked)
doc.replace("old text", "new text")  # Tracked replacement
doc.delete("text to delete")          # Tracked deletion
doc.insert_after("anchor", "new text") # Tracked insertion

# Save and close
doc.save()  # Overwrites original
# or doc.save("reviewed.docx")  # Save to new file
doc.close()
```

### Track Changes API

```python
from docx_edit import Document

doc = Document.open("document.docx", author="Editor")

# Replace text (creates tracked deletion + insertion)
doc.replace("30 days", "60 days")

# Delete text (creates tracked deletion)
doc.delete("unnecessary clause")

# Insert after anchor text (creates tracked insertion)
doc.insert_after("Section 3.", " Additional terms apply.")

doc.save("edited.docx")
doc.close()
```

### Comments API

```python
from docx_edit import Document

doc = Document.open("document.docx", author="Reviewer")

# Add a comment anchored to text
doc.add_comment("ambiguous term", "Please clarify this term")

# List all comments
comments = doc.list_comments()
for c in comments:
    print(f"ID: {c['id']}, Author: {c['author']}, Text: {c['text']}")

# Reply to a comment
doc.reply_to_comment(comment_id=1, "I agree, needs clarification")

# Resolve or delete comments
doc.resolve_comment(comment_id=1)
doc.delete_comment(comment_id=2)

doc.save()
doc.close()
```

### Revision Management API

```python
from docx_edit import Document

doc = Document.open("reviewed.docx", author="Editor")

# List all tracked revisions
revisions = doc.list_revisions()
for r in revisions:
    print(f"ID: {r['id']}, Type: {r['type']}, Author: {r['author']}")

# Accept or reject individual revisions
doc.accept_revision(revision_id=1)
doc.reject_revision(revision_id=2)

# Accept or reject all revisions
doc.accept_all()
# or
doc.reject_all()

doc.save()
doc.close()
```

## Redlining Workflow (Document Review)

For comprehensive document review with tracked changes:

### Step 1: Analyze the document

```bash
# Get readable text with any existing tracked changes
pandoc --track-changes=all contract.docx -o contract.md
```

Review the markdown to understand document structure and identify needed changes.

### Step 2: Plan your changes

Organize changes by section or type:
- Date changes
- Party name updates
- Term modifications
- Clause additions/removals

### Step 3: Implement changes

```python
from docx_edit import Document

doc = Document.open("contract.docx", author="Legal Reviewer")

# Section 2 changes
doc.replace("30 days", "60 days")
doc.replace("January 1, 2024", "March 1, 2024")

# Section 5 changes
doc.delete("and any affiliates")
doc.insert_after("termination.", " Notice must be provided in writing.")

# Add review comments
doc.add_comment("indemnification clause", "Review with counsel")

doc.save("contract-reviewed.docx")
doc.close()
```

### Step 4: Verify changes

```bash
pandoc --track-changes=all contract-reviewed.docx -o verification.md
```

Check that all changes appear correctly in the output.

## Converting Documents to Images

To visually analyze Word documents, convert them to images:

```bash
# Step 1: Convert DOCX to PDF
soffice --headless --convert-to pdf document.docx

# Step 2: Convert PDF pages to JPEG images
pdftoppm -jpeg -r 150 document.pdf page
# Creates: page-1.jpg, page-2.jpg, etc.
```

Options for pdftoppm:
- `-r 150`: Resolution in DPI (adjust for quality/size)
- `-jpeg` or `-png`: Output format
- `-f N`: First page to convert
- `-l N`: Last page to convert

## Code Style Guidelines

When generating code for DOCX operations:
- Write concise code
- Avoid verbose variable names and redundant operations
- Avoid unnecessary print statements

## Dependencies

Required dependencies (install if not available):

- **docx_edit**: `pip install docx-edit` (for editing documents)
- **pandoc**: `sudo apt-get install pandoc` (for text extraction)
- **docx**: `npm install -g docx` (for creating new documents)
- **LibreOffice**: `sudo apt-get install libreoffice` (for PDF conversion)
- **Poppler**: `sudo apt-get install poppler-utils` (for pdftoppm)
