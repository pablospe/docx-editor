# Contributing to `docx-editor`

Contributions are welcome, and they are greatly appreciated!
Every little bit helps, and credit will always be given.

You can contribute in many ways:

# Types of Contributions

## Report Bugs

Report bugs at https://github.com/pablospe/docx-editor/issues

If you are reporting a bug, please include:

- Your operating system name and version.
- Any details about your local setup that might be helpful in troubleshooting.
- Detailed steps to reproduce the bug.

## Fix Bugs

Look through the GitHub issues for bugs.
Anything tagged with "bug" and "help wanted" is open to whoever wants to implement a fix for it.

## Implement Features

Look through the GitHub issues for features.
Anything tagged with "enhancement" and "help wanted" is open to whoever wants to implement it.

## Write Documentation

docx-editor could always use more documentation, whether as part of the official docs, in docstrings, or even on the web in blog posts, articles, and such.

## Submit Feedback

The best way to send feedback is to file an issue at https://github.com/pablospe/docx-editor/issues.

If you are proposing a new feature:

- Explain in detail how it would work.
- Keep the scope as narrow as possible, to make it easier to implement.
- Remember that this is a volunteer-driven project, and that contributions
  are welcome :)

# Get Started!

Ready to contribute? Here's how to set up `docx-editor` for local development.
Please note this documentation assumes you already have `uv` and `Git` installed and ready to go.

1. Fork the `docx-editor` repo on GitHub.

2. Clone your fork locally:

```bash
cd <directory_in_which_repo_should_be_created>
git clone git@github.com:YOUR_NAME/docx-editor.git
```

3. Now we need to install the environment. Navigate into the directory

```bash
cd docx-editor
```

Then, install and activate the environment with:

```bash
uv sync
```

4. Install pre-commit to run linters/formatters at commit time:

```bash
uv run pre-commit install
```

5. Create a branch for local development:

```bash
git checkout -b name-of-your-bugfix-or-feature
```

Now you can make your changes locally.

6. Don't forget to add test cases for your added functionality to the `tests` directory.

7. When you're done making changes, check that your changes pass the formatting tests.

```bash
make check
```

Now, validate that all unit tests are passing:

```bash
make test
```

9. Before raising a pull request you should also run tox.
   This will run the tests across different versions of Python:

```bash
tox
```

This requires you to have multiple versions of python installed.
This step is also triggered in the CI/CD pipeline, so you could also choose to skip this step locally.

10. Commit your changes and push your branch to GitHub:

```bash
git add .
git commit -m "Your detailed description of your changes."
git push origin name-of-your-bugfix-or-feature
```

11. Submit a pull request through the GitHub website.

# Pull Request Guidelines

Before you submit a pull request, check that it meets these guidelines:

1. The pull request should include tests.

2. If the pull request adds functionality, the docs should be updated.
   Put your new functionality into a function with a docstring, and add the feature to the list in `README.md`.

# Releasing a New Version

This project uses GitHub Releases to trigger automated publishing to PyPI and docs deployment.

## Steps

1. **Update the version** in both files on `main`:

   - `pyproject.toml` → `version = "X.Y.Z"`
   - `.claude-plugin/plugin.json` → `"version": "X.Y.Z"`

2. **Commit and push** the version bump:

   ```bash
   git add pyproject.toml .claude-plugin/plugin.json
   git commit -m "bump version to X.Y.Z"
   git push origin main
   ```

3. **Create a GitHub Release**:

   - Go to [Releases](https://github.com/pablospe/docx-editor/releases/new)
   - Create a new tag matching the version: `X.Y.Z` (e.g., `0.3.0`)
   - Set the target branch to `main`
   - Add release notes (use "Generate release notes" for a changelog)
   - Click **Publish release**

4. **Automated CI** (`.github/workflows/on-release-main.yml`) will:

   - Update `pyproject.toml` version from the release tag
   - Build the package with `uv build`
   - Publish to [PyPI](https://pypi.org/project/docx-editor/) via trusted publishing
   - Deploy documentation to GitHub Pages with `mkdocs gh-deploy`

## Notes

- The release tag **must** match the version format (e.g., `0.3.0`, no `v` prefix)
- PyPI publishing uses [trusted publishing](https://docs.pypi.org/trusted-publishers/) (no API tokens needed)
- If you need to build and publish manually, you can use `make build-and-publish`
