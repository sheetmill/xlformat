# Contributing

## Issues and bug reports

Bug reports are welcome. The most useful report gives the value, the format string,
what Excel shows and what `xlformat` shows, along with your Excel version and platform.
A screenshot of the cell helps when the difference is a space or a fill character.

## Pull requests

The project is just getting started, so pull requests aren't being accepted yet while
the test-capture workflow, release process and review practices settle. This will be
revisited after the first releases. In the meantime, if you have a change in mind,
please open an issue describing it.

When pull requests open, these rules will apply:

- AI-assisted contributions must be disclosed in the pull request description.
- The contributor must understand every line of the change and be able to explain it when asked.
- Undisclosed or unexplained AI-generated pull requests may be closed without review.

## Setting up

Python 3.11 or later. Either route installs the package in editable mode with the development
tools; add `--group docs` for the documentation tools.

```bash
uv sync --group docs
```

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e . --group dev --group docs      # needs pip 25.1 or later
```

`uv.lock` pins the versions CI tests with, and `uv sync` installs exactly those versions.
pip and other tools ignore it and resolve the latest versions the dependency groups in
`pyproject.toml` allow, which is normally fine but is not what CI runs.

## Checks to run before proposing a change

All of these must pass; CI runs the same set on Python 3.11 to 3.14. Prefix each with `uv run`
on the uv route.

```bash
pytest                                   # unit tests, every captured row, the doc examples
ruff check .
ruff format --check .
mypy
sphinx-build -n -W docs docs/_build/html # the strict documentation build
```

`pytest` checks every row captured from Excel in a few seconds; `pytest --tier=fast` runs a
per-format sample. The examples in `README.md` and `docs/*.md` run as doctests, so a change of
behavior shows up there too.

## The expected outputs

The TSV files under `tests/captures/` are ground truth: each row is what Excel for Mac displayed
for a value and a format string. They are never edited by hand. A suite is regenerated only by
capturing it again from Excel for Mac with the scripts in `capture/` (see `CAPTURING.md`), and a
regeneration that changes any expected text is reviewed as a change of evidence. If the code
and a row disagree, the code is wrong unless the row's input can be shown to differ from what
the test assumes.

## Commit messages

A commit made with the help of an AI tool ends with the trailer

```
Assisted-by: <tool>:<model ID>
```

as its last line, after a blank line, naming the tool and the exact model, for example
`Assisted-by: Claude Code:claude-opus-5-5`. One trailer per tool and model used.

## How this project's own commits are made

Most of the code is written with Claude Code, reviewed before each commit and tested against
expected values captured from real Excel (see `CAPTURING.md`). Every such commit carries the git
trailer `Assisted-by: Claude Code:<model ID>`, so the history records which model took part.
