"""``python -m capture <suite> [<suite> ...] [--diff-2023] [--out DIR]``: build each suite's
workbook, capture it from Excel for Mac and write ``tests/captures/<suite>.tsv`` (or into DIR;
see ``capture/capture.py`` and CAPTURING.md)."""

import sys

from capture.capture import main

main(sys.argv[1:])
