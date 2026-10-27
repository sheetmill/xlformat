"""``python -m capture <suite> [--diff-2023]``: build the suite's workbook, capture it from Excel
for Mac and write ``tests/captures/<suite>.tsv`` (see ``capture/capture.py`` and CAPTURING.md)."""

import sys

from capture.capture import main

main(sys.argv[1:])
