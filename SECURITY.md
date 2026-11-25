# Security policy

xlformat formats values with Excel number-format strings and measures text; it opens no workbook,
reads no file a caller passes it and writes no markup. A security problem here is a value and a
format string (or a font and a width) that crash it, hang it, or make it take time or memory out of
proportion to the input.

## Reporting a vulnerability

Please report security problems privately, through GitHub's private vulnerability reporting:
open the repository's **Security** tab and choose **Report a vulnerability**
(https://github.com/sheetmill/xlformat/security/advisories/new). Do not open a public issue.

Useful in a report: the version, a value and format string that reproduce the problem (with the
font, column width and platform, if they matter), and what happens (a crash, a hang, time or memory
out of proportion).

A problem in one of the other sheetmill packages belongs in that package's repository, reported
the same way: sheetpeek, which renders workbooks
(https://github.com/sheetmill/sheetpeek/security/advisories/new); xlrecalc, which evaluates
formulas (https://github.com/sheetmill/xlrecalc/security/advisories/new); and xlio, which reads
xlsx and CSV files (https://github.com/sheetmill/xlio/security/advisories/new).

## Supported versions

xlformat is in alpha. Fixes go into the next release; only the latest release receives fixes.

## What happens after a report

The report is acknowledged, reproduced and fixed in the next release, whose changelog states the
fix in its Security section. Reporters are credited there unless they prefer not to be.
