# Security policy

sheetpeek renders workbooks, which may come from anyone, to HTML, Markdown and text, and serves
them from a small local server. A security problem here is a workbook whose output carries markup
or script it should have escaped (text, a link, a color, a sheet name), or that crashes sheetpeek,
hangs it, or makes it take time or memory out of proportion to the file, past the limits it was
given.

## Reporting a vulnerability

Please report security problems privately, through GitHub's private vulnerability reporting:
open the repository's **Security** tab and choose **Report a vulnerability**
(https://github.com/sheetmill/sheetpeek/security/advisories/new). Do not open a public issue.

Useful in a report: the version, a workbook (a crafted one is fine) or the steps that reproduce
the problem, the output it gives (the HTML, Markdown or text), and what happens (output that is not
escaped, a crash, a hang, time or memory out of proportion, a limit that does not hold).

A problem in what sheetpeek is given belongs in the package that gives it, reported the same way:
reading a workbook, to xlio (https://github.com/sheetmill/xlio/security/advisories/new);
computing its formulas, to xlrecalc (https://github.com/sheetmill/xlrecalc/security/advisories/new);
formatting a value with a number format, to xlformat
(https://github.com/sheetmill/xlformat/security/advisories/new). If you are not sure which, report
it here.

## Supported versions

sheetpeek is in alpha. Fixes go into the next release; only the latest release receives fixes.

## What happens after a report

The report is acknowledged, reproduced and fixed in the next release, whose changelog states the
fix in its Security section. Reporters are credited there unless they prefer not to be.
