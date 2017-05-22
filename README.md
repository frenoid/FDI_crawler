# FDI_crawler
to extract projects, signals and company reports for fDi markets

Files
=====
1) company_report.py
This file downloads Company Report from the Companies page in fDi markets, one-by-one. It organizes them by page number.

2) project_report.py
This file downloads listed Projects from the fDi projects page. It downloads them by country of origin, constrained by
date-range to not exceed the export limit.

3) signal_report.py
This file downloads listed Signals from the fDi singals page. it downloads them by country of origin, constrained by date-range
to not exceed the export limit.

4) zero_signal_file.xlsx
This file is used when signal_report.py encounters an instance where a country has 0 signals for a given date-range. This file
stands in for that particular download.
