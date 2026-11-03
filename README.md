# Scopus Research Insights — Faculty Publications Dashboard

A **ready-to-host GitHub Pages dashboard** for the Excel workbooks produced by the Scopus / Google Scholar / CiteScore reporting Python program. No backend, deployment token, database, Scopus API key, or server is needed for the website.

## Features

- Upload one or multiple Scopus reporting `.xlsx` files; select a year or combine years
- Search authors by typing their names; click a profile
- Author totals, Q1 / Q2 / Q3 / Q4 and N/A, monthly Jan–Dec publication chart, journal CiteScore
- Institution overview with co-authored papers deduplicated
- Publication list with title, DOI, Scopus URL, journal, quartile, CiteScore, date and filters
- Author ranking, source/journal comparison, downloadable CSV files
- Clearly distinguish the **publication month** from the actual Scopus indexing month
- All uploaded Excel data stays within the visitor's browser (no API calls containing author report data)


## Data definitions and cautions

- **Year / month:** Scopus publication cover date, **not** the date the paper first appeared in Scopus.
- **Q1–Q4:** Best CiteScore percentile among source subject categories, not paper quality, not SJR quartile, not JCR quartile.
- **Metric year:** Latest *completed* CiteScore year in the imported report, which can differ from the publication year.
- **N/A:** Missing CiteScore or percentile. Never silently classify N/A as Q4.
- **Institution totals:** Unique papers deduplicated across faculty by paper key. **Author rankings:** Count the paper under each author; author totals may exceed the institution's unique-paper count.
- **Retrieval errors:** Check the Python reporting program's `Errors` worksheet if an author has incomplete API results. The dashboard warns when errors are present.
- **Overlapping report files:** Re-importing the same paper does not double-count an author-paper record with the same year and unique paper key.

## How to update the dashboard

Run your existing Scopus Python reporter, export an updated workbook, open your GitHub Pages dashboard, and upload it. You do not need to edit the website code for each new reporting period.

## Troubleshooting

- **Blank charts:** Verify the workbook contains `Scopus Records`, not just author URLs.
- **Excel reader unavailable:** Check your network access to `https://cdn.sheetjs.com/`.
- **Q1–Q4 all N/A:** Inspect the output `Scopus Records` and `Journal Metrics` for completed CiteScore percentile availability; it is not possible to infer quartile from CiteScore alone.
- **Incorrect indexing month:** The Python input does not contain Scopus first-indexed timestamps. The dashboard can only display publication months.
- **Duplicate author names:** Each author is resolved by Scopus Author ID where available.

## Technical stack

Plain HTML / CSS / JavaScript and browser-only SheetJS Excel parsing. No npm, API secrets, GitHub Actions, cloud database, or external frontend framework. License: MIT for the provided dashboard code; SheetJS remains under its own separate license and is loaded from its CDN.


