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

## Which Excel file to upload?

Use the **output report** from `Scopus_Scholar_Monthly_Report_WITH_CITESCORE_NO_PDF.py`, **not** the two-column input authors list. The report should contain:

1. `Monthly Summary` — author names and Scopus Author IDs
2. `Scopus Records` — one row per author-paper with title, date, journal and quartile
3. Optional: `Journal Metrics`, `Run Info` and `Errors`

Both variants of the reporter (with or without PDF columns) are supported. This web app does **not** retrieve new publications. Generate a fresh Excel report using your Python program whenever you want new data.

## Run locally

You may open `index.html` directly in your browser while connected to the internet, or run a local web server:

```bash
python -m http.server 8000
```

Open `http://localhost:8000/`. Click **Try demo** to see fictional data, then **Upload Excel** to load your real report. The Excel-reading JavaScript (`SheetJS`, version 0.20.3) is loaded from its official CDN and requires internet access.

## Publish on GitHub Pages (no Python server)

1. On GitHub, create a new repository, e.g. `scopus-research-dashboard`.
2. Upload **all** website files to the repository root: `index.html`, `styles.css`, `dashboard-core.js`, `app.js`, `.nojekyll`, and this README. The `data/` folder is optional. You do **not** need to upload your actual report or your Scopus API key.
3. Go to **Settings → Pages → Build and deployment**.
4. Select **Deploy from a branch**, choose `main` and `/(root)`, then **Save**.
5. Your hosted dashboard will be available at a URL shaped like `https://YOUR-USERNAME.github.io/scopus-research-dashboard/` after publication. Open it and upload the Excel report from your own computer.

GitHub Pages serves the static site. It does not run Python or fetch new Scopus records on its own.

Official references:
- https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site
- https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site
- https://docs.sheetjs.com/docs/getting-started/installation/standalone/

## Privacy and access choices

By default, **do not commit the faculty Excel report** to a public GitHub repository. Visitors upload their local workbook; it is processed solely in their browser. This also means each visitor must supply the report themselves. The website is not access-controlled: anyone with the Pages URL can use the interface, but cannot see another visitor's report unless that report has been separately published.

**Optional automatic public dashboard:** If you explicitly want all visitors to see the same dataset without uploading Excel, review and approve a sanitized report first, then commit it to your GitHub repository with the exact path `data/report.xlsx`. The site automatically detects and loads that file, and it becomes publicly downloadable. **Never publish a report that must remain private.** If `data/report.xlsx` does not exist, the website starts in private upload mode. The provided package deliberately contains no public report file.

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

## Correct Scopus paper links

When reading an Excel report, `dashboard-core.js` converts every available `Scopus EID` into the corresponding web URL:

`https://www.scopus.com/inward/record.uri?partnerID=HzOxMe3b&scp=<SCOPUS_DOCUMENT_ID>&origin=inward`

This fixes older reports whose URLs pointed at Elsevier XML APIs or `record/display.uri`. No fresh Scopus fetch is required. Every paper has its own numeric Scopus document ID; its DOI link is unchanged. Viewing Scopus may require institutional login.
