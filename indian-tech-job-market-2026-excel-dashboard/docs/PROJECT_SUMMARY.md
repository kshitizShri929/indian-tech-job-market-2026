# Project Summary — Indian Tech Job Market 2026

## 1. Overview

| | |
|---|---|
| **Dataset** | 23,201 job postings, 32 raw columns |
| **Source** | naukri.com (scraped snapshot) |
| **Snapshot date** | 10-Jun-2025 |
| **Roles covered** | Data Scientist, Data Analyst, Business Analyst, Machine Learning Engineer, Data Engineer, Python Developer |
| **Tools used** | Python (pandas) for one-time cleaning & text-mining, Microsoft Excel for the live dashboard |
| **Final deliverable** | `excel/Indian_Tech_Job_Market_2026_Dashboard.xlsx` — 7 sheets, 69,000+ live formulas, 7 charts |

This document explains **how** the workbook was built: the data quality issues found in the raw scrape, exactly how each was handled, and the key findings once the data was clean.

---

## 2. Data Cleaning Methodology

The raw CSV was inspected column by column before anything was loaded into Excel. Three real data quality issues were found and fixed; everything else in the source was already usable as-is.

### 2.1 Experience field scraping glitch (90 rows)

**Problem:** For 90 rows (~0.4% of the dataset), the `experience_raw` column contained a **date fragment** (e.g. `"13 Jun"`, `"11 Jun - 10 Jul"`) instead of a year range like `"5-10 Yrs"`. This corrupted the derived `experience_min_yrs` / `experience_max_yrs` columns — in several cases `experience_min_yrs` ended up **greater than** `experience_max_yrs` (e.g. min=13, max=8), which is logically impossible.

**Detection:** Flagged every row where `experience_min_yrs > experience_max_yrs`.

**Fix:** For those 90 rows, `experience_min_yrs` and `experience_max_yrs` were reset to blank, `experience_tier` was reset to `"Data Not Available"`, and a new `data_quality_flag` column marks them as `"Experience Corrected"`. Nothing was guessed or estimated — the honest choice was to mark these as unknown rather than invent a number.

### 2.2 City name fragmentation (198 → 46 variants)

**Problem:** The `primary_city` column had **198 unique values** for what were really far fewer actual cities, because sub-locality tags were baked into the city name — e.g. `"Mumbai(Mahalaxmi)"`, `"Mumbai(Mulund)"`, `"Pune(Baner)"`, `"Pune(Hinjewadi Phase 1)"`, `"Mumbai(Goregaon East +3)"` were all being counted as separate cities from plain `"Mumbai"` / `"Pune"`. This made city-level analysis meaningless without cleanup — "Mumbai" alone was undercounted by ~70 postings hidden inside suburb-tagged variants.

**Fix:** A cleaning function stripped anything in parentheses and anything after a `+` (used for "and N more locations"), then trimmed whitespace. This consolidated 198 raw values into **46 clean city names**, all stored in a new `city` column.

### 2.3 Salary "Not Disclosed" placeholders (20,433 rows)

**Problem:** For undisclosed salaries, the source data stored `0.0` in `salary_min_lpa` / `salary_max_lpa` / `salary_midpoint_lpa` rather than leaving them blank. Averaging these columns directly (as-is) would silently pull every salary average toward zero, since 88% of the dataset is undisclosed.

**Fix:** Wherever `salary_disclosed = False`, the salary columns were converted to blank (`NaN`) instead of `0`. Every average salary figure in the workbook (`AVERAGEIFS` formulas) is therefore computed **only over the ~12% of postings that actually disclosed a salary**, which is stated explicitly next to every salary figure in the workbook.

### 2.4 Columns dropped from the Excel Raw Data sheet

To keep the file lightweight for Excel 2016 on modest hardware, two large free-text columns were excluded from the workbook's Raw Data sheet (they added no analytical value and significantly bloat file size): `job_description` (long unstructured text) and `scraped_at` (a single constant value for every row, `2025-06-10`). All other columns were retained.

---

## 3. From Cleaned Data to Dashboard

### 3.1 Why formulas, not hardcoded numbers

Every summary table (Role Summary, City Summary, Other Summary) is built with `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS` formulas that reference the `RawData` Excel Table using structured references, e.g.:

```
=COUNTIFS(RawData[role_category], B6)
=IFERROR(AVERAGEIFS(RawData[salary_midpoint_lpa], RawData[role_category], B6, RawData[salary_disclosed], TRUE), 0)
```

This means the workbook isn't a static snapshot report — if you paste in new job postings (keeping the same columns) or edit existing rows in `Raw Data`, every KPI, table, and chart on the Dashboard recalculates automatically. No pivot table refresh, no macro needed.

### 3.2 The one exception: Top Skills

The `skills_required` column stores multiple skills as comma-separated free text inside a single cell (e.g. `"Python, SQL, Machine Learning, Data Visualization"`). Counting how often each individual skill appears across 23,201 such cells isn't something `COUNTIFS` can do directly — it would need array-based text parsing that Excel 2016 (and even modern Excel, reliably) doesn't handle well at this scale.

Instead, this one table was generated with a short Python text-mining pass (splitting each cell on commas, trimming and title-casing each skill, then counting frequency across all rows) and written into the workbook as a **static snapshot**, clearly labeled as such. If the underlying data changes, this table needs to be regenerated — every other table in the workbook does not.

---

### 3.3 Interactive filtering without pivot tables

The Dashboard also includes a small **Interactive Explorer** — two Data Validation dropdowns (Role, City) wired to four KPI tiles via `COUNTIFS`/`AVERAGEIFS` formulas using a wildcard pattern:

```
=COUNTIFS(RawData[role_category], IF($RoleCell="All Roles","*",$RoleCell),
          RawData[city],          IF($CityCell="All Cities","*",$CityCell))
```

When the dropdown is left on "All Roles" / "All Cities", the `IF` resolves to a `"*"` wildcard that matches every row; picking a specific value narrows the formula to that exact match. This gives pivot-table-style interactivity while staying fully compatible with Excel 2016 (no Slicers, no Power Pivot, no macros).

---

## 4. Key Findings

### Market composition
- **Data Scientist** is the single largest role category: 6,455 postings (27.8% of the market) — nearly 1.4x the next closest role (Data Analyst, 20.4%).
- Combined, the three "analyst/scientist" roles (Data Scientist, Data Analyst, Business Analyst) make up **67.6%** of all postings — core "developer" roles like Python Developer are a much smaller slice (6.8%).

### Geography
- **Mumbai (14.4%), Bangalore (12.2%), Chennai (11.7%), Pune (11.6%) and Noida (11.1%)** are the top 5 hiring hubs, together accounting for roughly 61% of all postings.
- **Remote-tagged postings (2,010, 8.7% of total)** rank 7th among "cities" — meaningful, but still a minority compared to any of the top physical hubs.

### Skills
- **Python (18.4% of postings), Data Analysis (16.3%), Machine Learning (15.2%), and SQL (12.7%)** are the four most-requested skills, appearing in roughly 1 in 6 postings or more.
- Traditional "soft"/process skills like **Analytical** (10.5%) and **Business Analysis** (10.3%) rank just below the core technical skills — suggesting most roles expect a blend, not pure technical depth.

### Work mode & experience
- **81% of jobs are On-site**, 10% Hybrid, and just **9.3% Remote** — despite widespread remote-work narratives, the Indian tech hiring market in this snapshot remains overwhelmingly office-based.
- **Mid-level (3–5 Yrs) experience is the largest single bracket at 43.6%** of postings, more than double the Junior (0–2 Yrs, 21.2%) and Senior (6–8 Yrs, 20.7%) brackets combined with Fresher (9.6%). This is a mid-level-heavy market, not an entry-level-heavy one.

### Salary (where disclosed)
- Only **11.9% of postings (2,768 of 23,201)** disclose a salary range at all — any salary conclusion should be read as indicative, not representative of the full market.
- Among postings that do disclose, the **average is ₹15.3 LPA**. By role, **Data Engineer (₹17.8 LPA avg) and Data Scientist (₹18.6 LPA avg)** pay highest on average; **Data Analyst (₹10.6 LPA avg)** is the lowest-paid of the six roles tracked.
- By skill domain, **Cloud & DevOps (₹19.0 LPA avg)** and **AI/ML/DL (₹18.2 LPA avg)** command the highest average pay; **Business Intelligence (₹12.3 LPA avg)** — the largest skill domain by volume (47.9% of postings) — pays the least on average.

### Company landscape
- **6,993 unique companies** are represented across the 23,201 postings — the market is highly fragmented, not dominated by a handful of employers.
- **Tata Consultancy Services (653 postings)** and **Accenture (438 postings)** are, by a clear margin, the two most active hirers in the dataset — together posting more than the next 6 companies on the list combined.

---

## 5. Limitations & Honest Caveats

- **Single snapshot, not live data.** Everything above reflects postings active on 10-Jun-2025; it is not a live feed of current openings.
- **Salary figures are self-selected.** Only ~12% of postings disclose salary, and companies that disclose salary may systematically differ (e.g. more transparent employers, or specific salary bands) from the 88% that don't — so salary insights should be treated as directional, not a market-wide average.
- **Source is naukri.com only.** Postings from other job boards (LinkedIn, Indeed, company career pages) are not represented, so this is a partial view of the market, not the whole market.
- **90 rows (0.4%) have unknown experience** due to the scraping glitch described in Section 2.1 — they're retained in the dataset (for accurate job-count totals) but excluded from any experience-based average.

---

## 6. Possible Next Steps

- Track postings over multiple scrape dates to turn this into a trend analysis (growth/decline by role, city, or skill over time) rather than a single snapshot.
- Cross-reference `company_rating` against `salary_midpoint_lpa` to see whether higher-rated companies pay more, controlling for role.
- Extend the Top Skills text-mining step to compute skill *co-occurrence* (which skills are most often requested together), which would need a small Python script beyond what Excel formulas alone can do.
