# 📊 Indian Tech Job Market 2026 — Excel Analytics Dashboard

An end-to-end Excel data analytics project that cleans, analyzes, and visualizes **23,201 real tech job postings** scraped from naukri.com, covering Data Science, Data Analytics, Business Analysis, Machine Learning, Data Engineering, and Python Developer roles across India.

Built entirely in **Microsoft Excel (2016 compatible)** — a full BI-style dashboard with live formulas, an interactive role/city filter, data-bar visualizations, and native charts. No add-ins, no Power Query, no pivot tables required.

![Dashboard Overview](images/dashboard-overview-1.png)
![Dashboard Charts](images/dashboard-overview-2.png)

---

## 🎯 Project Objective

To analyze the current Indian tech hiring market and answer questions like:

- Which roles are companies hiring the most for right now?
- Which cities have the most job opportunities?
- What skills actually show up most often in job descriptions?
- How do salaries vary by role and experience level?
- What's the real split between remote, hybrid, and on-site jobs?

---

## ✨ What Makes This "Advanced" (Still Pure Excel 2016)

| Feature | How it's built |
|---|---|
| **🔎 Interactive Explorer** | Two dropdowns (Role, City) built with Data Validation lists. Four KPI tiles next to them recalculate live using `COUNTIFS`/`AVERAGEIFS` with wildcard "All" logic — filtering behavior without needing a single pivot table or slicer. |
| **Doughnut charts with per-slice coloring** | Work Mode, Skill Domain, and Experience Tier are visualized as clean doughnut charts with a custom color per segment and % data labels. |
| **Data-bar heatmaps on every summary table** | Job posting counts in Role/City/Other Summary sheets use in-cell conditional-formatting data bars — instant visual ranking, no charts needed. |
| **Data labels directly on every bar chart** | No need to hover or cross-reference a table — every chart shows its numbers inline. |
| **Custom KPI tile design** | Color-coded, bordered KPI cards (not default Excel styling) built entirely from cell fills, merges, and borders. |
| **Zero-gridline, branded chart styling** | Every chart has its gridlines removed, a consistent navy/teal/orange palette, and a styled title — built through openpyxl's chart XML, not Excel's default theme. |

---

## 🗂️ Project Structure

```
indian-tech-job-market-2026-excel-dashboard/
│
├── README.md                                    → you are here
├── data/
│   └── indian_tech_jobs_2026.csv                 → raw source dataset (23,201 rows, 32 columns)
├── excel/
│   └── Indian_Tech_Job_Market_2026_Dashboard.xlsx → the full Excel project (data + formulas + dashboard)
├── images/
│   └── *.png                                      → dashboard & sheet screenshots used in this README
└── docs/
    └── PROJECT_SUMMARY.md                         → methodology, data cleaning log, and key insights
```

---

## 🛠️ Tools & Skills Used

| Tool | Purpose |
|---|---|
| **Microsoft Excel** (2016 compatible) | Data cleaning, formulas, pivoting logic, dashboarding |
| **Excel Tables** (structured references) | `RawData[column_name]` style formulas that stay correct as data changes |
| **COUNTIFS / SUMIFS / AVERAGEIFS / IFERROR** | Live aggregation formulas — no manual pivot refresh needed |
| **Native Excel Charts** | Bar, horizontal bar, and pie charts wired directly to the summary tables |
| **Python (pandas)** | One-time data cleaning & text-mining pass (documented in `docs/PROJECT_SUMMARY.md`) before loading into Excel |

---

## 📁 Inside the Workbook

The `.xlsx` file has **7 visible sheets** (+ 1 hidden helper sheet for dropdown lists):

| Sheet | What's in it |
|---|---|
| **Dashboard** | Banner, 6 KPI cards, Interactive Explorer (role/city filter + 4 live mini-KPIs), and 7 charts (2 bar, 3 doughnut, 1 horizontal bar, 1 column) |
| **Read Me** | In-workbook notes on data cleaning, how the Interactive Explorer works, and performance tips |
| **Raw Data** | All 23,201 cleaned job postings as a proper Excel Table (`RawData`) |
| **Role Summary** | Postings, avg experience, avg salary, avg rating — by role, with data-bar visualization (formula-driven) |
| **City Summary** | Postings & avg salary — by city, top 15, with data bars (formula-driven) |
| **Other Summary** | Skill domain, experience tier, work mode, company size, salary tier, posting recency — all with data bars (formula-driven) |
| **Top Skills & Companies** | Top 20 in-demand skills (text-mined) and top 20 hiring companies, with data bars |

**Every summary number outside "Top Skills & Companies" — including every number behind the Interactive Explorer — is a live formula** referencing the `RawData` table. Edit a row in Raw Data and the whole dashboard, filters included, recalculates.

---

## 📈 Dashboard Preview

### Role Summary (formula-driven, with data bars)
![Role Summary](images/role-summary.png)

### City Summary (formula-driven, with data bars)
![City Summary](images/city-summary.png)

### Other Summary — Skill Domain, Experience, Work Mode, Company Size
![Other Summary](images/other-summary.png)

### Top Skills & Top Hiring Companies
![Top Skills and Companies](images/top-skills-companies.png)

---

## 🔑 Key Insights

- **Data Scientist** is the most in-demand role (6,455 postings, 27.8% of the market), followed by Data Analyst and Business Analyst.
- **Mumbai, Bangalore, Chennai, Pune, and Noida** are the top 5 hiring cities, together accounting for ~60% of all postings.
- **Python, Data Analysis, Machine Learning, and SQL** are the four most frequently mentioned skills across all job descriptions.
- **81% of jobs are On-site** — remote roles make up only 9.3% of the market.
- Only **~12% of postings disclose salary**; among those that do, the average is ₹15.3 LPA, with Data Engineer and Data Scientist roles paying the highest on average.
- **Mid-level (3–5 Yrs)** experience is the single largest bracket (43.6% of postings) — the market is not primarily fresher-friendly.

*(Full methodology and a longer insights write-up is in [`docs/PROJECT_SUMMARY.md`](docs/PROJECT_SUMMARY.md).)*

---

## 🚀 How to Use This Project

1. Download `excel/Indian_Tech_Job_Market_2026_Dashboard.xlsx`.
2. Open it in Excel 2016 or later — the Dashboard sheet is the default landing tab.
3. Try the **Interactive Explorer**: pick a Role and City from the two dropdowns near the top — the 4 tiles next to them (Matching Postings, Avg Salary, Avg Experience, % Remote) update instantly.
4. To explore the full dataset yourself, click into the **Raw Data** sheet and use the Table filter arrows (faster than sorting/filtering the whole sheet).
5. To see how a number is calculated, click any cell in the Role/City/Other Summary sheets — you'll see the live `COUNTIFS`/`SUMIFS`/`AVERAGEIFS` formula.
6. Want to re-run the analysis on updated data? Replace the rows inside the `RawData` table (keep the same columns) — every formula, chart, and the Interactive Explorer all update automatically.

### 💻 Performance Note
This project was built and tested to run comfortably on a modest machine (8 GB RAM, HDD). If you're on similar hardware: keep only this workbook open, and switch **Formulas → Calculation Options → Manual** if you're doing heavy edits to Raw Data, then press `F9` to refresh.

---

## 📌 Data Source & Disclaimer

- Source: job listings scraped from **naukri.com**, single snapshot dated **10-Jun-2025**.
- This is **not live data** — it reflects a point-in-time snapshot, not current openings.
- Salary figures reflect only the ~12% of postings that disclosed a salary range and may not be representative of the full market.

---

## 👤 Author

**Shrikant Kshitij**
B.Tech CSE · Data Analytics Trainee (SQL, Python, Power BI, Excel) · Ex-Linux Administrator

