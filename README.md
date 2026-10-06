# 🏠 Rent Portfolio Analyzer

> Python data pipeline that automatically detects revenue leakage 
> across residential portfolios — built by a property manager, 
> for property managers.

![Python](https://img.shields.io/badge/Python-3.8+-blue)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-green)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20Reports-orange)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

---

## 📌 The Problem

Managing 350+ residential units means thousands of data points — 
rent rates, lease dates, vacancy status, market comps. Doing this 
manually means revenue gaps go undetected for months.

This tool fixes that. Run it once. Know exactly where your money is.

---

## 🚀 What It Does

Ingests tenant data exported from property management software 
(RentManager, AppFolio, Buildium) and automatically:

- 🔴 **Identifies vacant units** generating zero income
- 🟡 **Flags below-market leases** where tenants pay under current rate
- 💸 **Calculates total monthly and annual revenue leakage**
- 📊 **Generates a professional 4-tab Excel report** instantly

---

## 📊 Output — 4 Excel Sheets Generated Automatically

| Sheet | Contents |
|---|---|
| 📊 Executive Summary | 8 KPI boxes — vacancies, leakage, occupancy rate, current vs potential income |
| 🏠 All Units | Full portfolio sorted by priority — red, yellow, green status |
| 🔴 Vacancies | Every vacant unit with monthly/annual loss and days vacant |
| 🟡 Below Market Rate | Every under-market tenant sorted by largest gap first |

---

## 💡 Real World Impact

> "On first run against a 350-unit portfolio this tool identified 
> revenue gaps that had gone undetected for months. The entire 
> analysis runs in under 3 seconds."

---

## 🛠️ Installation
```bash
# Clone the repository
git clone https://github.com/michael-chaves-dev/rent-portfolio-analyzer

# Navigate to project directory
cd rent-portfolio-analyzer

# Install dependencies
pip install pandas openpyxl
```

---

## 📋 Usage

### Option 1 — Run with your own data
```bash
python rent_analyzer.py your_tenant_export.csv
```

### Option 2 — Run demo with sample data
```bash
python rent_analyzer.py
```
*No file needed — generates a realistic 50-unit demo automatically*

---

## 📥 How To Export From RentManager

1. Log into RentManager Express
2. Go to **Reports → Rent Roll**
3. Click **Export to Excel**
4. Save the file
5. Run: `python rent_analyzer.py your_file.xlsx`

---

## 📁 Required Data Columns

| Column | Description |
|---|---|
| Unit | Unit number or identifier |
| Tenant Name | Full tenant name |
| Lease Start | Lease start date |
| Lease End | Lease end date |
| Market Rent | Current correct/market rate |
| Actual Rent | What tenant currently pays |
| Status | Occupied / Vacant |

> Column names are configurable in the COLUMN_MAP section 
> of the script to match your specific export format.

---

## 🧰 Tech Stack

- **Python 3.8+**
- **Pandas** — Data ingestion, cleaning, and analysis
- **OpenPyXL** — Professional Excel report generation
- **datetime** — Lease date calculations

---

## 🗺️ Roadmap

- [x] CSV and Excel import support
- [x] 4-tab professional Excel output
- [x] Vacancy detection and reporting
- [x] Below-market rate detection
- [x] Revenue leakage calculation
- [ ] Web interface — upload CSV, download report (RentRadar v2)
- [ ] Direct RentManager API integration
- [ ] Live market rent comparison via Zillow API
- [ ] Email report scheduling and automation

---

## 👤 About The Author

Built by **Michael Chaves** — property manager, business owner, 
and Python developer in training.

This tool was born from a real operational need managing 350+ 
residential units. Every feature solves a problem I lived personally.

- 🔒 Google Cybersecurity Certificate
- 🛡️ TCM Security Academy (In Progress)
- 🤖 DeepLearning.AI (In Progress)
- 🌍 Bilingual: English / Spanish / Portuguese
- 💼 [GitHub Portfolio](https://github.com/michael-chaves-dev)

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.

---

## 🤝 Contributing

Pull requests welcome. If you manage a residential portfolio 
and want a feature added — open an issue. This tool is built 
for real operators by a real operator.

---

*⭐ If this tool saves you time or money, drop a star on the repo!*



# rent-portfolio-analyzer
Python data pipeline that detects revenue leakage across residential portfolios — identifies vacant units and below-market leases, generates professional Excel reports

<!-- SAS-IP-FOOTER-v1 -->
---
**Subtract Architect Studios™**  
Copyright © 2026 Michael F. Chaves. All rights reserved in original Subtract Architect Studios materials except as expressly licensed. See [IP_NOTICE.md](./IP_NOTICE.md). Existing open-source and third-party licenses remain controlling for materials they cover.
