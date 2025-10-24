## 📘 Overview
This project automates the extraction of expert profiles from [Alternatives.org.uk](https://www.alternatives.org.uk/experts) — capturing names, specialization tags, and biographies in the exact order they appear on the website. The scraper delivers a clean, well-structured Excel report designed for professional use, analytics, or client presentation.

## 🎯 Objectives
- Automate expert data collection with accuracy and order preservation.
- Replace missing or blank values with “N/A”.
- Eliminate duplicates and inconsistent spacing.
- Deliver a fully formatted, presentation-ready Excel dataset.

## ⚙️ Technology Stack
- **Python** – Core scripting
- **Playwright (async/await)** – Dynamic content scraping
- **BeautifulSoup4** – HTML parsing
- **Pandas** – Data structuring and cleaning
- **OpenPyXL** – Excel styling and automation
- **Tenacity** – Retry logic for stability

## 🧩 Key Features
- Maintains exact sequential order of profiles.
- Auto-retries on temporary network or timeout issues.
- Professionally styled Excel output:
  - Dark blue header (#1F4E78), white bold centered text
  - Alternating row shading (#F5F5F5)
  - “N/A” cells shown in light grey italic
  - Auto-fit column widths, frozen header, and auto-filter
- Metadata footer:
📊 Sourced from (https://www.alternatives.org.uk/experts
)
Generated on: [timestamp]

## 🚀 How It Works
1. Launches Playwright to load and render the dynamic expert page.
2. Uses BeautifulSoup to extract each expert’s name, tags, and biography.
3. Cleans, normalizes, and validates all records.
4. Saves the final structured dataset as **speakers.xlsx** with full professional styling.

## 📦 Output
A refined Excel file containing all expert profiles with consistent formatting, ready for analytics, reporting, or portfolio demonstration.

---

**Author:** Onyekachi Ejimofor  
**Purpose:** Demonstrate professional web scraping, data cleaning, and reporting automation for business and data-driven applications.
