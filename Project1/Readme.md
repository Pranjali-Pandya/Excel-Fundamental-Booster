# 📊 Excel Data Analysis & Advanced Functions Project

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)
![Functions](https://img.shields.io/badge/Formulas-Dynamic_%26_Automated-blue?style=for-the-badge)

A comprehensive Microsoft Excel data analysis workbook demonstrating real-world implementations of logical, lookup, dynamic array, text, date/time, and mathematical functions across three structured business datasets: **Students Grade**, **Sales Data**, and **Employee Data**.

---

## 📑 Table of Contents

- [Repository Structure](#-repository-structure)
- [Workbook Breakdown](#-workbook-breakdown)
  - [1. Project Instructions](#1-project-instructions)
  - [2. Students Grade](#2-students-grade)
  - [3. Sales Data](#3-sales-data)
  - [4. Employee Data](#4-employee-data)
- [Functions & Formulas Summary](#-functions--formulas-summary)
- [Getting Started](#-getting-started)

---

## 📁 Repository Structure

```text
├── Project1_3.xlsx    # Main Excel workbook containing datasets and dynamic formulas
└── README.md          # Project documentation and formula breakdown
```

---

## 📋 Workbook Breakdown

### 1. `Project Instructions`
Serves as the project roadmap, listing all required Excel competencies categorized by function group and mapped to the corresponding analytical tasks across the workbook.

---

### 2. `Students Grade`
Evaluates student academic performance across three subjects (`Math`, `Science`, `English`), enrollment timelines, and automated grade classification.

| Feature / Metric | Cell Range | Excel Functions Used | Description |
| :--- | :--- | :--- | :--- |
| **Total & Average Scores** | `G2:I21` | `SUM`, `AVERAGE` | Computes cumulative marks and percentage averages per student. |
| **Grade Classification** | `K2:K21` | Nested `IF` | Assigns letter grades (`A`, `B`, `C`, `F`) based on percentage thresholds. |
| **Pass / Fail Evaluation** | `L2:L21` | `IF` | Determines pass/fail status dynamically. |
| **Multi-Criteria Status** | `Y2:Y21` | `IF`, `AND`, `OR` | Evaluates combined subject score conditions for honors/eligibility. |
| **Conditional Summaries** | `M2`, `Z2`, `AA2` | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Aggregates student counts and scores filtered by grade and criteria. |
| **Record & Dynamic Lookup** | `AC3`, `AD2` | `VLOOKUP`, `INDIRECT` | Retrieves specific student details and references dynamic cell ranges. |
| **Top Performers Extraction** | `N2:X17` | `FILTER` | Extracts all records where Total Score exceeds `180` (`=FILTER(A2:K21, G2:G21>180)`). |
| **Date & Tenure Metrics** | `AE2:AF21` | `TODAY`, `DATEDIF` | Calculates elapsed days (`=TODAY()-F2`) and full years since enrollment (`DATEDIF`). |

---

### 3. `Sales Data`
Tracks product sales transactions, tax calculations, tiered discounts, regional performance, and financial rounding.

| Feature / Metric | Cell Range | Excel Functions Used | Description |
| :--- | :--- | :--- | :--- |
| **Subtotal, Tax & Discounts** | `E2:I21` | `IF` | Computes gross sales, 18% tax, and conditional discount percentages. |
| **Regional & High-Value Flags** | `L2:L21` | `IF`, `AND`, `OR` | Flags orders meeting multi-condition revenue and region rules. |
| **Sales Aggregations** | `M2` | `SUMIFS`, `COUNTIFS` | Summarizes revenue and order volume by region and category. |
| **Product & Two-Way Lookups** | `J2`, `O2`, `Q5` | `VLOOKUP`, `INDEX`, `MATCH` | Fetches unit prices, product details, and cross-tabulated values. |
| **Modern Lookups & Offset** | `S4`, `T3`, `W3:W4` | `XLOOKUP`, `XMATCH`, `OFFSET` | Performs exact-match lookups, position indexing, and dynamic range shifts. |
| **Financial Rounding** | `Y2:AA21` | `ROUND`, `CEILING`, `FLOOR` | Rounds tax figures to standard precision (`ROUND`), upper multiples (`CEILING`), and lower multiples (`FLOOR`). |

---

### 4. `Employee Data`
Manages HR records, name parsing, text standardization, employee age/tenure tracking, and department lookups.

| Feature / Metric | Cell Range | Excel Functions Used | Description |
| :--- | :--- | :--- | :--- |
| **First Name Extraction** | `C2:C21` | `LEFT`, `FIND` | Dynamically parses first names from full employee names. |
| **Text Case Formatting** | `O2:P21` | `UPPER`, `LOWER` | Standardizes text strings into uppercase and lowercase formats. |
| **Age, Tenure & Days Elapsed** | `G2:I21` | `DATEDIF`, `TODAY` | Computes employee age from DOB, years of service, and total days elapsed. |
| **Employee Lookups** | `N4`, `R4`, `G24` | `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH` | Retrieves employee records and relative row positions dynamically. |

---

## 🛠️ Functions & Formulas Summary

* **Logical:** `IF`, `AND`, `OR`
* **Conditional Aggregation:** `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`
* **Lookup & Reference:** `VLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH`, `INDIRECT`, `OFFSET`
* **Dynamic Arrays:** `FILTER`
* **Text Manipulation:** `LEFT`, `FIND`, `UPPER`, `LOWER`
* **Date & Time:** `TODAY`, `DATEDIF`
* **Mathematical:** `ROUND`, `CEILING`, `FLOOR`

---

## 🚀 Getting Started

1. **Clone or Download** this repository to your local machine:
   ```bash
   git clone https://github.com/<your-username>/<your-repo-name>.git
   ```
2. Open **`Project1_3.xlsx`** using **Microsoft Excel (Microsoft 365 or Excel 2021+)** so that modern dynamic functions (`XLOOKUP`, `XMATCH`, `FILTER`) calculate properly.
3. Ensure automatic calculation is enabled (**Formulas** $\rightarrow$ **Calculation Options** $\rightarrow$ **Automatic**).
