<div align="center">

![Banner](https://capsule-render.vercel.app/api?type=waving&color=0:000000,50:0D0D0D,100:000000&height=180&section=header&text=Excel%20Formula%20Lab&fontSize=55&fontColor=00E5FF&animation=fadeIn&desc=Lookups%2C%20Dynamic%20Arrays%20%26%20Smart%20Formulas%20in%20Microsoft%20Excel&descSize=18&descAlignY=62&fontAlignY=35)

[![Typing SVG](https://readme-typing-svg.demolab.com?font=Fira+Code&size=22&pause=1000&color=F72585&background=000000&center=true&vCenter=true&width=650&lines=IF+%E2%86%92+LOOKUP+%E2%86%92+FILTER+%E2%86%92+OFFSET+%E2%86%92+ANALYZE;Turning+Raw+Cells+into+Real+Insights;Built+with+%F0%9F%92%9C+by+Vaidehi+Vyas)](https://git.io/typing-svg)

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-365-000000?style=for-the-badge&logo=microsoftexcel&logoColor=217346&labelColor=000000)](https://www.microsoft.com/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-IF%20%7C%20SUMIFS%20%7C%20COUNTIFS-000000?style=for-the-badge&logo=databricks&logoColor=9C27B0&labelColor=000000)](#)
[![Lookups](https://img.shields.io/badge/Lookups-VLOOKUP%20%7C%20XLOOKUP%20%7C%20XMATCH-000000?style=for-the-badge&logo=databricks&logoColor=FF6F00&labelColor=000000)](#)
[![Dynamic Arrays](https://img.shields.io/badge/Dynamic%20Arrays-FILTER%20%7C%20OFFSET-000000?style=for-the-badge&logo=cachet&logoColor=4CAF50&labelColor=000000)](#)
[![Status](https://img.shields.io/badge/Status-Completed-000000?style=for-the-badge&logo=checkmarx&logoColor=F72585&labelColor=000000)](#)
[![Made In](https://img.shields.io/badge/Made%20in-India-000000?style=for-the-badge&logo=india&logoColor=FF9933&labelColor=000000)](#)
[![Author](https://img.shields.io/badge/Author-Vaidehi%20Vyas-000000?style=for-the-badge&logo=googlescholar&logoColor=00E5FF&labelColor=000000)](#-author)

<br/>

[![Skills](https://skillicons.dev/icons?i=excel&theme=dark)](#)

<br/>

> *"Spreadsheets hold the data — formulas tell its story."*

<img src="https://raw.githubusercontent.com/andreasbm/readme/master/assets/lines/rainbow.gif" width="100%">

<br/>

<table>
<tr>
<td align="center"><img width="45"src="https://img.icons8.com/fluency/48/microsoft-excel-2019.png"/><br/><b>10</b><br/>Worksheets</td>
<td align="center"><img width="45"src="https://img.icons8.com/fluency/48/product.png"/><br/><b>20</b><br/>Products</td>
<td align="center"><img width="45"src="https://img.icons8.com/fluency/48/student-male.png"/><br/><b>20</b><br/>Students</td>
<td align="center"><img width="45"src="https://img.icons8.com/fluency/48/conference-call.png"/><br/><b>20</b><br/>Employees</td>
<td align="center"><img width="45"src="https://img.icons8.com/fluency/48/source-code.png"/><br/><b>20</b><br/>Functions</td>
</tr>
</table>

</div>

---

## <span style="color:#00E5FF">📋 Table of Contents</span>

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🛒 Part A — Sales Data: Logic, Discounts & Lookups](#-part-a--sales-data-logic-discounts--lookups)
- [🔎 Part B — Modern Lookups: XLOOKUP & XMATCH](#-part-b--modern-lookups-xlookup--xmatch)
- [📅 Part C — Date & Time Functions](#-part-c--date--time-functions)
- [🧮 Part D — Math & Rounding Functions](#-part-d--math--rounding-functions)
- [🧩 Part E — Dynamic Functions: INDIRECT, OFFSET & FILTER](#-part-e--dynamic-functions-indirect-offset--filter)
- [🎓 Part F — Students Grade & Conditional Aggregates](#-part-f--students-grade--conditional-aggregates)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [⚠️ Known Quirks & Fixes](#️-known-quirks--fixes)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## <span style="color:#217346">📌 Overview</span>

**Excel Formula Lab** is a hands-on **Microsoft Excel** workbook (`Excel_p1.xlsx`) built around one central idea: a spreadsheet only becomes powerful once its cells can **look things up, make decisions, and update themselves**. Across ten themed worksheets, the workbook walks through the formulas used most often in real reporting and analysis work.

This project is designed to:
- Build **nested `IF` logic** for discounts, eligibility flags, and letter grades
- Compare classic lookups (`VLOOKUP`, `INDEX` + `MATCH`) with modern ones (`XLOOKUP`, `XMATCH`)
- Use **date functions** (`DATEDIF`, date subtraction) to calculate ages and day differences
- Control number precision with `ROUND`, `CEILING`, and `FLOOR`
- Create **dynamic references** with `INDIRECT` and `OFFSET`
- Extract live subsets of data using the **`FILTER`** dynamic-array function
- Summarize data with `SUMIFS`, `COUNTIFS`, and `AVERAGEIFS` over structured **Excel Tables**

---

## <span style="color:#FF6F00">🎯 Problem Statement</span>

> **Objective:** Given plain tables of products, students, employees, and monthly sales, use Excel formulas alone — no macros, no manual edits — to categorize, look up, round, filter, and summarize the data automatically.

You're handed raw lists that mirror everyday business and classroom data: products with prices, students with marks, employees with salaries, and a row of monthly sales figures. The task is to **let formulas do the thinking** — flag discount-eligible products, grade students, find a salary from an employee ID, total sales for any number of months, and pull out only the top scorers.

| 📂 Feature | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| Sales Data | Sheet | Product prices, discount logic, `SUMIFS`, `VLOOKUP`, `INDEX`/`MATCH` |
| XLOOKUP / XMATCH | Sheet | Modern, flexible replacements for older lookup functions |
| DATE&TIME | Sheet | Age and duration calculations with `DATEDIF` |
| MATH | Sheet | Rounding prices to the nearest hundred three different ways |
| INDIRECT / OFFSET RANGE | Sheet | Formulas whose references are built or shifted dynamically |
| FILTER | Sheet | Spill-range output of rows that meet a condition |
| Employee Data | Sheet | Salary lookup by employee ID using `XLOOKUP` |
| Students Grade | Sheet | Nested-`IF` grading, `COUNTIFS`, `AVERAGEIFS`, table-based `FILTER` |

The goal is to demonstrate **real-world Excel formula skills** — the kind used daily in reports, dashboards, and quick analyses.

---

## <span style="color:#4CAF50">✨ Key Features</span>

| Feature | Description |
|--------|-------------|
| 📑 **Ten Themed Worksheets** | Each sheet isolates one formula family so it can be studied on its own |
| 🏷️ **Tiered Discount Logic** | Nested `IF` assigns a discount and a `YES`/`NO` eligibility flag per product |
| 🔎 **Old vs. New Lookups** | `VLOOKUP` and `INDEX`+`MATCH` side by side with `XLOOKUP` and `XMATCH` |
| 🛡️ **Graceful Not-Found Handling** | `XLOOKUP` returns `"-"` and `XMATCH` is wrapped in `IFNA` for missing values |
| 📅 **Date Intelligence** | `DATEDIF(...,"Y")` for exact age; simple subtraction for day counts |
| 🔢 **Three Rounding Styles** | `ROUND(…,-2)`, `CEILING(…,100)`, and `FLOOR(…,100)` compared directly |
| 🧩 **Dynamic References** | `INDIRECT("B2:B6")` and `OFFSET(B2,0,0,1,B5)` respond to changing inputs |
| 🎯 **Live Filtering** | `FILTER` spills only the rows that meet a score threshold |
| 📊 **Structured Tables** | `Table1` (Students) and `Table2` (Sales) enable readable `Table1[Score]` references |
| 🎓 **Automatic Grading** | Seven-tier grade scale from `A` to `Fail` driven by the average score |

---

## <span style="color:#9C27B0">🏗️ Project Structure</span>

```
📦 excel-formula-lab/
│
├── 📊 Excel_p1.xlsx          ← Main workbook (10 worksheets)
│     ├── Sales Data          ← Discounts, SUMIFS, VLOOKUP, INDEX/MATCH
│     ├── XLOOKUP             ← Modern single-value lookup
│     ├── XMATCH              ← Position lookup with IFNA
│     ├── DATE&TIME           ← DATEDIF ages & day differences
│     ├── INDIRECT            ← Dynamic cell references
│     ├── MATH                ← ROUND / CEILING / FLOOR
│     ├── FILTER              ← Dynamic-array filtering
│     ├── OFFSET RANGE        ← Variable-width range sums
│     ├── Employee Data       ← XLOOKUP salary finder
│     └── Students Grade      ← Grades, COUNTIFS, AVERAGEIFS, FILTER
│
└── 📄 README.md              ← Project documentation
```

---

## <span style="color:#00758F">🔄 Project Workflow</span>

```
Start
  │
  ▼
┌──────────────────────────────────┐
│ Enter raw data: Products,        │
│ Students, Employees, Sales       │
└────────────────┬─────────────────┘
                 │
      ┌──────────┼───────────┐
      ▼          ▼           ▼
┌───────────┐┌───────────┐┌────────────────┐
│ IF / AND  ││ VLOOKUP / ││ XLOOKUP /      │
│ Logic     ││ INDEX+MATCH││ XMATCH         │
└─────┬─────┘└─────┬──────┘└────────┬───────┘
      │            │                │
      └────────────┼────────────────┘
                   ▼
      ┌─────────────────────────────┐
      │ Date & Math: DATEDIF,       │
      │ ROUND, CEILING, FLOOR       │
      └──────────────┬──────────────┘
                     │
                     ▼
      ┌─────────────────────────────┐
      │ Dynamic: INDIRECT, OFFSET,  │
      │ FILTER                      │
      └──────────────┬──────────────┘
                     │
                     ▼
      ┌─────────────────────────────┐
      │ Summaries: SUMIFS, COUNTIFS,│
      │ AVERAGEIFS                  │
      └──────────────┬──────────────┘
                     │
                     ▼
                  Done ✅
```

---

## <span style="color:#E91E63">🛒 Part A — Sales Data: Logic, Discounts & Lookups</span>

### 📝 1. Product Table & Discount Logic

The `Sales Data` sheet holds **20 products** (IDs `201`–`220`) in an Excel Table (`Table2`, range `A1:H21`) with five calculated columns.

```excel
Discount            =IF(C2>=F5, C2*50%, IF(C2>=5000, C2*20%, 0))
Final price         =C2-D2
Discount Eligible?  =IF(C2>=5000, "YES", "NO")
```

| Product id | Product name | Product price | Discount | Final price | Eligible? |
|:----------:|--------------|--------------:|---------:|------------:|:---------:|
| 201        | Laptop       | 300           | 0        | 300         | NO        |
| 204        | Fridge       | 6700          | 1340     | 5360        | YES       |
| 205        | Ac           | 8000          | 1600     | 6400        | YES       |
| 206        | Washing machine | 9000       | 1800     | 7200        | YES       |
| 210        | Cpu          | 8700          | 1740     | 6960        | YES       |
| 213        | Headphone    | 3400          | 0        | 3400        | NO        |

> 💡 Products priced **₹5,000 or above** are flagged eligible and receive a **20% discount**; everything below is charged full price.

### 🔗 2. SUMIFS — Total Price for a Product Name

```excel
=SUMIFS($C$2:$C$21, $B$2:$B$21, B2)
```

*Sums the price column wherever the product name matches — for `Laptop` it returns **300**.*

### 🔍 3. VLOOKUP — Price from Product ID

| Input | Formula | Output |
|-------|---------|--------|
| Product id `213` | `=VLOOKUP(I9, A2:E21, 3, FALSE)` | **3400** |

*Exact-match (`FALSE`) lookup returning the 3rd column — Product price — for the ID typed in `I9`.*

### 🎯 4. INDEX + MATCH — Flexible Lookup

```excel
=INDEX($C$2:$C$21, MATCH(218, $A$2:$A$21, 0))
```

*`MATCH` finds the row of product `218` (Remote), and `INDEX` returns its price — **3400**. Unlike `VLOOKUP`, the return column can sit on either side of the lookup column.*

---

## <span style="color:#3F51B5">🔎 Part B — Modern Lookups: XLOOKUP & XMATCH</span>

### 🔍 5. XLOOKUP — Name from ID

The `XLOOKUP` sheet holds five products with IDs, names, prices, and companies.

```excel
=XLOOKUP(B10, A2:A6, B2:B6, "-")
```

| Input (ID) | Output (Name) |
|:----------:|:-------------:|
| 102        | **CPU**       |

*The fourth argument `"-"` is returned if the ID isn't found — no `#N/A` errors on screen.*

### 📍 6. XMATCH — Position of a Value

```excel
=XMATCH(F2, C:C, 0)
=IFNA(XMATCH(F2, C:C, 0), "not found")
```

| Product | Price (C) | Answer (F) | XMATCH Position |
|---------|----------:|-----------:|:---------------:|
| MOBILE  | 50000     | 50000      | 2               |
| TV      | 56000     | 76000      | 5               |
| SPECKAR | 78000     | 78000      | 4               |
| WASHING MACHINE | 45000 | 45000  | 6               |

*`XMATCH` returns the **position** of a value inside a column. Wrapping it in `IFNA` replaces a missing-value error with a friendly `"not found"` message.*

### 👔 7. XLOOKUP on Employee Data — Find a Salary

The `Employee Data` sheet lists **20 employees** (IDs `301`–`320`) across HR, IT, Sales, and Marketing.

```excel
=XLOOKUP(I8, A2:A21, D2:D21)
```

| Input (Employee ID) | Output (Salary) |
|:-------------------:|----------------:|
| 318 (VIBHUTI)       | **47000**       |

---

## <span style="color:#F72585">📅 Part C — Date & Time Functions</span>

### 🎂 8. Age and Day Difference

```excel
AGE              =DATEDIF(B2, D2, "Y")
DAYS DIFFERENCE  =D2-C2
```

| Student | DOB        | Start Date | End Date   | Age | Days Difference |
|---------|------------|------------|------------|:---:|----------------:|
| VAIDEHI | 2007-10-01 | 2026-01-05 | 2026-09-30 | 18  | 268             |
| VEDANT  | 2006-12-05 | 2026-02-12 | 2026-09-30 | 19  | 230             |
| JAYESH  | 1981-10-07 | 2026-07-02 | 2026-09-30 | 44  | 90              |
| MEGHA   | 1985-01-03 | 2026-01-04 | 2026-09-30 | 41  | 269             |
| RANJAN  | 1955-10-14 | 2026-01-07 | 2026-09-30 | 70  | 266             |
| OM      | 2013-02-16 | 2026-08-05 | 2026-09-30 | 13  | 56              |

*`DATEDIF` with the `"Y"` unit returns **completed years** between two dates, while plain subtraction gives the exact number of days.*

---

## <span style="color:#009688">🧮 Part D — Math & Rounding Functions</span>

### 🔢 9. ROUND vs. CEILING vs. FLOOR

```excel
ROUND    =ROUND(C2, -2)
CEILING  =CEILING(C2, 100)
FLOOR    =FLOOR(C2, 100)
```

| Product    | Price     | ROUND | CEILING | FLOOR |
|------------|----------:|------:|--------:|------:|
| LAPTOP     | 233.58    | 200   | 300     | 200   |
| MOBILE     | 456.70    | 500   | 500     | 400   |
| SMARTWATCH | 5437.80   | 5400  | 5500    | 5400  |
| HEADPHONE  | 43245.876 | 43200 | 43300   | 43200 |
| TV         | 54356.87  | 54400 | 54400   | 54300 |

> 💡 `ROUND` goes to the **nearest** hundred, `CEILING` always rounds **up**, and `FLOOR` always rounds **down**.

---

## <span style="color:#795548">🧩 Part E — Dynamic Functions: INDIRECT, OFFSET & FILTER</span>

### 🔗 10. INDIRECT — References Built from Text

```excel
=SUM(INDIRECT("B2:B6"))     → 200000
=INDIRECT(D2)               → D2 holds "B3", so the result is 30000
```

*`INDIRECT` converts a text string into a real cell reference — change the text in `D2` and the output follows.*

### 📏 11. OFFSET — Variable-Width Sales Total

The `OFFSET RANGE` sheet stores monthly sales from **JAN** to **JUL**.

```excel
=SUM(OFFSET($B$2, 0, 0, 1, $B$5))
```

| Input (Months) | Months Covered | Total Sales |
|:--------------:|----------------|------------:|
| 4              | JAN – APR      | **146000**  |

*Change `B5` to `6` and the same formula sums six months — no formula editing needed.*

### 🎯 12. FILTER — Top Scorers Only

```excel
=FILTER(B2:C6, C2:C6>=75)
```

| Student | Score |
|---------|------:|
| VAIDEHI | 78    |
| VEDANT  | 76    |
| MEGHA   | 75    |

*`FILTER` is a **dynamic-array** function: one formula spills a whole result block, and students scoring below 75 (JAYESH, OM) are excluded automatically.*

---

## <span style="color:#3776AB">🎓 Part F — Students Grade & Conditional Aggregates</span>

### 🏅 13. Nested-IF Grading

The `Students Grade` sheet holds **20 students** with Maths and Science marks in `Table1` (range `A1:H21`).

```excel
Score  =AVERAGE(C2:D2)
Grade  =IF(E2>=90,"A",IF(E2>=80,"B",IF(E2>=70,"C",IF(E2>=60,"D",IF(E2>=50,"E",IF(E2>=40,"F","Fail"))))))
Above 80?                =IF(E2>80, "Yes", "No")
Math & Science Above 80? =IF(AND(C2>80, D2>80), "OK", "NOT OK")
```

| Score Range | Grade |
|:-----------:|:-----:|
| 90 and above | A |
| 80 – 89.9 | B |
| 70 – 79.9 | C |
| 60 – 69.9 | D |
| 50 – 59.9 | E |
| 40 – 49.9 | F |
| Below 40 | Fail |

| Student   | Maths | Science | Score | Grade | Above 80? | Both > 80? |
|-----------|------:|--------:|------:|:-----:|:---------:|:----------:|
| Vaidehi   | 90    | 100     | 95.0  | A     | Yes       | OK         |
| Megha     | 90    | 89      | 89.5  | B     | Yes       | OK         |
| Jayesh    | 86    | 69      | 77.5  | C     | No        | NOT OK     |
| Ranjan    | 80    | 90      | 85.0  | B     | Yes       | NOT OK     |
| Prayag    | 100   | 97      | 98.5  | A     | Yes       | OK         |
| Aayushi   | 23    | 54      | 38.5  | Fail  | No        | NOT OK     |

### 📊 14. COUNTIFS, AVERAGEIFS & Table-Based FILTER

```excel
=COUNTIFS(Table1[Score], ">60")                       → 17
=AVERAGEIFS(Table1[Score], Table1[Score], ">60")      → 90.38
=FILTER(Table1[], Table1[Score] > 80)                 → 15 rows
```

| Metric | Result |
|--------|-------:|
| 🎯 Students scoring above 60 | **17** |
| 📈 Average score of those students | **90.38** |
| 🏆 Students scoring above 80 (FILTER spill) | **15** |

*Structured references like `Table1[Score]` expand automatically when new students are added to the table.*

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🧠 Logical Functions | `IF` (nested up to 6 levels), `AND`, `IFNA` |
| 🔎 Lookup Functions | `VLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH` |
| 📅 Date Functions | `DATEDIF`, date subtraction |
| 🔢 Math Functions | `ROUND`, `CEILING`, `FLOOR`, `SUM`, `AVERAGE` |
| 🧩 Dynamic References | `INDIRECT`, `OFFSET` |
| 📊 Conditional Aggregates | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| 🎯 Dynamic Arrays | `FILTER` with spill ranges |
| 🗂️ Structured References | `Table1[Score]`, `Table1[]` |

---

## <span style="color:#3776AB">🛠️ Tech Stack</span>

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 365 / 2021+ | Spreadsheet engine (required for `XLOOKUP`, `XMATCH`, `FILTER`) |
| 🧠 **Logical Functions** | `IF`, `AND`, `IFNA` | Decision-making and error handling |
| 🔎 **Lookup Functions** | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`/`MATCH` | Retrieving values by key |
| 📅 **Date Functions** | `DATEDIF` | Age and duration calculation |
| 🧩 **Reference Functions** | `INDIRECT`, `OFFSET` | Dynamic ranges |
| 🎯 **Dynamic Arrays** | `FILTER` | Spill-range results |
| 🗂️ **Excel Tables** | `Table1`, `Table2` | Structured, auto-expanding data |

---

## <span style="color:#4CAF50">📈 Results & Insights</span>

After opening the workbook, the following outcomes are produced:

- ✅ **10 Worksheets** covering logic, lookups, dates, math, dynamic references, and filtering
- 🏷️ **Automatic discounts** — 20% off for every product priced ₹5,000 or more
- 🔎 **Five lookup techniques** demonstrated: `VLOOKUP`, `INDEX`+`MATCH`, `XLOOKUP`, `XMATCH`, and table-based `FILTER`
- 📅 **Exact ages and day counts** computed for six people with `DATEDIF`
- 🔢 **Rounding compared** three ways on the same set of prices
- 🧩 **Dynamic totals** — `OFFSET` sums any number of months (4 months → ₹1,46,000)
- 🎓 **20 students graded automatically**, with 17 scoring above 60 and a 90.38 average among them
- 🎯 **Live filtering** — 15 students scoring above 80 extracted with a single formula

---

## <span style="color:#F44336">⚠️ Known Quirks & Fixes</span>

A few spots in the workbook behave unexpectedly. They're documented here so they can be fixed or used as debugging exercises.

| Sheet | Cell(s) | Issue | Suggested Fix |
|-------|---------|-------|---------------|
| Sales Data | `D2:D21` | Discount formula compares price to `F5`, `F6`, … which hold `YES`/`NO` text (and blank cells for rows 19–21), so rows `218`, `219`, `220` receive an unintended **50%** discount | Use `=IF(C2>=5000, C2*20%, 0)` |
| Sales Data | `O4` | `VLOOKUP` is typed with semicolons (`;`) and stored as text, so it never calculates | Use commas, or enter it as a real formula |
| Students Grade | `F11` | Naitik's formula points at `E110` instead of `E11` and is missing the `F` tier, so a score of 44 shows `Fail` | Copy the formula from `F10` |
| Employee Data | `C16` | Department `"IT "` has a trailing space | Wrap with `TRIM()` or retype |
| XMATCH | `F3`, `G5` | `F3` (76000) doesn't match the row's own price, and `G5` looks up `F2` instead of `F5` | Align lookup cells with their rows |

---

## <span style="color:#FFB300">🏆 Advantages</span>

| Advantage | Detail |
|-----------|--------|
| 🎓 **Progressive Difficulty** | Moves from basic `IF` logic to lookups, then dynamic arrays |
| 🔎 **Old vs. New Side by Side** | `VLOOKUP` and `INDEX`/`MATCH` compared with `XLOOKUP` and `XMATCH` |
| 📊 **Realistic Data** | Products, students, and employees mirror everyday spreadsheets |
| 📚 **Educational** | Every formula is paired with its actual output |
| 🖥️ **No Macros or Add-ins** | Runs on standard Microsoft Excel 365 / 2021 |
| 🧪 **Extensible** | Easy to add `PIVOT`, `XLOOKUP` with wildcards, `LET`, or `SORT` sheets |
| 🗂️ **Table-Driven** | Structured references keep formulas readable and auto-expanding |
| 📖 **Clearly Sectioned** | One sheet per concept keeps the workbook easy to navigate |

---

## <span style="color:#607D8B">📄 License</span>

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## <span style="color:#0077B5">👤 Author</span>

<div align="center">

### Vaidehi Vyas

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-000000?style=for-the-badge&logo=github&logoColor=white&labelColor=000000)](https://github.com/)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-000000?style=for-the-badge&logo=linkedin&logoColor=0077B5&labelColor=000000)](https://www.linkedin.com/)

> *"Every formula starts with a question the raw data can't answer alone."*

**🎓 Role:** Excel & Data Analytics Learner \
**📍 Location:** India \
**🛠️ Skills:** Excel · XLOOKUP · FILTER · Nested IF · Date Functions · Data Cleaning

</div>

---

## <span style="color:#795548">🙏 Acknowledgements</span>

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official Excel function reference
- 🔗 [W3Schools Excel](https://www.w3schools.com/excel/) — Beginner-friendly Excel reference
- 🧮 [ExcelJet](https://exceljet.net/) — Clear formula examples and explanations
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Chandoo.org](https://chandoo.org/) — Practical Excel tutorials

---

<div align="center">

---

### <span style="color:#00E5FF">⭐ If this project helped you learn Excel formulas better, consider giving it a star!</span>

[![Stars](https://img.shields.io/badge/⭐-Star%20this%20repo-000000?style=for-the-badge&labelColor=000000&color=FFD700)](#)
[![Forks](https://img.shields.io/badge/🍴-Fork%20it-000000?style=for-the-badge&labelColor=000000&color=4CAF50)](#)
[![Visitors](https://api.visitorbadge.io/api/visitors?path=excel-formula-lab%2Freadme&label=Visitors&countColor=%2300E5FF&style=for-the-badge)](#)

<br/>

![Footer](https://capsule-render.vercel.app/api?type=waving&color=0:000000,50:0D0D0D,100:000000&height=100&section=footer)

*Made with 💜 and 📗 by **Vaidehi Vyas** — Last updated: 02 October, 2026*

</div>
