# 💰 Excel Finance Tracker with VBA

This project is a **dynamic personal finance tracker** built in **Microsoft Excel** using **VBA (Visual Basic for Applications)**. It allows users to automatically generate a Month-to-Date (MTD) and Year-to-Date (YTD) financial dashboard based on data from a transaction sheet.

Created by [@RastaDataVibes](https://github.com/RastaDataVibes), this tool is ideal for individuals and small businesses managing their finances directly in Excel.

---

## 📊 Features

- ✅ **Auto-calculates Month-to-Date totals** for income and categorized expenses
- ✅ **Year-to-Date summaries** with monthly totals for income and each expense category
- ✅ **Built-in pie and bar charts** for MTD and YTD overviews
- ✅ **Automatic month detection** for current summaries
- ✅ **Formulas linked to Transactions sheet** so tracker updates automatically
- ✅ Built with **clean, modular VBA code**

---

## 🧾 Workbook Structure

### 1. `Transactions` Sheet (Input Data)

| Year | Month | Date | Description | Category | Income | Debit | Balance |
|------|-------|------|-------------|----------|--------|-------|---------|
| 2025 | 8     | 24   | Egg sales   | Income   | 20000  |       | 20000   |
| 2025 | 8     | 24   | Fuel        | Diesel   |        | 3000  | 17000   |

> The tracker automatically reads from this sheet using category and month filters.

---

### 2. `FinanceTracker` Sheet (Auto-Generated)

#### 🟩 Month-to-Date (MTD)
- Displays **income** and **expenses** grouped by category (e.g., Salary, Diesel, Feeds)
- Columns:  
  - `Total`: actual total this month  
  - `Budgeted`: what was planned  
  - `Vs Budget`: difference between total and budgeted  

#### 🟦 Year-to-Date (YTD)
- Displays monthly totals for:
  - **Income**
  - **Each expense category**
  - **Total expenses**
- Side-by-side with MTD, not stacked

#### 📈 Charts
- **Pie Chart** → MTD breakdown by category
- **Bar Chart** → YTD comparison of income vs total expenses (monthly)

---

## 🚀 How to Use

1. Open the Excel file: `Book2.xlsm`
2. Ensure a sheet named `Transactions` exists with properly formatted data
3. Press `Alt + F11` to open the **VBA Editor**
4. Run the macro: `CreateFinanceTracker_Updated`
5. The `FinanceTracker` sheet will be generated with summaries and charts

---

## 🛠 Technologies Used

- Microsoft Excel (`.xlsm`)
- VBA (Visual Basic for Applications)
- Excel formulas and charts

---

## 📸 Screenshots (optional)

> *(Add screenshots here if you want to showcase how it looks)*

---

## 👤 Author

**Opio Bethle** ([@RastaDataVibes](https://github.com/RastaDataVibes))  
Medical Doctor | Data Enthusiast | Aspiring Epidemiologist

---

## 📄 License

This project is licensed under the **MIT License** – use freely and modify.

---

## ⭐ If you find this helpful, give it a star!


