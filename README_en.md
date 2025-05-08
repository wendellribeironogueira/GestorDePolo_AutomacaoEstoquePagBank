<p align="center">
  <img src="logo.png" alt="GestorDePolo Logo" width="140">
</p>

<h1 align="center">GestorDePolo_AutomacaoEstoquePagBank</h1>

<p align="center"><strong>Automates local inventory control, reverse validation, and operational reporting using Excel VBA.</strong></p>

---

## 🔖 Overview

**Gestor de Polo** is an automated solution built in **VBA for Excel**, developed during the **PagResolve Project** at **PagBank**.  
Its purpose is to eliminate manual effort from local inventory control and to generate actionable reports from CSV files exported by the **Workfinity (iSolution)** platform.

---

## 🔹 Main Modules

### ✅ IMPORT
- Imports CSV exported from Workfinity.
- Updates the `ESTOQUE.xlsm` spreadsheet, marking equipment as “Activated”.
- Registers replaced items into the **REVERSA** tab.
- Prevents duplicates using `Scripting.Dictionary`.
- Generates process summary.

### ✅ REVERSE
- Serial validation from returned equipment.
- Integrates with a barcode reader for fast entry.
- Shows ✅ or ❌ depending on serial presence.

### ✅ REPORT
- Analyzes CSV data to produce dashboards.
- KPIs: SLA, reopened calls, activity by city, technician, and service type.
- Detects expired or about-to-expire service orders.

---

## 🧰 Technologies

- Excel VBA (Visual Basic for Applications)
- Dynamic Tables & Charts
- ActiveX TextBoxes and Buttons
- CSV parser + Dictionary structure

---

## 🧾 License

This project is licensed under **Creative Commons BY-NC-ND 4.0**.  
Originally created as a personal initiative while employed at PagBank, for internal process optimization.  
Use is allowed for educational, demonstration, or portfolio purposes only.

[Full license details](https://creativecommons.org/licenses/by-nc-nd/4.0/deed.en)

---

## 👨‍💻 Author

**Wendell Ribeiro Nogueira**  
Support & Automation Specialist  
[GitHub](https://github.com/wendellribeironogueira) • [LinkedIn](https://www.linkedin.com/in/wendell-ribeiro-nogueira)
