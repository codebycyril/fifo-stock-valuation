# FIFO Stock Valuation (Excel VBA Project)

Stock valuation during audits was taking a huge amount of time — sometimes weeks of repetitive manual work.  
To simplify the process, I built an **Excel VBA automation** that performs **Closing Stock Valuation using FIFO**, turning a slow task into a structured and efficient workflow.

---

## ✨ Features
- Takes input from a **Purchase Register** and a **Closing Stock** sheet.  
- Generates a **Detailed Report** with bill-wise consumption and references.  
- Generates a **Summary Report** with product-wise closing stock values.  
- Preserves headers in all reports (data is written **only below the headers**).  
- Adds totals, formatting, and clear headings for easy navigation.  

---

## ⚠️ Important Notes
- **Do not remove the headers** in your Purchase Register or Closing Stock sheets.  
- Paste data **only below the headers** in each sheet.  
- The macro copies only the required columns needed for the valuation.  

---

## 📊 Impact
- Transformed a process that used to take **weeks** into one that runs in **minutes**.  
- Turned weeks of manual effort into minutes.  
- Improved accuracy and consistency in stock valuation.  
- Demonstrated how even small automations (like VBA) can bring big efficiency gains in audits.  

---

## 🚀 How to Use
1. Download and open the Excel file (`.xlsm`).  
2. Enable macros when prompted.  
3. Run the macro: **`codebycyrilFIFOStockValuation`**.  
4. Enter sheet names when asked (defaults are `PurchaseRegister` and `ClosingStock`).  
5. Two reports will be generated:  
   - **ClosingStockValuation** (detailed report)  
   - **SummaryReport** (summary with totals)  

---

## 📂 Repository Contents
- `FIFOStockValuation.xlsm` → Macro-enabled workbook with the code.  
- `ValueClosingStock.bas` → Exported VBA module for reading/importing the macro.  
- `SampleData.xlsx` → Test dataset (PurchaseRegister + ClosingStock) for quick demo.  
- `Screenshots/` → Folder with example output (Detailed Report & Summary Report).  
- `LICENSE` → MIT License text.  

---

## 📷 Example Output

**1. Run the Macro (Alt + F8)**  
![Run Macro](Screenshots/1.Run%20macro%20ALT+F8.png)

**2. Detailed Report**  
![Detailed Report](Screenshots/Closing%20Stock%20valuation-Detailed%20Report.png)

**3. Summary Report**  
![Summary Report](Screenshots/Stock%20valuation%20summary%20result.png)

---

## 🛠 Tech Stack
- Excel + VBA  
- Problem-driven logic (designed during real audit work)  

---

## 📄 License
This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.  
You are free to use, modify, and share it, as long as credit is given.  

---

👤 Built by [Cyril](https://github.com/codebycyril)  
