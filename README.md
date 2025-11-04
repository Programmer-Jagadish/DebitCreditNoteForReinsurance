# 💼 Reinsurance Debit Note Generator

A simple Java tool that automates the creation of **Debit Notes** for reinsurance transactions.

### 🧮 Features
- Reads input data from Excel.
- Performs premium, brokerage, and commission calculations.
- Generates Word Debit Note files automatically.
- Marks processed records in Excel.
- Works both in IntelliJ and via `.jar + .bat` for end users.

### 📂 Folder Structure
resources/
├── DebitNoteCalculations.xlsx ← Input Excel file
├── DebitNoteTemplate.docx ← Word template for debit notes
└── output/ ← Auto-generated debit notes

### ⚙️ How to Run
1. Edit `resources/DebitNoteCalculations.xlsx`
2. Run `RunTool.bat`
3. Generated notes appear inside `resources/output/`

### 🧰 Tech Stack
- Java 17
- Apache POI (Excel + Word)
- IntelliJ IDEA
