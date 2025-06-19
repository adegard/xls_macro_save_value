# 📦 Excel Cell Settings Exporter/Importer

This personal Excel macro provides a simple way to **export selected cell values** to a CSV file and **import them later** into the same or a different workbook. It's a lightweight tool to **transfer settings or key values** between Excel workbooks—especially useful for dashboards, configuration sheets, and recurring reports.

</br>
<img src="https://github.com/adegard/xls_macro_save_value/blob/main/Immagine 2025-06-19 141159.jpg"  align="center">


## ✨ Features

- Save any list of cell values to a `.csv` file via a custom user form  
- Restore saved values to their original addresses with one click  
- Customizable input: just type or paste cell references (e.g., `C2, D5, G9`)  
- Uses standard `InputBox` dialogs for quick and intuitive use  
- No external dependencies—fully native Excel VBA  

## 🧰 How It Works

1. User enters a **comma-separated list of cell addresses** in the TextBox (e.g. `"C2, D5, G6"`).
2. Press **Save to CSV** to export their current values.
3. Press **Restore from CSV** to load and apply those values back into the sheet.

Both buttons display a file dialog to choose or save your CSV file. You'll get a confirmation message when complete.

## 📦 Installing as a Personal Macro

To add this feature to your **Personal Macro Workbook**, follow these steps:

1. Open Excel and press `Alt + F11` to open the VBA editor.
2. In the **Project Explorer**, locate `PERSONAL.XLSB` or create a new module inside it.
3. Insert a new **UserForm** and add the following elements:
    - `Label1` – for instructions
    - `TextBox1` – to enter cell addresses
    - `btnSave` – Save button
    - `btnRestore` – Restore button
4. Paste the provided VBA code into the UserForm’s code window.
5. Press `F5` to test the UserForm.

To display it later, you can assign a macro to a shortcut or button that runs (now it is Ctrl+Shift+S):
```vba
UserForm1.Show
```
## Sample Default Cell List
Please update in the code defaults values:

</br>
<img src="https://github.com/adegard/xls_macro_save_value/blob/main/Immagine 2025-06-19 141138.jpg"  align="center">

