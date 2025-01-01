# How to Use the `Transform035Field` Macro in Excel

This guide explains how to set up and run the `Transform035Field` macro in Excel. The macro extracts specific values from the "035 field" column, processes them, and outputs results into a new column named "Extracted OCLC Number."

## Prerequisites

1. **Excel Installed**: Ensure you have Microsoft Excel installed on your computer.
2. **Macro-Enabled Workbook**: Save your Excel workbook as a macro-enabled file (`.xlsm`).
3. **Basic Understanding of VBA**: Familiarity with how to open the VBA editor and run macros is helpful, though these instructions should suffice if you are unfamiliar with macros.

---

## Steps to Use the Macro

### 1. Open the Excel Workbook
Ensure your workbook contains a column labeled **`035 field`** in the first row. This should be there be default with exports from the BIG CAT

### 2. Open the VBA Editor
1. Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor. (Also available under Tools / Macro / Visual Basic Editor)
2. In the VBA editor, go to **Insert > Module**.

### 3. Paste the Macro Code
1. Copy the entire macro code provided.
2. Paste the code into the module window.

### 4. Save the Workbook
1. Save your workbook as a macro-enabled file (`.xlsm`).
2. You may now close the VBA editor if desired.

### 5. Run the Macro
1. Press `Alt + F8` in Excel. (Also available under Macro / Macros )
2. Select **`Transform035Field`** from the list.
3. Click **Run**.

---

## What the Macro Does

1. **Identifies the "035 field" Column**:
   - Locates the column labeled "035 field" in the first row.
   - If the column is missing, it shows an error message.

2. **Creates or Reuses the Output Column**:
   - If a column labeled "Extracted OCLC Number" exists, it reuses it.
   - Otherwise, it creates a new column to the right of "035 field."
   - Formats the new column as a **Text** column.

3. **Extracts and Processes Data**:
   - Scans each row in the "035 field" column.
   - Extracts values starting with `$a` and matching specific prefixes.
   - Removes prefixes like `(OCoLC)` before copying.
   - Ensures only **unique values** are added, separated by semicolons (`;`).

4. **Adjusts Column Width**:
   - Automatically resizes the new column to fit the longest entry.

5. **Displays a Success Message**:
   - Shows "Transformation complete!" when processing is finished.

---

## Notes

- Ensure the "035 field" column contains properly formatted data for `$`-delimited processing.
- This macro works row by row and preserves unique extracted values within each row.
- You can customize the valid prefixes in the `validPrefixes` array inside the VBA code if needed.

---

## Troubleshooting

- **Macro Not Listed**: Ensure the code is saved in the correct workbook module.
- **Error Messages**: Check if the "035 field" column exists and is spelled correctly.
- **Security Warnings**: Enable macros in your Excel settings by navigating to `File > Options > Trust Center > Trust Center Settings > Macro Settings` and selecting **Enable all macros**.

---

## Making the Macro Available in All Workbooks

If you'd like to make this macro available across all Excel workbooks, follow these steps:

### 1. Open Excel's Personal Macro Workbook
1. Open Excel.
2. Press `Alt + F11` to open the VBA Editor. (Also available under Tools / Macro / Visual Basic Editor)
3. In the VBA Editor, look for a workbook named `PERSONAL.XLSB` under **VBAProject**.
   - If you don’t see it, you need to create it:
     - Close the VBA Editor.
     - In Excel, record a dummy macro:
       - Go to **View > Macros > Record Macro**.
       - In the **Store macro in** dropdown, select **Personal Macro Workbook**.
       - Click **OK**, and stop the recording immediately by clicking **Stop Recording** on the status bar.
     - This creates the `PERSONAL.XLSB` workbook.

### 2. Add the VBA Script to the Personal Macro Workbook
1. Open the VBA Editor (`Alt + F11`). (Also available under Tools / Macro / Visual Basic Editor)
2. Locate `PERSONAL.XLSB` under **VBAProject**.
3. Expand the project by clicking the `+` icon next to it.
4. Right-click **Modules**, select **Insert > Module**, and paste [the macro code](https://github.com/samato88/VBATransform035/blob/main/Transform035Field.bas) into the module.
5. Save your changes:
   - Go to **File > Save PERSONAL.XLSB** in the VBA Editor.

### 3. Close and Reopen Excel
1. Close Excel to ensure the `PERSONAL.XLSB` file is saved properly.
2. Reopen Excel to load the `PERSONAL.XLSB` file automatically.

### 4. Run the Macro in Any Workbook
1. Open any Excel workbook.
2. Press `Alt + F8` to open the **Macro Dialog Box**.
3. Select **Transform035Field** from the list.
4. Click **Run**.

---
# Credits

This guide and the Transform035Field VBA script were created by ChatGPT 4o in Jan 1 2025. 
Enjoy your automated data processing with the `Transform035Field` macro!