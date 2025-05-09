# Excel Calc Pad

Excel Calc Pad is a macro-powered Excel utility designed to speed up typing of engineering formulas. It helps you format subscripts, superscripts, and Greek letters quickly using simple syntax and keyboard shortcuts.


## How It Works

### 1. **Subscripts and Superscripts**
- Use `.` before a subscripted character.
- Use `^` before a superscripted character.
- Then press **Ctrl + H** on the selected cell.

**Example:**
p.1p.2p.3p^4
→ becomes →  
![image](https://github.com/user-attachments/assets/c092a035-7fe7-45f6-919f-63d4d0c8c62e)



### 2. **Greek Letters**
- Use `,` before a Latin character to convert it to its Greek equivalent.
- Then press **Ctrl + G** on the selected cell.

**Example:**
,g
→ becomes →  
![image](https://github.com/user-attachments/assets/80a4a208-0b4c-465d-99a3-cb025afa9c46)



##  Files in This Repo

| File Name | Description |
|-----------|-------------|
| `Excel Calc Pad.xlsm` | The main macro-enabled workbook. Save this as your personal workbook or copy the macros into your own. |
| `Copy my contents.txt` | Plain text version of the macros. Copy this into the Visual Basic Editor (`Alt + F11`). |
| `import me.bas` | Exported VBA module. You can import it into your workbook using the VB Editor: `File > Import File...` |


##  Installation Options

### Option 1 – Use the Workbook Directly
1. Open `Excel Calc Pad.xlsm`.
2. Enable macros.
3. Start typing formulas and use the shortcuts.

### Option 2 – Add Macros to Your Own Workbook
1. Open your workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Either:
   - Copy-paste the code from `Copy my contents.txt`, or
   - Import `import me.bas` (`File > Import File...`).
4. Save your file as `.xlsm`.


## Notes
- Macros must be enabled for the shortcuts to work.
- This is for engineering-style formula entry in Excel cells.



## Suggestions?

Feel free to open an issue or submit a pull request.


