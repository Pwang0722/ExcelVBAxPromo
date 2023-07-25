<h2 align="center">Excel: Efficient Data Management with VBA Macro and Formula</h1>
</div>

### Spreadsheet Examples
- [Asia Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(Asia%20Content)%20Clean%20Template.xlsm)
- [English Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(English%20Content)%20Clean%20Template.xlsm)
---

### Outline
An example of using the FILTER function to auto-fill similar inputs across multiple worksheets, and of reformatting the data and exporting them as PDFs using VBA Macros.

---

### Ｍethod 
Take spreadsheet [Asia Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(Asia%20Content)%20Clean%20Template.xlsm) as an example:
- Fill in the data under columns A to M in the sheet titled "TITLE LIST". Based on the data you have filled in, a code will be generated from a formula in column N.
- There is a formula in cell B19 in the sheets titled from "1B. ###" to "13A. ###", which retrieves the codes from column N in the "TITLE LIST" sheet and automatically fills in the data based on different requirements in each sheet.

Formula example:
  ```bash
 =IFERROR(FILTER('TITLE LIST'!A:N,('TITLE LIST'!N:N="AENG FMALLN")+('TITLE LIST'!N:N="GMAND FMALLN")+('TITLE LIST'!N:N="OMAND FMALLN")+('TITLE LIST'!N:N="OBM FMALLN")+('TITLE LIST'!N:N="ASOT ONLYALLN")+('TITLE LIST'!N:N="GSOT ONLYALLN")+('TITLE LIST'!N:N="OSOT ONLYALLN")+('TITLE LIST'!N:N="AENG FM05BN")+('TITLE LIST'!N:N="GMAND FM05BN")+('TITLE LIST'!N:N="OMAND FM05BN")+('TITLE LIST'!N:N="OBM FM05BN")+('TITLE LIST'!N:N="ASOT ONLY05BN")+('TITLE LIST'!N:N="GSOT ONLY05BN")+('TITLE LIST'!N:N="OSOT ONLY05BN")+('TITLE LIST'!N:N="GMAND FMALLY")+('TITLE LIST'!N:N="GSOT ONLYALLY")+('TITLE LIST'!N:N="GMAND FM05BY")+('TITLE LIST'!N:N="GSOT ONLY05BY")),"")
  ```
 - To prevent Excel from lagging while filling in data, I set the Calculation Options to Manual. Hence, users have to run the Calculate Now function every time they finish data entry or make changes. To make this process more convenient for everyone, I created a Macro that runs the Calculate Now function and assigned it to a button.

 Macro example:
  ```bash
  Sub CalculateWorkbook()
    Application.CalculateFull
    MsgBox ("Done. Sheets are ready to check.")
End Sub

---
