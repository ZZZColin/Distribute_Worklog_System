Attribute VB_Name = "Convert_Num2Letter_Module"
Function Num2Letter(column_num)

    ColumnLetter = ThisWorkbook.Sheets(1).Cells(1, column_num).Address
    ColumnLetter = Mid(ColumnLetter, 2, InStr(2, ColumnLetter, "$") - 2)
    
    Num2Letter = ColumnLetter
    
End Function
