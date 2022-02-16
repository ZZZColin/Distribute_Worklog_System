Attribute VB_Name = "Get_Data_Module"
Function get_data(file_path, column_count, Optional OverwriteStrRow, Optional OverwriteStrColumn)
    
    Set eapp = CreateObject("excel.application")
    
    Set efile = eapp.Workbooks.Open(file_path)
    
    ColumnLetter = Num2Letter(column_count)
    
    Data = efile.Sheets(1).Range("A1:" & ColumnLetter & efile.Sheets(1).UsedRange.Rows.Count).Value
    
    Call add_log("       Data Range: " & "A1:" & ColumnLetter & efile.Sheets(1).UsedRange.Rows.Count)
    
    If IsMissing(OverwriteStrRow) = False And IsMissing(OverwriteStrColumn) = False Then
        
        OverwriteStr = efile.Sheets(1).Cells(OverwriteStrRow, OverwriteStrColumn).Text
        
        Set result = New Collection
        
        result.Add Data
        
        result.Add OverwriteStr
        
        Set get_data = result
        
        GoTo final:
        
    End If

    get_data = Data

final:
    
    eapp.Application.DisplayAlerts = False
    efile.Close
    Set efile = Nothing
    eapp.Application.DisplayAlerts = True
    
    eapp.Quit
    Set eapp = Nothing
    
End Function
