Attribute VB_Name = "Overwrite_Module"
Sub overwrite()

    'CLEAR LOG ---------------------------------------------------------------------------------
    ThisWorkbook.Sheets("Log").TextBox1.Text = ""
    
    'SET fso TO CHECK FolderPath EXIST AND GET LastModifiedTime --------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'INI ENVIR VAR -----------------------------------------------------------------------------
    FolderPath = ThisWorkbook.Sheets("Log").Cells(1, 2).Text
    If FolderPath = "" Then Call add_log("FolderPath Is Empty, Please Check."): Exit Sub
    If fso.FolderExists(FolderPath) = False Then Call add_log("FolderPath Not Exist,Please Check."): Exit Sub
    
    'DateColumn = ThisWorkbook.Sheets("Log").Cells(2, 2).Value
    'If IsNumeric(DateColumn) = False Or DateColumn = "" Then Call add_log("DateColumn Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    'DateColumnLetter = Num2Letter(DateColumn)
    
    KeyColumn = ThisWorkbook.Sheets("Log").Cells(3, 2).Value
    If IsNumeric(KeyColumn) = False Or KeyColumn = "" Then Call add_log("KeyColumn Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    
    ColumnCount = ThisWorkbook.Sheets("Log").Cells(4, 2).Value
    If IsNumeric(ColumnCount) = False Or ColumnCount = "" Then Call add_log("ColumnCount Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    
    'If ThisWorkbook.Sheets("Log").TextBox2.BackColor <> RGB(128, 255, 128) Or _
    '   ThisWorkbook.Sheets("Log").TextBox3.BackColor <> RGB(128, 255, 128) Then
    '    Call add_log("Please Check Start Date and End Date."): Exit Sub
    'End If
    'StartDate = ThisWorkbook.Sheets("Log").TextBox2.Text: EndDate = ThisWorkbook.Sheets("Log").TextBox3.Text
    
    If InStr(1, ThisWorkbook.Sheets("Log").Cells(5, 2).Text, ",") = 0 Then Call add_log("Overwrite String Cell Invalid Format, Please Check."): Exit Sub
    OverwriteStrRow = Trim(Split(ThisWorkbook.Sheets("Log").Cells(5, 2).Text, ",")(0))
    If OverwriteStrRow = "" Or IsNumeric(OverwriteStrRow) = False Then Call add_log("Overwrite String Cell Row Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    OverwriteStrColumn = Trim(Split(ThisWorkbook.Sheets("Log").Cells(5, 2).Text, ",")(1))
    If OverwriteStrColumn = "" Or IsNumeric(OverwriteStrColumn) = False Then Call add_log("Overwrite String Cell Column Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    OverwriteStrRow = CInt(OverwriteStrRow): OverwriteStrColumn = CInt(OverwriteStrColumn)
    
    'GET KEY -----------------------------------------------------------------------------------
    'STRUCTURE:
    'KEY: KEY
    'VALUE: ROW AS COLLECTION
    Set KeyDic = get_key(KeyColumn)
    
    With ThisWorkbook.Sheets("Main")
        
        'STORE OVERWRITE SUCCESS KEY -----------------------------------------------------------
        Set OverwriteDic = CreateObject("scripting.dictionary")
        
        FilePath = Dir(FolderPath & "\*.xlsx")
    
        Do While FilePath <> ""
            
            FilePath = FolderPath & "\" & FilePath
            
            Call add_log("--> " & FilePath)
            
            'GET LastModifiedTime --------------------------------------------------------------
            Set FileObj = fso.GetFile(FilePath)
            
            LastModifiedTime = FileObj.DateLastModified
            
            Call add_log("       Last Modified Time: " & LastModifiedTime)
            
            'GET DATA BASED ON ColumnCount PROVIDED AND OverwriteStr ---------------------------
            Set result = get_data(FilePath, ColumnCount, OverwriteStrRow, OverwriteStrColumn)
            Data = result(1)
            OverwriteStr = result(2)
            Call add_log("       Overwrite String: " & OverwriteStr)
            
            'COUNT OVERWRITE -------------------------------------------------------------------
            OverwriteCount = 0

            'GET OverwriteArr ------------------------------------------------------------------
            OverwriteArr = Split(OverwriteStr, " ")
            
            'FOR EACH ITEM IN OverwriteArr
            'NEED CHECK IF IT WAS NUMBER AND WITHIN DATA RANGE
            For Each a In OverwriteArr
            
                OverwriteRow = Trim(a)
                
                'CHECK IF NUMBER ---------------------------------------------------------------
                If OverwriteRow = "" Or IsNumeric(OverwriteRow) = False Then
                    
                    Call add_log("             Is Empty Or Is Not Number Format: " & OverwriteRow)
                    
                    GoTo next_overwrite
                    
                End If
                
                'CHECK IF WITHIN DATA RANGE ----------------------------------------------------
                If CInt(OverwriteRow) < 1 Or CInt(OverwriteRow) > UBound(Data, 1) Then
                
                    Call add_log("             Not Within Data Range: " & OverwriteRow)
                    
                    GoTo next_overwrite
                    
                End If
                
                OverwriteRow = CInt(OverwriteRow)
                
                'GET KEY -----------------------------------------------------------------------
                Key = Data(OverwriteRow, KeyColumn)

                If KeyDic.exists(Key) = False Then
                    
                    'IF KEY NOT FOUND ----------------------------------------------------------
                    Call add_log("             Row: " & OverwriteRow & ", Key: " & Key & " Not Exist In Tracking Log")
                    
                    GoTo next_overwrite
                
                Else:

                    If KeyDic(Key).Count > 1 Then
                        
                        'IF MULTI RECORDS FOUND WITH SAME KEY ----------------------------------
                        RowStr = ""
                            
                        For Each r In KeyDic(Key)
                        
                            If RowStr = "" Then RowStr = r Else RowStr = RowStr & ", " & r
                            
                        Next
                            
                        Call add_log("             Row: " & OverwriteRow & ", Key: " & Key & " Found On Multi Rows " & RowStr & ", Please Check")
                    
                        GoTo next_overwrite
                        
                    Else:
                        
                        If OverwriteDic.exists(Key) Then
                            
                            'IF THIS KEY HAS BEEN OVERWRITTEN --------------------------------------
                            Call add_log("             Row: " & OverwriteRow & ", Key: " & Key & " Skipped As Duplicate Overwrite On Row " & OverwriteDic(Key) & ", Please Check")
                    
                            GoTo next_overwrite
                            
                        Else:
                        
                            'OVERWRITE -------------------------------------------------------------
                            TargetRow = KeyDic(Key)(1)
                            
                            For i = 1 To UBound(Data, 2)
                            
                                .Cells(TargetRow, i) = Data(OverwriteRow, i)
                                
                            Next i
                            
                            Call add_log("             Row: " & OverwriteRow & ", Key: " & Key & " Overwritten On Row " & TargetRow)
                            
                            OverwriteDic(Key) = TargetRow
                            
                            OverwriteCount = OverwriteCount + 1
                            
                        End If
                        
                    End If
                
                End If
                
next_overwrite:
                
            Next
                
            Call add_log("       " & OverwriteCount & " Record(s) Overwritten In Total." & Chr(10))
            
            FilePath = Dir
            If FilePath = "" Then
                Exit Do
            End If
    
        Loop
        
    End With
    
    Call add_log("Done!")

End Sub
