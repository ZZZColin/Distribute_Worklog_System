Attribute VB_Name = "Update_Module"
Sub update()
    
    'CLEAR LOG ---------------------------------------------------------------------------------
    ThisWorkbook.Sheets("Log").TextBox1.Text = ""
    
    'SET fso TO CHECK FolderPath EXIST AND GET LastModifiedTime --------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'INI ENVIR VAR -----------------------------------------------------------------------------
    FolderPath = ThisWorkbook.Sheets("Log").Cells(1, 2).Text
    If FolderPath = "" Then Call add_log("FolderPath Is Empty, Please Check."): Exit Sub
    If fso.FolderExists(FolderPath) = False Then Call add_log("FolderPath Not Exist,Please Check."): Exit Sub
    
    DateColumn = ThisWorkbook.Sheets("Log").Cells(2, 2).Value
    If IsNumeric(DateColumn) = False Or DateColumn = "" Then Call add_log("DateColumn Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    DateColumnLetter = Num2Letter(DateColumn)
    
    KeyColumn = ThisWorkbook.Sheets("Log").Cells(3, 2).Value
    If IsNumeric(KeyColumn) = False Or KeyColumn = "" Then Call add_log("KeyColumn Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    
    ColumnCount = ThisWorkbook.Sheets("Log").Cells(4, 2).Value
    If IsNumeric(ColumnCount) = False Or ColumnCount = "" Then Call add_log("ColumnCount Is Empty Or Is Not Number Format, Please Check."): Exit Sub
    
    If ThisWorkbook.Sheets("Log").TextBox2.BackColor <> RGB(128, 255, 128) Or _
       ThisWorkbook.Sheets("Log").TextBox3.BackColor <> RGB(128, 255, 128) Then
        Call add_log("Please Check Start Date and End Date."): Exit Sub
    End If
    StartDate = ThisWorkbook.Sheets("Log").TextBox2.Text: EndDate = ThisWorkbook.Sheets("Log").TextBox3.Text
    
    'GET KEY -----------------------------------------------------------------------------------
    'STRUCTURE:
    'KEY: KEY
    'VALUE: ROW AS COLLECTION
    Set KeyDic = get_key(KeyColumn)

    With ThisWorkbook.Sheets("Main")
        
        'GET FIRST EMPTY ROW BASED ON DATE COLUMN ----------------------------------------------
        FirstRow = .Range(DateColumnLetter & "1000000").End(xlUp).Row + 1
        
        Call add_log("Update Data From Row " & FirstRow & Chr(10))
        
        FilePath = Dir(FolderPath & "\*.xlsx")
    
        Do While FilePath <> ""
            
            FilePath = FolderPath & "\" & FilePath
            
            Call add_log("--> " & FilePath)
            
            'GET LastModifiedTime --------------------------------------------------------------
            Set FileObj = fso.GetFile(FilePath)
            
            LastModifiedTime = FileObj.DateLastModified
            
            Call add_log("       Last Modified Time: " & LastModifiedTime)
            
            'GET DATA BASED ON ColumnCount PROVIDED --------------------------------------------
            Data = get_data(FilePath, ColumnCount)
            
            'COUNT UPDATE ROWS -----------------------------------------------------------------
            UpdateCount = 0
            
            For i = 1 To UBound(Data, 1)
            
                If IsDate(Data(i, DateColumn)) Then
                
                    If CDate(Data(i, DateColumn)) >= CDate(StartDate) And _
                       CDate(Data(i, DateColumn)) <= CDate(EndDate) Then
                        
                        Key = Data(i, KeyColumn)
                        
                        If KeyDic.exists(Key) Then
                            
                            'IF DUPLICATE KEY, ADD LOG TO SHOW DUPLICATE ROW -------------------
                            RowStr = ""
                            
                            For Each r In KeyDic(Key)
                            
                                If RowStr = "" Then RowStr = r Else RowStr = RowStr & ", " & r
                                
                            Next
                            
                            Call add_log("             Row: " & i & ", Value: " & Key & " Duplicate On Row " & RowStr)
                            
                        Else:
                            
                            'IF NO DUPLICATE KEY THEN UPDATE -----------------------------------
                            For l = 1 To UBound(Data, 2)
                            
                                .Cells(FirstRow, l) = Data(i, l)
                                
                            Next l
                            
                            Call add_log("             Row: " & i & ", Value: " & Key & " Updated On Row " & FirstRow)
                            
                            'ADD NEW KEY -------------------------------------------------------
                            Set KeyDic(Key) = New Collection
                            
                            KeyDic(Key).Add FirstRow
                            
                            FirstRow = FirstRow + 1
                            
                            UpdateCount = UpdateCount + 1

                        End If
                        
                    End If
                    
                End If
                
            Next i
            
            Call add_log("       " & UpdateCount & " Record(s) Updated In Total." & Chr(10))
            
            FilePath = Dir
            If FilePath = "" Then
                Exit Do
            End If
    
        Loop
        
    End With
    
    Call add_log("Done!")
    
End Sub
