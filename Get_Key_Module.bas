Attribute VB_Name = "Get_Key_Module"
Function get_key(KeyColumn)
    
    Set key_dic = CreateObject("scripting.dictionary")
    
    KeyColumnLetter = Num2Letter(KeyColumn)
    
    With ThisWorkbook.Sheets("Main")
        
        For i = 1 To .Range(KeyColumnLetter & "1000000").End(xlUp).Row
            
            If key_dic.exists(.Cells(i, KeyColumn).Value) Then
            
                key_dic(.Cells(i, KeyColumn).Value).Add i
                
            Else:
            
                Set key_dic(.Cells(i, KeyColumn).Value) = New Collection
                
                key_dic(.Cells(i, KeyColumn).Value).Add i
                
            End If
            
        Next i
        
    End With
    
    Set get_key = key_dic

End Function
