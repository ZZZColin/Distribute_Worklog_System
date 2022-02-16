Attribute VB_Name = "Add_Log_Module"
Function add_log(info)

    With ThisWorkbook.Sheets("Log").TextBox1
    
        If .Text = "" Then
        
            .Text = info
            
        Else:
        
            .Text = .Text & Chr(10) & info
            
        End If
        
    End With
    
    DoEvents
    
End Function
