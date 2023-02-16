Function getRecordByName(name)
    
    strsql = "SELECT * FROM [dbo].[users] " & vbnewline & _
            "WHERE fname = '" & name & "'"

    getRecordByName = strsql
    
End Function

