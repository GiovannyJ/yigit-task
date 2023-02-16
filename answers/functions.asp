Function getRecordByName(name)
    
    strsql = "SELECT * FROM [dbo].[users] " & vbnewline & _
            "WHERE fname = '" & name & "'"

    getRecordByName = strsql
    
End Function


Function getEilgible()
    strsql = "SELECT * FROM [dbo].[users]" & vbnewline & _
             "WHERE eligible=1 AND deleted = 0"

    getEilgible = strsql
End Function


Function getNonEilgible()
    strsql = "SELECT * FROM [dbo].[users]" & vbnewline & _
             "WHERE eligible=0 AND deleted = 0"

    getNonEilgible = strsql
End Function

    