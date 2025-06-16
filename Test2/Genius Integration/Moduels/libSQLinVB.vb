Module libSQLinVB

    Public Function sqlTextInVBcode(nm As String) As String
        Dim ar As Object

        ar = Split(nm, "'SQL'")
        If UBound(ar) < 1 Then
            sqlTextInVBcode = ""
        Else
            'sqlTextInVBcode = ar(1)
            sqlTextInVBcode = Join(Split(
            ar(1), vbCrLf & "'"
        ), vbCrLf)
        End If
    End Function

    Public Function sqlTextInDict(
    nm As String, dc As Scripting.Dictionary
) As String
        sqlTextInDict = sqlTextInVBcode(
        vbTextOfProcInDict(nm, dc)
    )
    End Function

    Public Function sqlTextInProject(
    nm As String, pj As VBIDE.VBProject
) As String
        sqlTextInProject = sqlTextInDict(
        nm, dcOfVbProcsFlat(pj)
    )
    End Function

End Module