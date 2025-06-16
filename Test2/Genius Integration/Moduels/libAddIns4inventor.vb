Module libAddIns4inventor
    Dim inventorApp As Inventor.Application
    Public Function DcAddIns4inventor(
    Optional app As Inventor.Application = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim adIn As Inventor.ApplicationAddIn

        If app Is Nothing Then
            DcAddIns4inventor = DcAddIns4inventor(inventorApp)
        Else
            rt = New Scripting.Dictionary
            For Each adIn In app.ApplicationAddIns
                rt.Add(adIn.ClassIdString, adIn)
                System.Windows.Forms.Application.DoEvents()
            Next
            DcAddIns4inventor = rt
        End If
    End Function

    Public Function AddIn4inventor(
    clsId As String
) As Inventor.ApplicationAddIn
        With DcAddIns4inventor()
            If .Exists(clsId) Then
                AddIn4inventor = .Item(clsId)
            Else
                AddIn4inventor = Nothing
            End If
        End With
    End Function

    Public Function AddInILogic() As Inventor.ApplicationAddIn
        AddInILogic = AddIn4inventor(
        "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
    )
    End Function

End Module