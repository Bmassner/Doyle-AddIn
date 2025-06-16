Class AiPropLink
    Private dc As Scripting.Dictionary

    Private Sub Class_Initialize()
        dc = New Scripting.Dictionary
    End Sub

    Private Sub Class_Terminate()
        dc.RemoveAll()
        dc = Nothing
    End Sub

    Public Function PrepFrom(AiDoc As Inventor.Document) As AiPropLink
        Dim PropertySet As Inventor.PropertySet
        Dim InvProperty As Inventor.Property
        Dim PropertySetName As String
        Dim PropertyName As String

        With AiDoc
            For Each PropertySet In .PropertySets
                PropertySetName = PropertySet.Name
                For Each InvProperty In PropertySet
                    PropertyName = InvProperty.Name
                    If dc.Exists(PropertyName) Then
                        Stop
                    Else
                        dc.Add(PropertyName, PropertySetName)
                    End If
                Next
            Next
        End With
        PrepFrom = Me
    End Function
End Class
