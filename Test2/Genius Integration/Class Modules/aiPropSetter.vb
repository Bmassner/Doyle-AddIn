Class AiPropSetter
    Private ls() As String

    Public Function PropList() As String()
        PropList = ls
    End Function

    Public Function DcPropsIn(ad As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim PropertySet As Inventor.PropertySet
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        PropertySet = ad.PropertySets.Item(gnCustom)

        If dc Is Nothing Then
            DcPropsIn = DcPropsIn(ad,
            New Scripting.Dictionary
        )
        Else
            If TypeOf ls Is Array Then
                For Each ky In ls
                    InvProperty = aiGetProp(PropertySet, CStr(ky), 1)
                    If InvProperty Is Nothing Then
                        'nothing we can do (as yet?)
                    Else
                        dc.Add(ky, InvProperty)
                    End If
                Next
                DcPropsIn = dc
            ElseIf VarType(ls) = vbString Then
                Stop 'shouldn't wind up here
                DcPropsIn = DcPropsIn(ad, dc)
            Else
                Stop 'or here, either
            End If
        End If
    End Function

    Private Sub Class_Initialize()
        ls = Split("andrew patrick thompson", " ")
    End Sub
End Class