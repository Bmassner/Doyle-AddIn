Module dvlContentCenterCheck

    Public Function ckCtCtr() As Long
        '
    End Function

    Public Function ccc0g0f0(ck As Inventor.Document) As Inventor.Document
        Dim PartDoc As Inventor.PartDocument

        PartDoc = aiDocPart(ck)
        If PartDoc Is Nothing Then
            ccc0g0f0 = PartDoc
        ElseIf PartDoc.ComponentDefinition.IsContentMember Then
            ccc0g0f0 = PartDoc
        ElseIf PartDoc.Propertys.Count > 4 Then
            ccc0g0f0 = PartDoc
        ElseIf 0 Then
        Else
            ccc0g0f0 = Nothing
        End If
    End Function
End Module