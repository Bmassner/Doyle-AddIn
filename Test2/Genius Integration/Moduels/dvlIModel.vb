Module dvlIModel

    Public Function dImG1f1iPart(md As Inventor.PartDocument) As Inventor.PartDocument
        If md Is Nothing Then
            dImG1f1iPart = Nothing
        Else
            With md.ComponentDefinition
                If .IsiPartFactory Then
                    If .iPartFactory Is Nothing Then
                        Stop
                        dImG1f1iPart = Nothing
                    Else
                        dImG1f1iPart = aiDocPart(
            .iPartFactory.Parent)
                    End If
                ElseIf .IsiPartMember Then
                    If .iPartMember Is Nothing Then
                        Stop
                        dImG1f1iPart = Nothing
                    Else
                        dImG1f1iPart = dImG1f1iPart(
            .iPartMember.ParentFactory.Parent)
                    End If
                    'ElseIf .IsContentMember Then
                    'ElseIf .IsModelStateFactory Then
                    'ElseIf .IsModelStateMember Then
                Else
                    dImG1f1iPart = Nothing
                End If : End With : End If
        'dImG0f1 = "REV[2023.01.19.1046]"
    End Function

    Public Function dImG1f2iPart(md As Inventor.Document) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim PartDoc As Inventor.PartDocument
        Dim ck As Inventor.PartDocument

        rt = New Scripting.Dictionary

        With dcAiPartDocs(dcAiDocComponents(md))
            For Each ky In .Keys
                PartDoc = aiDocPart(.Item(ky))
                If PartDoc Is Nothing Then
                Else
                    ck = dImG1f1iPart(PartDoc)
                    If ck Is Nothing Then
                    Else
                        'ck.File.FullFileName
                        'Stop
                        Debug.Print("") 'Breakpoint Landing
                    End If
                End If
            Next
        End With

        dImG1f2iPart = rt
        ''Debug.Print(ConvertToJson(dImG1f2iPart(aiDocActive()), vbTab)
    End Function

    Public Function dImG0f0() As String
        dImG0f0 = "REV[2023.01.19.1046]"
    End Function
End Module