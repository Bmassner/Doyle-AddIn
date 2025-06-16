Module lib1
    Public Function Repeat(
    Count As Long, Text As String
) As String
        Repeat = Replace(Space$(Count), " ", Text)
    End Function

    Public Function txBlk(
    Lines As Long, Chars As Long,
    Optional Use As String = "+"
    ) As String
        txBlk = Mid$(Repeat(
        Lines, vbCrLf _
        & New String(CChar(Use), Chars)
    ), 1 + Len(vbCrLf))
    End Function

    'Public Sub MakeActivePurchased()
    '    Dim md As Inventor.Document
    '    Dim ck As VbMsgBoxResult

    '    md = InventorApp.ActiveDocument
    '    If md Is ThisDocument Then
    '        ck = vbNo
    '    Else
    '        ck = mkAiDocPurchased(md)
    '    End If

    '    If ck = vbOK Then
    '        ck = MsgBox(Join({
    '        "Model BOM Structure",
    '        "now Purchased."
    '    }, vbCrLf),
    '        vbOKOnly + vbInformation,
    '        "Success!"
    '    )
    '    ElseIf ck = vbNo Then
    '        ck = MsgBox(Join({
    '        "Document is not",
    '        "a valid Model.",
    '        "",
    '        "Please select a",
    '        "Part or Assembly."
    '    }, vbCrLf),
    '        vbOKOnly + vbExclamation,
    '        "No Model"
    '    )
    '    ElseIf ck = vbAbort Then
    '        ck = MsgBox(Join({
    '        "Failed to update",
    '        "model's BOM Structure!",
    '        "",
    '        "Check for locks",
    '        "or other issues."
    '    }, vbCrLf),
    '        vbOKOnly + vbCritical,
    '        "Change Failed!"
    '    )
    '    Else
    '        ck = MsgBox(Join({
    '        "Change Operation returned",
    '        "unexpected result code.",
    '        "",
    '        "Please review model status."
    '    }, vbCrLf),
    '        vbOKOnly + vbQuestion,
    '        "Result Unknown"
    '    )
    '    End If
    'End Sub

    Public Function mkAiDocPurchased(
    AiDoc As Inventor.Document
) As VbMsgBoxResult
        Dim ck As VbMsgBoxResult

        If TypeOf AiDoc Is Inventor.PartDocument Then
            ck = mkAiPartPurchased(AiDoc)
        ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
            ck = mkAiAssyPurchased(AiDoc)
        Else
            ck = vbNo
        End If

        mkAiDocPurchased = ck
    End Function

    Public Function mkAiPartPurchased(
    AiDoc As Inventor.PartDocument
) As VbMsgBoxResult
        If AiDoc Is Nothing Then
            mkAiPartPurchased = vbNo
        Else
            With AiDoc.ComponentDefinition
                On Error Resume Next
                Err.Clear()
                .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                If Err.Number = 0 Then
                    mkAiPartPurchased = vbOK
                Else
                    mkAiPartPurchased = vbAbort
                End If
                On Error GoTo 0
            End With : End If
    End Function

    Public Function mkAiAssyPurchased(
    AiDoc As Inventor.AssemblyDocument
) As VbMsgBoxResult
        If AiDoc Is Nothing Then
            mkAiAssyPurchased = vbNo
        Else
            With AiDoc.ComponentDefinition
                On Error Resume Next
                Err.Clear()
                .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                If Err.Number = 0 Then
                    mkAiAssyPurchased = vbOK
                Else
                    mkAiAssyPurchased = vbAbort
                End If
                On Error GoTo 0
            End With : End If
    End Function

    Public Function dcTemplate0A(
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        If dc Is Nothing Then
            rt = dcTemplate0A(
           New Scripting.Dictionary
       )
        Else
            rt = dc
        End If

        dcTemplate0A = rt
    End Function

    Public Function send2clipBd_OBSOLETE(src As Object) As Object
        With New MSForms.DataObject
            .SetText(src)
            .PutInClipboard
        End With
        send2clipBd_OBSOLETE = src
    End Function

End Module