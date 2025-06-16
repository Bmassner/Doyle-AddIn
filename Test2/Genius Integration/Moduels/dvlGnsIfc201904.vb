Module dvlGnsIfc201904
    Dim inventorApp As Inventor.Application
    Public Function dgiG0t0() As Scripting.Dictionary
        Dim dcTree As Scripting.Dictionary
        Dim dcFlat As Scripting.Dictionary
        Dim nm As String
        Dim dt As String

        nm = nuSelAiDoc().GetReply()
        If Len(Trim(nm)) > 0 Then
            With InventorApp.Documents
                dcTree = dgiAiDocClassified(.ItemByName(nm))
                dt = dgiG2f2(dgiG2f1(dcTree))
                If MsgBox(
                "Send this text to the clipoard?" & vbCrLf & vbCrLf & dt,
                vbYesNo + vbQuestion,
                "Send to Clipboard?"
            ) = vbYes Then
                    On Error Resume Next
                    Err.Clear
                    send2clipBdWin10(dt)
                    If Err.Number = 0 Then
                        'MsgBox("PROMPT", vbOKOnly, "TITLE"
                        MsgBox(CStr(Len(dt)) & " characters" & vbCrLf _
                    & "were copied to the clipboard.",
                    vbOKOnly, "COPY SUCCESSFUL!")
                    Else
                        If MsgBox(
                        "Error Code " & Hex$(Err.Number) & ":" _
                        & vbCrLf & Err.Description & vbCrLf _
                        & vbCrLf & "Stop to attempt Debug?",
                        vbYesNo, "COPY FAILED!"
                    ) = vbYes Then
                            Stop
                        End If
                    End If
                    On Error GoTo 0
                Else
                    MsgBox("No data sent to clipboard",
                    vbOKOnly, "COPY CANCELED")
                End If
            End With
        End If
    End Function

    Public Function dgiAiDocClassified(AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ''
        ''  Classify supplied Inventor Document
        ''  by basic Document Type. Retrieve or
        ''  generate sub Dictionary associated
        ''  with Document Type, and reference
        ''  Document there by its Full Name/Path
        ''
        Dim dt As Inventor.DocumentTypeEnum
        Dim fp As String
        'Dim st As String

        If dc Is Nothing Then
            dgiAiDocClassified =
        dgiAiDocClassified(AiDoc,
        New Scripting.Dictionary)
        Else
            With AiDoc
                fp = .FullDocumentName
                dt = .DocumentType
                'st = .SubType
            End With

            If Len(fp) > 0 Then
                With dc
                    If Not .Exists(dt) Then
                        .Add(dt, New Scripting.Dictionary)
                    End If
                    With dcOb(.Item(dt))
                        If Not .Exists(fp) Then .Add(fp, AiDoc)
                    End With
                End With
            End If

            If dt = DocumentTypeEnum.kAssemblyDocumentObject Then
                dgiAiDocClassified = dgiMembersClassified(AiDoc, dc)
            Else
                dgiAiDocClassified = dc
            End If
        End If
    End Function

    Public Function dgiMembersClassified(
    AiDoc As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ''
        ''  Given an Assembly Document,
        ''  categorize its Components.
        ''
        Dim oc As Inventor.ComponentOccurrence
        Dim rt As Scripting.Dictionary

        rt = dc
        With AiDoc.ComponentDefinition
            For Each oc In .Occurrences
                With oc.Definition
                    rt = dgiAiDocClassified(.Document, rt)
                End With
            Next
        End With
        dgiMembersClassified = rt
    End Function

    Public Function dgiFlatListed(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ''
        ''  Flatten Dictionary
        ''  of Dictionaries of
        ''  Inventor Documents
        ''  into one singular
        ''  Dictionary for rescan.
        ''
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim fp As Object
        Dim ct As Long

        rt = New Scripting.Dictionary

        If dc Is Nothing Then
        Else
            With dc
                For Each ky In .Keys
                    With dcOb(.Item(ky))
                        For Each fp In .Keys
                            rt.Add(fp, .Item(fp))
                        Next
                    End With
                Next
            End With
        End If

        dgiFlatListed = rt
    End Function

    Public Function nuSelAiDoc(
        Optional ByVal [Default] As String = "%$#@*&!"
    ) As FmSelectorList
        nuSelAiDoc = nuSelector() _
            .HdrCancel("Cancel Operation?") _
            .HdrNoSelection("No Document Selected!") _
            .HdrOK("Proceed With Operation?") _
            .MsgNoSelection("No changes will be applied to any open Document.") _
            .MsgNoSelection("Do you wish to cancel the Operation?" & vbCrLf & "(Click NO to return to list)") _
            .MsgOK("The following Document(s) will be affected: " & vbCrLf & "%%%" & vbCrLf & "(Click CANCEL to quit with no changes)") _
            .WithList(dcAiDocsVisible().Keys) _
            .SelectIfIn([Default])
        'lsAiDocsVisible()
        'lsWorkbooks()
    End Function
    Public Function AskUser4aiDoc(
            Optional ByVal defaultDoc As Object = Nothing,
            Optional ByVal dc As Object = Nothing
        ) As Object
        Dim nm As String

        If dc Is Nothing Then
            AskUser4aiDoc = AskUser4aiDoc(defaultDoc, dcAiDocsVisible())
        Else
            If defaultDoc Is Nothing Then
                AskUser4aiDoc = AskUser4aiDoc(InventorApp.ActiveDocument, dc)
            Else
                nm = d0g6f0(defaultDoc)
                With dc
                    nm = nuSelAiDoc().WithList(.Keys).SelectIfIn(nm).GetReply()
                    If .Exists(CStr(nm)) Then
                        AskUser4aiDoc = .Item(CStr(nm))
                    Else
                        AskUser4aiDoc = Nothing
                    End If
                End With
            End If
        End If
    End Function

    '
    '
    '
    Public Function dgiG0f0() As Object
        dgiG0f0 = vbEmpty
    End Function

    Public Function dgiG0f1(AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ''
        ''  "Junk" function originally intended to
        ''  collect and categorize Inventor Documents.
        ''  See following functions for preferred approach.
        ''
        If dc Is Nothing Then
            dgiG0f1 = dgiG0f1(AiDoc,
        New Scripting.Dictionary)
        Else
            With AiDoc
                If .DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    If Not dc.Exists(DocumentTypeEnum.kAssemblyDocumentObject) Then
                        dc.Add(DocumentTypeEnum.kAssemblyDocumentObject,
                    New Scripting.Dictionary)
                    End If
                    With dcOb(dc.Item(DocumentTypeEnum.kAssemblyDocumentObject))
                        If Not .Exists(AiDoc.FullDocumentName) Then
                            .Add(AiDoc.FullDocumentName, AiDoc)
                        End If
                    End With
                ElseIf .DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Else
                    Stop
                    If .DocumentType = DocumentTypeEnum.kDesignElementDocumentObject Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kForeignModelDocumentObject Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kNoDocument Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kSATFileDocumentObject Then
                        Stop
                    ElseIf .DocumentType = DocumentTypeEnum.kUnknownDocumentObject Then
                        Stop
                    End If
                End If
            End With
            dgiG0f1 = dc
        End If
    End Function

    Public Function dgiG1f0(
    dc As Scripting.Dictionary
) As Long
        ''
        ''  Return the grand total count
        ''  of entries in all Dictionaries
        ''  within supplied Dictionary.
        ''
        ''  This is meant to check for
        ''  any additions to the collection
        ''  after each processing pass
        ''
        Dim ky As Object
        Dim ct As Long

        ct = 0

        If dc Is Nothing Then
        Else
            With dc
                For Each ky In .Keys
                    With dcOb(.Item(ky))
                        ct = ct + .Count
                    End With
                Next
            End With
        End If

        dgiG1f0 = ct
    End Function

    Public Function dgiG1f1(
    dc As Scripting.Dictionary,
    Optional ck As Long = -1
) As Scripting.Dictionary
        ''
        ''  Build up Dictionary of Inventor
        ''  Part and Assembly Documents
        ''
        Dim AiDoc As Inventor.Document
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim fp As Object
        Dim ct As Long

        If dc Is Nothing Then
            dgiG1f1 = New Scripting.Dictionary
        ElseIf ck < 0 Then
            dgiG1f1 = dgiG1f1(dc, dgiG1f0(dc))
        Else
            rt = dc
            With dgiFlatListed(rt)
                For Each ky In .Keys
                    AiDoc = aiDocument(obOf(.Item(ky)))
                    rt = dgiAiDocClassified(AiDoc, rt)

                    If AiDoc.DocumentType _
                = DocumentTypeEnum.kAssemblyDocumentObject Then
                        rt = dgiMembersClassified(AiDoc, rt)
                    End If
                Next
            End With
            ct = dgiG1f0(dc)

            If ct > ck Then
                dgiG1f1 = dgiG1f1(rt, ct)
            ElseIf ct = ck Then
                dgiG1f1 = rt
            Else
                Stop 'cuz something went wrong
            End If
        End If
    End Function

    Public Function dgiG2f0(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim id As Inventor.Document
        Dim ky As Object
        Dim sb As String
        Dim fp As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                id = aiDocument(.Item(ky))

                With id
                    fp = .FullDocumentName
                    sb = .SubType
                End With

                With rt
                    If Not .Exists(sb) Then
                        .Add(sb, New Scripting.Dictionary)
                    End If
                    With dcOb(.Item(sb))
                        If Not .Exists(fp) Then .Add(fp, id)
                    End With
                End With
            Next
        End With
        dgiG2f0 = rt
    End Function

    Public Function dgiG2f1(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim id As Inventor.Document
        Dim ky As Object
        Dim sb As String
        Dim fp As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                rt.Add(ky, dgiG2f0(dcOb(.Item(ky))))
            Next
        End With
        dgiG2f1 = rt
    End Function

    Public Function dgiG2f2(
    dc As Scripting.Dictionary,
    Optional pfx As String = "",
    Optional dlm As String = "|",
    Optional brk As String = vbCrLf
) As String
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim tx As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                tx = CStr(ky)
                If Len(pfx) > 0 Then tx = pfx & dlm & tx
                If TypeOf .Item(ky) Is Scripting.Dictionary Then
                    rt.Add(dgiG2f2(
                    dcOb(obOf(.Item(ky))),
                    tx, dlm, brk
                ), 0)
                Else 'If TypeOf .Item(ky) Is Inventor.Document Then
                    rt.Add(tx, 0)
                    'Else
                    'Stop
                End If
            Next
        End With
        dgiG2f2 = Join(rt.Keys, brk)
    End Function
    'Debug.Print(dgiG2f2(dgiG2f1(dgiAiDocClassified(InventorApp.Documents.VisibleDocuments(3))))
    'send2clipBd dgiG2f2(dgiG2f1(dgiAiDocClassified(InventorApp.Documents.VisibleDocuments(3))))
    'send2clipBd dgiG2f2(dgiG2f1(dgiAiDocClassified(InventorApp.Documents.ItemByName(nuSelAiDoc().getReply()))))

End Module