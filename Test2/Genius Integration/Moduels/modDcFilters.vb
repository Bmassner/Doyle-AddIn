Module modDcFilters
    Dim ThisApplication As Inventor.Application
    Public Function dcAiDocsVisible() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim AiDoc As Inventor.Document

        rt = New Scripting.Dictionary
        For Each AiDoc In ThisApplication.Documents.VisibleDocuments
            'rt.Add(aiDoc.FullDocumentName, aiDoc
            rt.Add(d0g6f0(AiDoc), AiDoc)
        Next
        dcAiDocsVisible = rt
    End Function

    Public Sub lsAiDocsVisible()
        Debug.Print(TxDumpLs(dcAiDocsVisible().Keys))
    End Sub

    Public Function dcAiDocsByType(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' Split Dictionary of Inventor Documents
        ' into separate "sub" Dictionaries,
        ' keyed by Document Type
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim AiDoc As Inventor.Document
        Dim tp As Inventor.DocumentTypeEnum
        Dim fn As String
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                AiDoc = aiDocument(.Item(ky))
                With AiDoc
                    tp = .DocumentType
                    fn = .FullFileName
                End With

                With rt
                    If .Exists(tp) Then
                        gp = .Item(tp)
                    Else
                        gp = New Scripting.Dictionary
                        .Add(tp, gp)
                    End If
                End With

                With gp
                    If .Exists(fn) Then
                        Stop
                    Else
                        .Add(fn, AiDoc)
                    End If
                End With
            Next
        End With
        dcAiDocsByType = rt
    End Function
    'Debug.Print(Join(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, ", ")

    Public Function dcAiDocsOfType(
    tp As Inventor.DocumentTypeEnum,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' Retrieve subDictionary for
        ' given Inventor Document type
        '
        With dcAiDocsByType(dc)
            If .Exists(tp) Then
                dcAiDocsOfType = .Item(tp)
            Else
                dcAiDocsOfType = New Scripting.Dictionary
            End If
        End With
    End Function

    Public Function dcAiPartDocs(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        dcAiPartDocs = dcAiDocsOfType(DocumentTypeEnum.kPartDocumentObject, dc)
    End Function
    'Debug.Print(Join(dcAiPartDocs(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbCrLf)

    Public Function dcAiAssyDocs(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        dcAiAssyDocs = dcAiDocsOfType(DocumentTypeEnum.kAssemblyDocumentObject, dc)
    End Function

    Public Function dcOf_iPartFactories(
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim PartDoc As Inventor.PartDocument
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If dc Is Nothing Then
            rt = dcOf_iPartFactories(
            dcAiDocsVisible()
        )
        Else
            rt = New Scripting.Dictionary

            With dcAiPartDocs(dc)
                For Each ky In .Keys
                    PartDoc = aiDocPart(.Item(ky))
                    With PartDoc
                        If .ComponentDefinition.IsiPartFactory Then
                            rt.Add(.FullFileName, PartDoc)
                        End If
                    End With
                Next
            End With
        End If

        dcOf_iPartFactories = rt
    End Function

    Public Function dcOf_iAssyFactories(
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim sm As Inventor.AssemblyDocument
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If dc Is Nothing Then
            rt = dcOf_iAssyFactories(
            dcAiDocsVisible()
        )
        Else
            rt = New Scripting.Dictionary

            With dcAiAssyDocs(dc)
                For Each ky In .Keys
                    sm = aiDocAssy(.Item(ky))
                    With sm
                        If .ComponentDefinition.IsiAssemblyFactory Then
                            rt.Add(.FullFileName, sm)
                        End If
                    End With
                Next
            End With
        End If

        dcOf_iAssyFactories = rt
    End Function

    Public Function dcOf_iAll_Factories(
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If dc Is Nothing Then
            rt = dcOf_iAll_Factories(
            dcAiDocsVisible()
        )
        Else
            rt = dcOf_iPartFactories(dc)

            With dcOf_iAssyFactories(dc)
                For Each ky In .Keys
                    rt.Add(ky, .Item(ky))
                Next
            End With
        End If

        dcOf_iAll_Factories = rt
    End Function

    Public Function dcAiSheetMetal(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim ky As Object
        Dim PartDoc As Inventor.PartDocument
        Dim rt As Scripting.Dictionary

        rt = New Scripting.Dictionary
        With dcAiPartDocs(dc)
            For Each ky In .Keys
                With aiDocPart(.Item(ky))
                    If .DocumentSubType.DocumentSubTypeID = guidSheetMetal Then
                        rt.Add(.FullFileName, .ComponentDefinition.Document)
                    End If
                End With
            Next
        End With
        dcAiSheetMetal = rt
    End Function
    'Debug.Print(Join(dcAiSheetMetal(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbCrLf)
    '
    'Debug.Print(dcAiPartDocs(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count
    'Debug.Print(dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count

    Public Function dcAssyPartsPrimary(
    aiAssy As Inventor.AssemblyDocument
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim oc As Inventor.ComponentOccurrence

        rt = New Scripting.Dictionary
        With aiAssy.ComponentDefinition
            For Each oc In .Occurrences
                With aiDocument(oc.Definition.Document)
                    If Not rt.Exists(.FullDocumentName) Then
                        rt.Add(.FullDocumentName,
                    .Propertys.Parent)
                    End If
                End With
            Next
        End With
        dcAssyPartsPrimary = rt
    End Function

    Public Function dcAiDocsByPtNum(
    dcIn As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim pn As String
        'Dim oc As Inventor.ComponentOccurrence

        rt = New Scripting.Dictionary
        With dcIn 'aiAssy.ComponentDefinition
            For Each ky In .Keys
                With aiDocument(.Item(ky)).Propertys.Item(gnDesign)
                    pn = .Item(pnPartNum).Text
                    If rt.Exists(pn) Then
                        Stop
                    Else
                        rt.Add(pn, .Parent.Parent)
                    End If
                End With
            Next
        End With
        dcAiDocsByPtNum = rt
    End Function
    ' dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: Debug.Print(txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbCrLf & vbTab): Next
    'tx = "":  dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: tx = tx & vbCrLf & ky & vbCrLf & vbTab & txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbCrLf & vbTab): Next: send2clipBd tx:  dc = Nothing

    Public Function dcItemsNotInGenius(
    dcPts As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcItemsNotInGenius --
        '     takes a Dictionary of Items
        '     (keyed by Item/Part Number)
        '     and returns a Dictionary of
        '     Items not yet found in Genius
        '
        ' NOTE: originally designed to take
        '     a Dictionary of Inventor
        '     Documents, it SHOULD be able
        '     to process a Dictionary of
        '     ANY sort of Items keyed
        '     to Item/Part Number
        '
        ' dcPts = dcRemapByPtNum( _
        Dim AiDoc As Inventor.Document
        dcAiDocComponents(AiDoc)


        dcItemsNotInGenius _
    = DcKeysMissing(dcPts, dcOb(
        dcRecSetDcDx4json(DcFromAdoRS(
            CnGnsDoyle().Execute(
            q1g1x2v2(dcPts)
        ))
    ).Item("Item")))
        'Debug.Print(txDumpLs(dcItemsNotInGenius(aiDocActive()).Keys)
    End Function

    Public Function dcAiPartsNotInGenius(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        '
        ' dcAiPartsNotInGenius --
        '     calls dcItemsNotInGenius
        '     against a Dictionary of Items
        '     from supplied Inventor Document
        '     to return a sub of Items
        '     not yet added to Genius
        '

        dcAiPartsNotInGenius _
    = dcItemsNotInGenius(
        dcRemapByPtNum(
        dcAiDocComponents(
        AiDoc
    )))
        'Debug.Print(txDumpLs(dcAiPartsNotInGenius(aiDocActive()).Keys)
    End Function

    Public Function mdf0g0f0(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim bk As String

        bk = vbCrLf & vbTab
        rt = New Scripting.Dictionary
        With dcAiDocsByPtNum(dcAssyPartsPrimary(AiDoc))
            For Each ky In .Keys
                rt.Add(ky & bk & TxDumpLs(
                dcAiDocsByPtNum(dcAssyPartsPrimary(
                    aiDocument(.Item(ky))
                )).Keys, bk
            ), .Item(ky))
            Next
        End With
        mdf0g0f0 = rt
    End Function
    'send2clipBd txDumpLs(mdf0g0f0(ThisApplication.ActiveDocument))

End Module