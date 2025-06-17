Module mod0
    Dim ThisApplication As Inventor.Application
    Public Function m0g3f1(rs As ADODB.Record) As Object
        Dim rt As Object
        Dim ar As Object
        Dim BillRow As Object
        Dim mx As Long
        Dim dx As Long
        Dim ct As Long
        Dim fd As Long

        With rs
            .Filter = .Filter
            If .BOF Or .EOF Then
                rt = {{"<NODATA>"}}
            Else
                ct = .Fields.Count - 1
                ar = Split(.GetString(.adClipString, , vbTab, vbVerticalTab), vbVerticalTab)
                mx = UBound(ar) - 1 'Last row is empty/blank
                ReDim rt(mx, ct)
                For dx = 0 To mx
                    BillRow = Split(ar(dx), vbTab)
                    For fd = 0 To ct
                        rt(dx, fd) = BillRow(fd)
                    Next
                Next
            End If
        End With

        m0g3f1 = rt
    End Function

    Public Function m0g3f2(AiDoc As Inventor.Document) As String
        Dim ar As Object
        Dim dx As Long
        Dim ky As Object

        With newFmTest1()
            .AskAbout(AiDoc)
            With .ItemData
                If .Count > 0 Then
                    ar = .Keys
                    For dx = 0 To UBound(ar) ' Each ky In ar
                        'Debug.Print(ky, .Item(ky)
                        ar(dx) = ar(dx) & "=" & .Item(ar(dx))
                    Next
                Else
                    ar = {"<NODATA>"}
                End If
                m0g3f2 = Join(ar, vbCrLf)
            End With
        End With
    End Function

    Public Function m0g2f1(AiDoc As Inventor.Document) As Long
        m0g2f1 = newFmTest0().ft0g0f0(AiDoc.Thumbnail)
    End Function

    Public Function m0g2f2() As Long
        m0g2f2 = m0g2f1(ThisApplication.ActiveDocument)
    End Function

    Public Function m0g2f3() As Long
        Dim AiDoc As Inventor.Document
        'Dim dc As Scripting.Dictionary
        Dim ky As Object

        ' dc = New Scripting.Dictionary
        With New Scripting.Dictionary 'fmTest0
            For Each AiDoc In ThisApplication.Documents
                On Error Resume Next
                '.ft0g0f0 aiDoc.Thumbnail
                If .Exists(AiDoc.FullFileName) Then
                    .Item(AiDoc.FullFileName) = 1 + .Item(AiDoc.FullFileName)
                Else
                    .Add(AiDoc.FullFileName, 1)
                End If

                For Each ky In .Keys
                    If .Item(ky) > 1 Then
                        Debug.Print(.Item(ky), ky)
                    End If
                Next

                On Error GoTo 0
            Next
        End With
    End Function

    Public Function m0g1f0(
    AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim tp As Inventor.DocumentTypeEnum
        Dim ky As String

        If dc Is Nothing Then
            rt = m0g1f0(AiDoc,
        New Scripting.Dictionary)
        Else
            With AiDoc
                tp = .DocumentType
                ky = .FullFileName
            End With

            rt = dc
            If rt.Exists(ky) Then 'we've been here before
                'Don't do anything more, for now

            Else 'we've got a new document to process
                rt.Add(ky, AiDoc)

                If tp = DocumentTypeEnum.kAssemblyDocumentObject Then
                    rt = m0g1f0assy(AiDoc, dc)
                ElseIf tp = DocumentTypeEnum.kPartDocumentObject Then
                Else
                End If
            End If
        End If

        m0g1f0 = rt
    End Function
    ' dc = m0g1f0(ThisApplication.ActiveDocument): For Each ky In dc.Keys: Debug.Print(aiDocument(dc.Item(ky)).Propertys(gnDesign).Item(pnPartNum).Text, aiDocument(dc.Item(ky)).Propertys(gnDesign).Item(pnMaterial).Text: Next

    Public Function m0g1f0part(
    AiDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        If dc Is Nothing Then
            rt = m0g1f0part(AiDoc,
            New Scripting.Dictionary
        )
        Else
            rt = dc
        End If

        m0g1f0part = rt
    End Function

    Public Function m0g1f0assy(
    AiDoc As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim aiOcc As Inventor.ComponentOccurrence

        If dc Is Nothing Then
            rt = m0g1f0assy(AiDoc,
            New Scripting.Dictionary
        )
        Else
            rt = dc
            For Each aiOcc In AiDoc.ComponentDefinition.Occurrences
                rt = m0g1f0(aiOcc.Definition.Document, rt)
            Next
        End If

        m0g1f0assy = rt
    End Function

    Public Function dcAssyComp2A(
    AiDoc As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim invOcc As Inventor.ComponentOccurrence
        Dim tp As Inventor.ObjectTypeEnum

        If dc Is Nothing Then
            rt = dcAssyComp2A(AiDoc, New Scripting.Dictionary)
        Else
            rt = dc
            For Each invOcc In AiDoc.ComponentDefinition.Occurrences
                With invOcc
                    'Remove suppressed and excluded parts from the process
                    'Moved out here from inner checks
                    If .Visible And Not .Suppressed And Not .Excluded Then
                        tp = .Definition.Type

                        If tp <> ObjectTypeEnum.kAssemblyComponentDefinitionObject _
                    And tp <> ObjectTypeEnum.kWeldmentComponentDefinitionObject _
                    Then
                            If tp <> ObjectTypeEnum.kWeldsComponentDefinitionObject Then
                                rt = dcAddAiDoc(.Definition.Document, rt)
                            End If
                        Else 'assembly, check BOM Structure
                            If .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then 'it's purchased
                                rt = dcAddAiDoc(.Definition.Document, rt)
                            ElseIf .BOMStructure = BOMStructureEnum.kNormalBOMStructure Then 'we make it
                                rt = dcAssyComp2A(.SubOccurrences,
                                dcAddAiDoc(.Definition.Document, rt)
                            ) 'NOT forgetting to add THIS document!
                            ElseIf .BOMStructure = BOMStructureEnum.kInseparableBOMStructure Then 'maybe weldment?
                                If tp = ObjectTypeEnum.kWeldmentComponentDefinitionObject Then 'it is
                                    rt = dcAssyComp2A(.SubOccurrences,
                                    dcAddAiDoc(.Definition.Document, rt)
                                )
                                Else 'it's not
                                    Stop 'and see if we can figure out what its type is
                                End If
                            ElseIf .BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then '"phantom" component
                                'Gather its components, but NOT the document itself
                                rt = dcAssyComp2A(.SubOccurrences, rt)
                            Else 'not sure what we've got
                                Stop 'and have a look at it
                            End If
                        End If 'part or assembly
                    End If
                End With
            Next
        End If

        dcAssyComp2A = New Scripting.Dictionary
    End Function

    Public Function dcAssyCompStops(
    Occurences As Inventor.ComponentOccurrences,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional dcStops As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ' Traverse the assembly,
        ' including any/all subassemblies,
        ' and collect all parts to be processed.
        Dim rt As Scripting.Dictionary
        Dim invOcc As Inventor.ComponentOccurrence
        Dim tp As Inventor.ObjectTypeEnum

        If dc Is Nothing Then
            rt = dcAssyCompStops(Occurences, New Scripting.Dictionary, dcStops)
        Else
            rt = dc
            For Each invOcc In Occurences
                With invOcc
                    If dcStops.Exists(aiDocument(.Definition.Document
                ).Propertys.Item(gnDesign).Item(pnPartNum).Text
                ) Then
                        'Stop
                    End If

                    'Remove suppressed and excluded parts from the process
                    'Moved out here from inner checks
                    If .Visible And Not .Suppressed And Not .Excluded Then
                        tp = .Definition.Type

                        MsgBox(Join({
                        "TYPE: " & tp,
                        "VISIBLE: " & .Visible,
                        "NAME: " & .Name,
                        "Suboccurence: " & .SubOccurrences.Count,
                        "Occurence Type: " & .Definition.Occurrences.Type,
                        "BOMStructure: " & .BOMStructure _
                    , vbCrLf}))

                        If tp <> ObjectTypeEnum.kAssemblyComponentDefinitionObject _
                    And tp <> ObjectTypeEnum.kWeldmentComponentDefinitionObject _
                    Then
                            '(moved suppression/exclusion check OUTSIDE)
                            If tp <> ObjectTypeEnum.kWeldsComponentDefinitionObject Then

                                ' rt = dcAddAiDoc(aiDocument(.Definition.Document), rt)
                                ' Recasting by aiDocument not likely necessary here.
                                ' Revised to following:
                                rt = dcAddAiDoc(.Definition.Document, rt)

                            End If 'inVisible, suppressed, excluded or Welds

                        Else 'assembly, check BOM Structure
                            If .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then 'it's purchased
                                'Just add it to the Dictionary
                                rt = dcAddAiDoc(.Definition.Document, rt)
                            ElseIf .BOMStructure = BOMStructureEnum.kNormalBOMStructure Then 'we make it
                                'Gather its components
                                rt = dcAssyCompStops(.SubOccurrences,
                                dcAddAiDoc(.Definition.Document, rt),
                                dcStops
                            ) 'NOT forgetting to add THIS document!
                            ElseIf .BOMStructure = BOMStructureEnum.kInseparableBOMStructure Then 'maybe weldment?
                                If tp = ObjectTypeEnum.kWeldmentComponentDefinitionObject Then 'it is
                                    'Treat it like an assembly
                                    rt = dcAssyCompStops(.SubOccurrences,
                                    dcAddAiDoc(.Definition.Document, rt),
                                    dcStops
                                )
                                Else 'it's not
                                    Stop 'and see if we can figure out what its type is
                                End If
                            ElseIf .BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then '"phantom" component
                                'Gather its components, but NOT the document itself
                                rt = dcAssyCompStops(.SubOccurrences, rt, dcStops)
                            Else 'not sure what we've got
                                Stop 'and have a look at it
                            End If
                        End If 'part or assembly
                    End If
                End With
            Next
        End If
        dcAssyCompStops = rt
    End Function

    Public Function dcStopAt(tx As String,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        If dc Is Nothing Then
            rt = dcStopAt(tx, New Scripting.Dictionary)
        Else
            rt = dc
        End If

        With rt
            If Not .Exists(tx) Then .Add(tx, tx)
        End With

        dcStopAt = rt
    End Function

    Public Function m0g0f0() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim tx As Object

        rt = New Scripting.Dictionary
        For Each tx In {
        "29-197A", "29-197B", "29-633",
        "29-634", "29-637", "29-638",
        "29-647", "29-648", "29-650",
        "29-651", "29-652", "HD182"}

            rt = dcStopAt(tx, rt)
        Next
        m0g0f0 = rt
    End Function

    Public Function m0g4f0(dcIn As Scripting.Dictionary) As Scripting.Dictionary
        ''  g4 currently reserved for development
        ''  relating to identification of purchased parts
        ''  by reference to Vault file path, as well as
        ''  BOM Structure and Cost Center tings
        ''
        ''  f0 separates members of incoming Dictionary
        ''  into design subgroups, including "doyle" and "purchased",
        ''  as indicated by the fourth component of each Vault path
        ''
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim ky As Object
        Dim ls As Object
        Dim mx As Long
        Dim dx As Long
        Dim g0 As String
        Dim g1 As String
        Dim fp As String
        Dim fn As String
        Dim rf As String

        rt = New Scripting.Dictionary
        With dcIn
            For Each ky In .Keys
                ls = Split(ky, "\")
                mx = UBound(ls)

                fn = ls(mx)
                g0 = ls(3)
                g1 = ls(4)
                fp = ""
                For dx = 5 To mx - 1
                    fp = fp & "\" & ls(dx)
                Next
                fp = Mid(fp, 2)
                rf = Join({fn, g1, fp, vbTab})

                With rt
                    'Stop
                    If .Exists(g0) Then
                        sd = .Item(g0)
                    Else
                        sd = New Scripting.Dictionary
                        .Add(g0, sd)
                    End If

                    With sd
                        If .Exists(rf) Then
                            Stop
                        Else
                            .Add(rf, dcIn.Item(ky))
                        End If
                    End With
                End With
            Next
        End With
        m0g4f0 = rt
    End Function

    Public Function m0g4f1(dcIn As Scripting.Dictionary) As Scripting.Dictionary
        ''  g4 currently reserved for development
        ''  relating to identification of purchased parts
        ''  by reference to Vault file path, as well as
        ''  BOM Structure and Cost Center tings
        ''
        ''  f1 scans documents in the "purchased" design group
        ''  for unexpected tings, specifically:
        ''  BOMStructure should be
        ''      BOMStructureEnum.kPurchasedBOMStructure (51973 / 0xCB05)
        ''  Design Tracking Property "Cost Center"
        ''      should be either "D-PTS" or "R-PTS"
        ''
        ''  Any documents not matching these parameters
        ''  are dropped into one or both subDictionaries
        ''  in the returned Dictionary, according to issue
        ''
        Dim rt As Scripting.Dictionary
        Dim rtBom As Scripting.Dictionary
        Dim rtFam As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim ivDoc As Inventor.Document
        Dim CpDef As Inventor.ComponentDefinition
        Dim FamilyProp As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        rtBom = New Scripting.Dictionary
        rtFam = New Scripting.Dictionary
        rt.Add("bom", rtBom)
        rt.Add("fam", rtFam)

        sd = dcIn.Item("purchased")
        With sd
            For Each ky In .Keys
                ivDoc = .Item(ky)
                CpDef = aiCompDefOf(ivDoc)

                If CpDef Is Nothing Then
                    Stop
                Else
                    With CpDef
                        If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                            rtBom.Add(ky, ivDoc)
                        End If

                        FamilyProp = aiDocument(.Document
                        ).Propertys.Item(
                        gnDesign).Item(pnFamily
                    )

                        If FamilyProp.Text <> "D-PTS" Or FamilyProp.Text <> "R-PTS" Then
                            rtFam.Add(ky, ivDoc)
                        End If
                    End With
                End If
            Next
        End With
        m0g4f1 = rt
    End Function
    'Debug.Print(Join(m0g4f1(m0g4f0(dcAssyDocComponents(ThisApplication.ActiveDocument))).Item("fam").Keys, vbCrLf)

    Public Function m0g4f1fixBom(
    dcIn As Scripting.Dictionary,
    Optional RiverView As Long = 0
) As Scripting.Dictionary
        ''  g4 currently reserved for development
        ''  relating to identification of purchased parts
        ''  by reference to Vault file path, as well as
        ''  BOM Structure and Cost Center tings
        ''
        ''  f1fixBom purports to fix incorrect BOM
        ''  Structure tings in purchased parts
        ''
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        'Dim ivDoc As Inventor.Document
        Dim CpDef As Inventor.ComponentDefinition
        Dim ky As Object

        sd = dcIn.Item("bom")
        rt = New Scripting.Dictionary
        With sd
            For Each ky In .Keys
                CpDef = aiCompDefOf(.Item(ky))
                If CpDef Is Nothing Then
                    Stop
                    rt.Add(ky, 0)
                Else
                    With CpDef
                        .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                        rt.Add(ky, IIf(.BOMStructure _
                        = BOMStructureEnum.kPurchasedBOMStructure,
                    1, 0))
                    End With
                End If
            Next
        End With
        m0g4f1fixBom = rt
    End Function

    Public Function m0g4f1fixFam(
    dcIn As Scripting.Dictionary,
    Optional RiverView As Long = 0
) As Scripting.Dictionary
        ''  g4 currently reserved for development
        ''  relating to identification of purchased parts
        ''  by reference to Vault file path, as well as
        ''  BOM Structure and Cost Center tings
        ''
        ''  f1fixFam purports to fix incorrect
        ''  Family tings in purchased parts
        ''
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim ivDoc As Inventor.Document
        Dim CpDef As Inventor.ComponentDefinition
        Dim ky As Object
        Dim txFam As String

        txFam = IIf(RiverView, "R", "D") & "-PTS"
        sd = dcIn.Item("fam")

        rt = New Scripting.Dictionary
        With sd
            For Each ky In .Keys
                ivDoc = .Item(ky)
                With ivDoc.Propertys.Item(gnDesign).Item(pnFamily)
                    .Text = txFam
                    rt.Add(ky, IIf(.Text = txFam, 1, 0))
                End With

                CpDef = aiCompDefOf(.Item(ky))
                If CpDef Is Nothing Then
                    Stop
                    rt.Add(ky, 0)
                Else
                    With CpDef
                        .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                        rt.Add(ky, IIf(.BOMStructure _
                        = BOMStructureEnum.kPurchasedBOMStructure,
                    1, 0))
                    End With
                End If
            Next
        End With
        m0g4f1fixFam = rt
    End Function
End Module