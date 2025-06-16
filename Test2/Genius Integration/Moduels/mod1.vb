Module mod1
    Dim inventorApp As Inventor.Application
    Public Function m1g0f0() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim gt0 As Long
        Dim eq0 As Long

        rt = New Scripting.Dictionary
        With dcAiSheetMetal(dcAiDocsByType(
        dcAssyDocComponents(inventorApp.ActiveDocument)
    ))
            gt0 = 0
            eq0 = 0
            For Each ky In .Keys
                With aiDocPart(.Item(ky))
                    With vcChkFlatPat(aiCompDefShtMetal(
                .ComponentDefinition
            ))
                        If Math.Abs(.X * .Y * .Z) > 0 Then
                            Debug.Print(.X * .Y * .Z, ky)
                            'Stop
                            gt0 = 1 + gt0
                        Else
                            'Stop
                            eq0 = 1 + eq0
                        End If
                        'If .HasFlatPattern Then
                        'With .FlatPattern
                        'If .Features.Count > 0 Then
                        'gt0 = 1 + gt0
                        'rt.Add(ky, .Document
                        'Else
                        'eq0 = 1 + eq0
                        ''Debug.Print(.Width * .Length
                        ''Stop
                        'End If
                        'End With
                        'Else
                        ''Debug.Print(aiDocument(.Document).FullFileName
                        ''Stop
                        'End If
                    End With
                End With
            Next
            Debug.Print(gt0, eq0)
        End With
        m1g0f0 = rt
    End Function

    Public Function m1g0f1() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim gt0 As Long
        Dim eq0 As Long

        rt = New Scripting.Dictionary
        With dcAssyDocComponents(
        inventorApp.ActiveDocument, , 1
    )
            gt0 = 0
            eq0 = 0
            For Each ky In .Keys
                With aiDocument(.Item(ky))
                    If .DocumentInterests.HasInterest(
                    guidDesignAccl
                ) Then
                        Stop
                    End If
                End With
            Next
            Debug.Print(gt0, eq0)
        End With
        m1g0f1 = rt
    End Function

    Public Function vcDiaBox(bx As Inventor.Box) As Inventor.Vector
        ''  Given a Box Object
        ''  containing diametrically opposed
        ''  MinPoint and MaxPoint Objects,
        ''  return the Vector from Min to Max
        With bx
            vcDiaBox = .MinPoint.VectorTo(.MaxPoint)
        End With
    End Function

    Public Function vc2flatPat(
    df As Inventor.SheetMetalComponentDefinition
) As Inventor.Vector
        ''  From an Inventor Sheet Metal Component Definition,
        ''  obtain a Vector representing the translation
        ''  of the Folded Model's bounding box diagonal
        ''  to that of the Flat Pattern.
        ''
        Dim rt As Inventor.Vector

        With df
            rt = vcDiaBox(.RangeBox)
            If .HasFlatPattern Then
                rt.SubtractVector(vcDiaBox(.FlatPattern.RangeBox))
            Else
                rt.SubtractVector(rt)
            End If
        End With
        ''
        ''  If they're the same point, this vector should
        ''  have zero length, however, this is NOT proof
        ''  positive of an invalid Flat Pattern. A valid
        ''  flat piece, with no folds, should produce
        ''  the same result.
        ''
        ''  A good follow-up check would probably be to
        ''  compare the Flat Pattern diagonal vector's
        ''  Z component to the model's Thickness

        vc2flatPat = rt
    End Function
    'Debug.Print(vc2flatPat(aiCompDefShtMetal(aiDocPart(InventorApp.Documents(2)).ComponentDefinition)).Length

    Public Function vcCubicThickness(
    df As Inventor.SheetMetalComponentDefinition
) As Inventor.Vector
        Dim hk As Double

        hk = df.Thickness.Text
        With inventorApp.TransientGeometry
            vcCubicThickness = .CreateVector(hk, hk, hk)
        End With
    End Function

    Public Function vcChkFlatPat(
    df As Inventor.SheetMetalComponentDefinition
) As Inventor.Vector
        ''  From an Inventor Sheet Metal Component Definition,
        ''  subtract a vector of cubic thickness from the
        ''  diagonal vector of either its Flat Pattern,
        ''  if available, or otherwise, the Folded Model.
        ''
        Dim rt As Inventor.Vector

        With df
            If .HasFlatPattern Then
                rt = vcDiaBox(.FlatPattern.RangeBox)
                'With .Thickness
                'rt.SubtractVector InventorApp.TransientGeometry.CreateVector(.Text, .Text, .Text)
                'With InventorApp.TransientGeometry.CreateVector(.Text, .Text, .Text)
                '.SubtractVector rt
                'If (.X * .Y * .Z) <> 0 Then
                'Debug.Print(.X * .Y * .Z
                'Stop
                'End If
                'End With
                'End With
                'rt.SubtractVector vcDiaBox(.RangeBox)
            Else
                rt = vcDiaBox(.RangeBox)
            End If
        End With
        rt.SubtractVector(vcCubicThickness(df))
        ''
        ''  If the model is a valid sheet metal part,
        ''  one of the dimensions of its flat pattern's
        ''  bounding box diagonal should either equal
        ''  the defined sheet metal thickness,
        ''  or fall very close. At least, in theory.
        ''
        ''  Plan at this point is to try to determine
        ''  just how often this bears out.
        ''  While this HAS failed one pretest, the Flat
        ''  Pattern of that model includes features;
        ''  a relatively infrequent occurrence, and
        ''  quite possibly one that can throw off
        ''  the boundaries.

        vcChkFlatPat = rt
    End Function
    'Debug.Print(vcChkFlatPat(aiCompDefShtMetal(aiDocPart(InventorApp.Documents(2)).ComponentDefinition)).Length

    Public Function m1tst0() As Object
        Debug.Print(iFacAssy(aiCompDefAssy(aiDocAssy(aiCompDefAssy(aiDocAssy(inventorApp.ActiveDocument).ComponentDefinition).Occurrences(1).Definition.Document).ComponentDefinition)))
    End Function

    Public Function m1tst1() As Object
        Dim ky As Object
        Dim InvProperty As Inventor.Property
        'Dim bm As Inventor.BOMStructureEnum

        For Each ky In Filter(dcAssyCompAndSub(
        aiDocDefAssy(inventorApp.ActiveDocument).Occurrences
    ).Keys, "(DC)")
            With aiDocPart(inventorApp.Documents.ItemByName(ky))
                InvProperty = .Propertys("Design Tracking Properties").Item("Cost Center")
                With .ComponentDefinition
                    If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                        .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                        InvProperty.Text = "D-PTS"
                        Debug.Print(InvProperty.Text, .BOMStructure, ky)
                    End If
                End With
                '.ComponentDefinition
                '.ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure) & "|" & ky
            End With
            'Debug.Print((aiDocPart(InventorApp.Documents.ItemByName((ky))).ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure) & "|" & ky
        Next
    End Function

    Public Function iFacAssy(
    ob As Inventor.AssemblyComponentDefinition
) As Inventor.iAssemblyFactory
        With ob
            If .iAssemblyFactory Is Nothing Then
                If .iAssemblyMember Is Nothing Then
                    iFacAssy = Nothing
                Else
                    iFacAssy = .iAssemblyMember.ParentFactory
                End If
            Else
                iFacAssy = .iAssemblyFactory
            End If
        End With
    End Function

    Public Function aiOccDoc(
    ob As Inventor.ComponentOccurrence
) As Inventor.Document
        aiOccDoc = ob.Definition.Document
    End Function

    Public Function aiDocDefAssy(
    ob As Inventor.AssemblyDocument
) As Inventor.AssemblyComponentDefinition
        aiDocDefAssy = ob.ComponentDefinition
    End Function

    Public Function dcAssyComponentsImmediate(
    ob As Inventor.AssemblyDocument
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim oc As Inventor.ComponentOccurrence
        Dim cp As Inventor.Document
        Dim fn As String

        rt = New Scripting.Dictionary
        For Each oc In ob.ComponentDefinition.Occurrences
            'aiDocument(oc.Definition.Document).FullFileName
            cp = oc.Definition.Document
            fn = cp.FullFileName
            With rt
                If Not .Exists(fn) Then .Add(fn, cp)
            End With
        Next
        dcAssyComponentsImmediate = rt
    End Function
    'For Each pn In {pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnArea, pnLength, pnWidth, pnThickness): Debug.Print(pn & "=" & aiDocument(InventorApp.ActiveDocument).Propertys.Item(gnCustom).Item((pn)).Text: Next

    Public Function iSyncPartFactory(PartDesc As Inventor.PartDocument) As Long
        ''
        ''  Backport iPart Member Properties to parent Factory
        ''
        Dim dcCols As Scripting.Dictionary
        Dim dcRows As Scripting.Dictionary
        Dim tr As Inventor.iPartTableRow
        Dim PropertySet As Inventor.Property
        Dim InvProperty As Inventor.Property
        Dim rt As Long
        Dim ck As VbMsgBoxResult

        rt = 0
        With PartDesc.ComponentDefinition
            PropertySet = aiDocument(.Document).Propertys.Item(gnCustom)
            If .iPartMember Is Nothing Then
                'Stop
            Else
                With .iPartMember
                    dcRows = dcIPartTbRows(.ParentFactory.TableRows)
                    If dcRows.Exists(PartDesc.DisplayName) Then
                        dcCols = dcIPartTbCols(.ParentFactory.TableColumns)

                        On Error Resume Next
                        Err.Clear()
                        tr = .Row

                        ' REV[2022.03.23.1624]
                        '     adding "pickup" attempt to capture
                        '     and recover from errors encountered
                        '     in trying to retrieve .Row directly
                        If Err.Number = 0 Then
                        Else
                            tr = dcRows.Item(PartDesc.DisplayName)
                            If tr.MemberName = PartDesc.DisplayName Then
                                Err.Clear()
                            ElseIf tr.PartName = PartDesc.DisplayName Then
                                Err.Clear()
                            Else
                                Stop
                            End If
                        End If

                        If Err.Number = 0 Then
                            Dim i As Integer
                            For i = 1 To PropertySet.Count
                                InvProperty = PropertySet.Item(i)
                                If dcCols.Exists(InvProperty.Name) Then
                                    Debug.Print(InvProperty.Name & " [" & InvProperty.Text & "]: ")
                                    With tr.Item(dcCols.Item(InvProperty.Name))
                                        Debug.Print("  " & .Text)
                                        If InvProperty.Text = .Text Then
                                            'Stop 'No change necessary
                                            Debug.Print(" (NO CHANGE)")
                                        Else
                                            On Error Resume Next
                                            .Text = InvProperty.Text
                                            If Err.Number = 0 Then
                                                rt = 1 + rt

                                                '' The update invalidated the object
                                                '' We'll have to grab it again
                                                With tr.Item(dcCols.Item(InvProperty.Name))
                                                    Debug.Print(" -> " & .Text)
                                                End With
                                            Else
                                                Debug.Print(" <!!ERROR!!> Couldn't Change")
                                                'Debug.Print(Err.Number, Err.Description
                                            End If

                                            On Error GoTo 0
                                        End If
                                    End With
                                Else
                                    'Stop
                                End If
                            Next i

                            Debug.Print("")  'Breakpoint Landing
                        Else
                            Debug.Print("=== CAN'T SYNC IFACTORY ===")
                            Debug.Print("   Failed to access Row")
                            Debug.Print("   for Member " & PartDesc.DisplayName)
                            Debug.Print("   of Factory " & aiDocument(.ParentFactory.Parent).DisplayName)
                            Debug.Print("=== PLEASE CHECK PARENT ===")
                            Debug.Print("====== FACTORY TABLE ======")
                            Debug.Print("Error 0x" & Hex$(Err.Number) & "(", CStr(Err.Number) & ")")
                            Debug.Print("    " & Err.Description)
                            Debug.Print("===========================")
                            ck = MsgBox(Join({"" _
                            & "iPart Member " & PartDesc.DisplayName _
                            , "in Factory" & aiDocument(.ParentFactory.Parent).DisplayName _
                            , "could not be accessed for updates." _
                            , "" _
                            , "Its Row might still be present" _
                            , "but somehow unavailable." _
                            , "" _
                            , "Please review iPart Factory."
                        }, vbCrLf), vbOKCancel,
                            "ERROR ACCESSING MEMBER ROW!"
                        )
                            If ck = vbCancel Then
                                Stop
                            End If
                        End If

                        On Error GoTo 0
                        Debug.Print("")  'Breakpoint Landing
                    Else
                        Debug.Print("==== CAN'T FIND MEMBER ====")
                        Debug.Print("   Failed to locate Row")
                        Debug.Print("   for Member " & PartDesc.DisplayName)
                        Debug.Print("   of Factory " & aiDocument(.ParentFactory.Parent).DisplayName)
                        Debug.Print("=== PLEASE CHECK PARENT ===")
                        Debug.Print("====== FACTORY TABLE ======")
                        'Debug.Print("Error 0x" & Hex$(Err.Number) & "(", CStr(Err.Number) & ")"
                        'Debug.Print("    " & Err.Description
                        'Debug.Print("==========================="
                        ck = MsgBox(Join({"" _
                        & "iPart Member " & PartDesc.DisplayName _
                        , "could not be located in Factory" _
                        , aiDocument(.ParentFactory.Parent).DisplayName _
                        , "" _
                        , "Its Row might have been removed" _
                        , "or separated from the main table." _
                        , "" _
                        , "Please review iPart Factory Table."
                    }, vbCrLf), vbOKCancel,
                        "WARNING!! MEMBER NOT FOUND!"
                    )
                        If ck = vbCancel Then
                            Stop
                        End If
                    End If
                End With
            End If
        End With
    End Function

    Public Function iSyncAssyFactory(PartDesc As Inventor.AssemblyDocument) As Long
        ''
        ''  Backport iPart Member Properties to parent Factory
        ''
        Dim dcCols As Scripting.Dictionary
        Dim dcRows As Scripting.Dictionary
        Dim tr As Inventor.iAssemblyTableRow
        Dim PropertySet As Inventor.PropertySet
        Dim InvProperty As Inventor.Property
        Dim rt As Long
        Dim i As Integer

        rt = 0
        With PartDesc.ComponentDefinition
            PropertySet = aiDocument(.Document).Propertys.Item(gnCustom)
            If .iAssemblyMember Is Nothing Then
                'Stop
            Else
                With .iAssemblyMember
                    dcRows = dcIAssyTbRows(.ParentFactory.TableRows)
                    If dcRows.Exists(PartDesc.DisplayName) Then
                        dcCols = dcIAssyTbCols(.ParentFactory.TableColumns)
                        tr = .Row
                        For i = 1 To PropertySet.Count
                            InvProperty = PropertySet.Item(i)
                            If dcCols.Exists(InvProperty.Name) Then
                                Debug.Print(InvProperty.Name & " [" & InvProperty.Text & "]: ")
                                With tr.Item(dcCols.Item(InvProperty.Name))
                                    Debug.Print("  " & .Text)
                                    If InvProperty.Text = .Text Then
                                        'Stop 'No change necessary
                                        Debug.Print(" (NO CHANGE)")
                                    Else
                                        On Error Resume Next
                                        .Text = InvProperty.Text
                                        If Err.Number = 0 Then
                                            rt = 1 + rt

                                            '' The update invalidated the object
                                            '' We'll have to grab it again
                                            With tr.Item(dcCols.Item(InvProperty.Name))
                                                Debug.Print(" -> " & .Text)
                                            End With
                                        Else
                                            Debug.Print(" <!!ERROR!!> Couldn't Change")
                                        End If

                                        On Error GoTo 0
                                    End If
                                End With
                            Else
                                'Stop
                            End If
                        Next i
                    Else
                        Stop
                    End If
                End With
            End If
        End With
    End Function

    Public Function dcColumnsIPart(
    PartDesc As Inventor.PartDocument
) As Scripting.Dictionary
        ''  Retrieve Dictionary of iPart Factory Table Columns
        ''  If supplied Part Document is NOT an iPart Factory
        ''  OR Member, returned Dictionary will be empty
        ''
        With PartDesc.ComponentDefinition
            If .iPartMember Is Nothing Then
                If .iPartFactory Is Nothing Then
                    dcColumnsIPart = New Scripting.Dictionary
                Else
                    dcColumnsIPart = dcIPartTbCols(
                    .iPartFactory.TableColumns
                )
                End If
            Else
                dcColumnsIPart = dcIPartTbCols(
                .iPartMember.ParentFactory.TableColumns
            )
            End If
        End With
    End Function

    Public Function dcColumnsIAssy(
    PartDesc As Inventor.AssemblyDocument
) As Scripting.Dictionary
        ''  Retrieve Dictionary of iAssembly Factory Table Columns
        ''  If supplied Assembly Document is NOT an iAssembly Factory
        ''  OR Member, returned Dictionary will be empty
        ''
        With PartDesc.ComponentDefinition
            If .iAssemblyMember Is Nothing Then
                If .iAssemblyFactory Is Nothing Then
                    dcColumnsIAssy = New Scripting.Dictionary
                Else
                    dcColumnsIAssy = dcIAssyTbCols(
                    .iAssemblyFactory.TableColumns
                )
                End If
            Else
                dcColumnsIAssy = dcIAssyTbCols(
                .iAssemblyMember.ParentFactory.TableColumns
            )
            End If
        End With
    End Function

    Public Function dcIPartTbCols(
    ls As Inventor.iPartTableColumns
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim it As Inventor.iPartTableColumn

        rt = New Scripting.Dictionary
        On Error Resume Next
        For Each it In ls
            rt.Add(it.DisplayHeading, it.Index)
            If Err.Number Then
                rt.Add(it.FormattedHeading, it.Index)
                Err.Clear()
            End If
        Next
        On Error GoTo 0
        dcIPartTbCols = rt
    End Function

    Public Function dcIPartTbRows(
    ls As Inventor.iPartTableRows
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim it As Inventor.iPartTableRow
        Dim ck As Long

        rt = New Scripting.Dictionary

        ' REV[2022.08.05.1003]
        ' added error trap code to more gracefully handle
        ' errors accessing iPart/iAssembly factory table
        ' (replicated changes also to dcIAssyTbRows)
        On Error Resume Next
        ck = ls.Count 'to trigger potential error

        If Err.Number = 0 Then 'should be good
            For Each it In ls
                ' REV[2022.03.23.1618]
                '     replacing Index of iPartTableRow
                '     with iPartTableRow itself, so it
                '     can just be pulled directly out
                '     of the Dictionary by the client
                '     process. if it needs the Index,
                '     it can just get it itself, right?
                rt.Add(it.MemberName, it) '.Index
                rt.Add(it.PartName, it) '.Index
                Debug.Print("")  'debug landing
            Next
        Else
            Debug.Print("ERROR " & CStr(Err.Number) & " (" & Hex$(Err.Number) & ")" & vbCrLf & Err.Description)
            Debug.Print(Join({"Could not access Table Rows", "for member of iPart factory.", vbCrLf}))
            'Stop
            Debug.Print("")  'Breakpoint Landing
        End If

        On Error GoTo 0

        dcIPartTbRows = rt
        'Debug.Print(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows.Count
        'Debug.Print(txDumpLs(dcIPartTbRows(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows).Keys)
    End Function

    Public Function dcIAssyTbCols(
    ls As Inventor.iAssemblyTableColumns
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim it As Inventor.iAssemblyTableColumn

        rt = New Scripting.Dictionary
        On Error Resume Next
        For Each it In ls
            rt.Add(it.DisplayHeading, it.Index)
            If Err.Number Then
                rt.Add(it.FormattedHeading, it.Index)
                Err.Clear()
            End If
        Next
        On Error GoTo 0
        dcIAssyTbCols = rt
    End Function

    Public Function dcIAssyTbRows(
    ls As Inventor.iAssemblyTableRows
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim it As Inventor.iAssemblyTableRow
        Dim ck As Long

        rt = New Scripting.Dictionary

        ' REV[2022.08.05.1003]
        ' added error trap code to more gracefully handle
        ' errors accessing iPart/iAssembly factory table
        ' (replicated from dcIPartTbRows)
        On Error Resume Next
        ck = ls.Count 'to trigger potential error

        If Err.Number = 0 Then 'should be good
            For Each it In ls
                rt.Add(it.MemberName, it.Index)
                rt.Add(it.DocumentName, it.Index)
                Debug.Print("")  'debug landing
            Next
        Else
            Debug.Print("ERROR " & CStr(Err.Number) & " (" & Hex$(Err.Number) & ")" & vbCrLf & Err.Description)
            Debug.Print(Join({"Could not access Table Rows", "for member of iAssembly factory.", vbCrLf}))
            Stop
        End If

        On Error GoTo 0

        dcIAssyTbRows = rt
        'Debug.Print(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows.Count
        'Debug.Print(txDumpLs(dcIAssyTbRows(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows).Keys)
    End Function

    Public Function m1g1f2(vr As Object) As Inventor.iPartTableColumn
        m1g1f2 = vr
    End Function

    Public Function m1g1f3(ad As Inventor.AssemblyDocument) As Long
        Dim oc As Inventor.ComponentOccurrence
        Dim rt As Long

        rt = 0
        For Each oc In ad.ComponentDefinition.Occurrences '(1)
            With oc.Definition
                'Debug.Print(Hex$(aiDocument(.Document).Type)
                'Stop
                If aiDocument(.Document).DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    rt = rt + iSyncPartFactory(aiDocPart(.Document))
                Else
                End If
            End With
        Next
        m1g1f3 = rt
    End Function

    Public Function m1g1f4() As Long
        m1g1f4 = m1g1f3(aiDocAssy(inventorApp.ActiveDocument))
    End Function

    Public Function m1g1f5(
    PartDesc As Inventor.PartDocument
    ) As Scripting.Dictionary
        ''  Retrieve Dictionary of Custom Properties
        ''
        Dim rt As Scripting.Dictionary
        Dim PropertySetMember As Inventor.PropertySet
        Dim PropertySetFactry As Inventor.PropertySet
        'Dim dcMember As Scripting.Dictionary
        Dim dcFactry As Scripting.Dictionary
        Dim InvProperty As Inventor.Property

        rt = New Scripting.Dictionary
        With PartDesc
            PropertySetMember = .Propertys.Item(gnCustom)
            With .ComponentDefinition
                If Not .iPartMember Is Nothing Then
                    With .iPartMember
                        With aiDocument(.ParentFactory.Parent)
                            PropertySetFactry = .Propertys.Item(gnCustom)
                            dcFactry = dcAiPropsIn(PropertySetFactry)
                        End With

                        For Each InvProperty In PropertySetMember
                            If Not dcFactry.Exists(InvProperty.Name) Then
                                'rt.Add(InvProperty.Name, InvProperty
                                rt = dcWithProp(
                                PropertySetFactry, InvProperty.Name, InvProperty.Text, rt
                            )
                            End If
                        Next
                    End With
                End If
            End With
        End With
        m1g1f5 = rt
    End Function

    Public Function m1g1f5t0() As String
        m1g1f5t0 = Join(m1g1f5(aiDocPart(
        aiDocAssy(inventorApp.ActiveDocument
        ).ComponentDefinition.Occurrences.Item(1
        ).Definition.Document
    )).Keys, vbCrLf)
    End Function

End Module