Module modMacros
    Dim inventorApp As Inventor.Application
    Public Function gnsPropAll() As Scripting.Dictionary
        gnsPropAll = nuDcPopulator(
        ).Setting("Part Number", "Item"
        ).Setting("Description", "Description1"
        ).Setting("Cost Center", "Family"
        ).Setting("GeniusMass", "Weight"
        ).Setting("Thickness", "Thickness"
        ).Setting("Extent_Area", "Diameter"
        ).Setting("Extent_Width", "Width"
        ).Setting("Extent_Length", "Length"
        ).Setting("RM", "Stock"
        ).Setting("RMQTY", "QuantityInConversionUnit"
        ).Setting("RMUNIT", "ConversionUnit"
        ).Setting("OFFTHK", "OFFTHK"
    ).Dictionary
        ' _
        '
    End Function

    Public Function gnsPropItem() As Scripting.Dictionary
        gnsPropItem = NuDcPopulator(
        ).Setting("Part Number", "Item"
        ).Setting("Cost Center", "Family"
        ).Setting("GeniusMass", "Weight"
        ).Setting("Thickness", "Thickness"
        ).Setting("Extent_Area", "Diameter"
        ).Setting("Extent_Width", "Width"
        ).Setting("Extent_Length", "Length"
        ).Setting("OFFTHK", "OFFTHK"
        ).Setting("Description", "Description1"
        ).Setting("RM", ""
        ).Setting("RMQTY", ""
        ).Setting("RMUNIT", ""
    ).Dictionary
    End Function

    Public Function gnsPropBomRaw() As Scripting.Dictionary
        '
        ' gnsPropBomRaw -- Property Names for Genius BOM
        ' !!!NOT READY!!! Just dup'd from gnsPropItem
        ' Needs adjustment to BOM Column/Field names
        '
        gnsPropBomRaw = NuDcPopulator(
        ).Setting("Part Number", "Product"
        ).Setting("RM", "Item"
        ).Setting("RMQTY", "QuantityInConversionUnit"
        ).Setting("RMUNIT", "ConversionUnit"
        ).Setting("Cost Center", "Family"
        ).Setting("GeniusMass", "Weight"
        ).Setting("Thickness", "Thickness"
        ).Setting("Extent_Area", "Diameter"
        ).Setting("Extent_Width", "Width"
        ).Setting("Extent_Length", "Length"
        ).Setting("OFFTHK", "OFFTHK"
        ).Setting("Description", "Description1"
    ).Dictionary
    End Function

    Public Function gnsPropsCurrent(
    Optional AiDoc As Inventor.Document = Nothing,
    Optional dcProps As Scripting.Dictionary = Nothing,
    Optional incTop As Long = 0,
    Optional inclPhantom As Long = 0
) As Scripting.Dictionary
        Dim rf As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        'Dim ActiveDoc As Document
        Dim ky As Object
        'Dim dx As Long

        If AiDoc Is Nothing Then
            gnsPropsCurrent = gnsPropsCurrent(
            aiDocActive(), dcProps, incTop, inclPhantom
        )
        ElseIf dcProps Is Nothing Then
            gnsPropsCurrent = gnsPropsCurrent(
            AiDoc, gnsPropAll(), incTop, inclPhantom
        ) 'gnsPropItem
        Else
            rf = dcProps(New Object() {gnsPropItem(),
                         "Part Number",
                         "Cost Center",
                         "GeniusMass",
                         "Extent_Area",
                         "Extent_Width", "Extent_Length", "RM", "RMQTY", "RMUNIT", "OFFTHK"})

            ' Collect Components for Processing
            ' (retained from Sub UpdateGeniusProperties_2023_0406_pre)
            ' NOTE[2021.08.09]:
            '     Function dcRemapByPtNum previously
            '     revised to address Key collisions
            '     in a crude manner. See that function
            '     for details.
            '
            rt = dcRemapByPtNum(
            dcAiDocComponents(
                AiDoc, , incTop
            )
        )
            'NOTE: incTop here is used to indicate
            '   whether to include top level assembly.
            '   This decision probably still needs
            '   to be made, but is not really
            '   addressed at the moment.

            ' Retrieve the full Component Collection
            With rt
                For Each ky In .Keys
                    ' Get Genius Properties for Item
                    ' Probably DON'T want to replace
                    ' Property Objects with their
                    ' values at this point, so that
                    ' they can be updated without
                    ' having to retrieve them again.
                    .Item(ky) = DcKeysInCommon(
                    dcOfPropsInAiDoc(.Item(ky)) _
                , dcProps, 1) 'dcProps replaces rf
                Next

                Debug.Print("") 'Breakpoint Landing
            End With

            gnsPropsCurrent = rt
            ' Don't index here.
            ' Let the client worry about that.
        End If
    End Function

    Public Sub testGnsPropsCurrent()
        Dim md As Inventor.Document
        Dim mPrpGnsAll As Scripting.Dictionary
        Dim mPrpGnsItm As Scripting.Dictionary
        Dim mPrpBomRaw As Scripting.Dictionary
        Dim dcPr As Scripting.Dictionary
        Dim dcVl As Scripting.Dictionary
        Dim dcGn As Scripting.Dictionary
        Dim dcGnDx As Scripting.Dictionary
        Dim dcBomDx As Scripting.Dictionary
        Dim nd As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim goAhead As VbMsgBoxResult
        Dim txOut As String
        Dim ky As Object

        md = aiDocActive()
        mPrpGnsAll = gnsPropAll() 'gnsPropItem
        ' mPrpBomRaw = gnsPropBomRaw()

        ' Retrieve all Documents'
        ' Genius Property Objects...
        dcPr = gnsPropsCurrent(md, mPrpGnsAll)

        ' ...and their current Values
        dcVl = New Scripting.Dictionary
        With dcPr : For Each ky In .Keys
                ' .Item(ky) = dcPropVals(dcOb(.Item(ky)))
                dcVl.Add(ky, dcPropVals(dcOb(.Item(ky))))
            Next : End With
        Debug.Print("") 'Breakpoint Landing
        If False Then
            ' send2clipBdWin10(ConvertToJson(dcVl, vbTab)) : Stop
        End If

        dcGnDx = dcRecSetDcDx4json(dcDxFromRecSetDc(
        DcFromAdoRS(CnGnsDoyle().Execute(q1g1x2(md)))
    ))
        With dcGnDx
            dcGn = dcOb(.Item("Item"))
            With dcOb(.Item(""))
                For Each ky In dcGn.Keys
                    dcGn.Item(ky) = .Item(dcGn.Item(ky)(0))
                Next
            End With
        End With
        Debug.Print("") 'Breakpoint Landing
        If False Then
            '  send2clipBdWin10(ConvertToJson(dcGn, vbTab)) : Stop
        End If

        ' This is to extract BOM from Assembly
        'bomInfoBkDn(bomViewStruct(aiDocAssy(

        dcBomDx = dcRecSetDcDx4json(dcDxFromRecSetDc(
        DcFromAdoRS(CnGnsDoyle().Execute(q1g2x2(md)))
    ))
        Debug.Print("") 'Breakpoint Landing
        If False Then
            '   send2clipBdWin10(ConvertToJson(dcBomDx, vbTab)) : Stop
        End If

        nd = New Scripting.Dictionary
        With dcPr : For Each ky In .Keys
                ck = DcKeysMissing(mPrpGnsAll, dcOb(.Item(ky)))
                If ck.Count > 0 Then nd.Add(ky, ck)
            Next : End With
        ''Debug.Print(ConvertToJson(nd, vbTab)
        Debug.Print("") 'Breakpoint Landing

        ' Index the Dictionary here
        ' (might be temporary)
        ' dcPr = dcRecDcDx4json( _
        dcDxFromRecSetDc(dcPr)

        ' Dump to JSON text format
        ' txOut = ConvertToJson(
        '    dcRecSetDcDx4json(
        '    dcDxFromRecSetDc(dcVl
        ')), vbTab)  'dcPr
        'Debug.Print(txOut

        goAhead = MsgBox(
        Join({
            "Assembly Name:",
            md.DisplayName,
            "Process Completed ",
            "Copy report text",
            "(JSON format)",
            "to Clipboard? ",
            "(Cancel for Debug)", vbCrLf, vbYesNoCancel, "Update Complete"}))
        If goAhead = vbCancel Then
            Stop
        ElseIf goAhead = vbYes Then
            send2clipBdWin10(txOut)
        End If
    End Sub

    Public Sub ExposeAllSheetMetalThicknesses()
        Dim PartDoc As Inventor.PartDocument
        Dim tk As Inventor.Parameter
        Dim ky As Object
        Dim ct As Long
        Dim xp As Long
        Dim nc As Long

        With dcAiSheetMetal(dcAiPartDocs(dcAiDocComponents(
        InventorApp.ActiveDocument
    )))
            ct = 0
            xp = 0
            nc = 0
            For Each ky In .Keys
                With aiCompDefShtMetal(aiDocPart(.Item(ky)).ComponentDefinition)
                    If .BOMStructure = BOMStructureEnum.kNormalBOMStructure Then
                        ct = 1 + ct

                        On Error Resume Next
                        Err.Clear()

                        ' REV[2023.01.18.1626]
                        ' added check for iPart member
                        ' to check exposure of Thickness
                        ' parameter of its Parent Factory
                        ' rather than its own.
                        '
                        ' this seeks to avoid errors
                        ' that now seem to arise when
                        ' attempting to  exposure
                        ' on the members themselves.
                        '
                        If .IsiPartMember Then
                            'Stop
                            With aiCompDefShtMetal(aiDocPart(
                        .iPartMember.ParentFactory.Parent))
                                tk = .Thickness
                            End With
                        Else
                            tk = .Thickness
                        End If

                        If Err.Number Then
                            Debug.Print("!ERROR!: " & ky)
                            tk = Nothing
                            nc = 1 + nc
                            Err.Clear()
                        Else
                            With tk
                                If .ExposedAsProperty Then
                                    Debug.Print("NOCHNGE: " & ky)
                                ElseIf aiCompDefShtMetal(
                                .Parent.Parent
                            ).IsiPartMember Then
                                    'Stop
                                    If aiDocPart(
                                    aiCompDefShtMetal(.Parent.Parent
                                    ).iPartMember.ParentFactory.Parent
                                ).ComponentDefinition.Parameters.Item(
                                    pnThickness
                                ).ExposedAsProperty Then
                                        Debug.Print("NOCHNGE: " & ky)
                                    Else
                                        nc = 1 + nc
                                        Debug.Print("FAILED!: " & ky)
                                    End If
                                Else
                                    .ExposedAsProperty = True
                                    If .ExposedAsProperty Then
                                        xp = 1 + xp
                                        Debug.Print("EXPOSED: " & ky)
                                    Else
                                        nc = 1 + nc
                                        Debug.Print("FAILED!: " & ky)
                                    End If
                                End If
                            End With
                        End If
                        On Error GoTo 0
                    End If
                End With
            Next
            If xp + nc > 0 Then
                MsgBox(Join({
                "Found " & CStr(ct) & " components.",
                "Thickness already exposed on " & CStr(ct - xp - nc),
                "   Exposed additional " & CStr(xp),
                "   Failed to expose " & CStr(nc)
            }, vbCrLf), vbOKOnly, "Sheet Metal Processed")
            Else
                MsgBox(Join({
                "Thickness already exposed",
                "on " & CStr(ct) & " components."
            }, vbCrLf), vbOKOnly, "No Change Required")
            End If
        End With
    End Sub

    Public Sub AddProps4Genius()
        Dim ky As Object
        Dim InvProperty As Inventor.Property

        With dcProps4genius(InventorApp.ActiveDocument)
            For Each ky In .Keys
                With aiProperty(obOf(.Item(ky)))
                    Debug.Print(.Parent.Name & ":" & .Name & "=" & .Text)
                End With
            Next
        End With
    End Sub
    'For Each itm In ActiveDocsComponents(InventorApp): Debug.Print(aiDocument(obOf(itm)).FullFileName: Next

    Public Sub MakeViewImageFiles()
        Debug.Print(d0g9f0())
    End Sub

    Sub iParts_GenerateAll()
        Dim oDoc As Inventor.PartDocument
        Dim oFactory As Inventor.iPartFactory
        Dim sFile As String
        Dim iCount As Long
        Dim i As Long
        Dim mx As Long
        Dim dx As Long
        Dim bk As Long

        oDoc = AskUser4aiDoc(, dcOf_iAll_Factories(
        InventorApp.Documents.VisibleDocuments(
    )))

        oDoc = AskUser4aiDoc(
    , dcOf_iPartFactories())

        If oDoc Is Nothing Then
            'do nothing
        ElseIf oDoc.ComponentDefinition.IsiPartFactory = True Then
            sFile = oDoc.FullFileName

            oFactory = oDoc.ComponentDefinition.iPartFactory

            'With oFactory
            mx = oFactory.TableRows.Count
            dx = 1
            Do 'For i = 1 To mx
                bk = 1 + mx - dx
                If bk > 10 Then bk = 10
                iCount = 0
                Do
                    InventorApp.StatusBarText _
                    = CStr(dx) & "/" & CStr(mx) & ": " _
                    & oFactory.TableRows.Item(dx).MemberName
                    'Member File creation
                    '.CreateMember dx
                    'disabled for testing
                    MsgBox(oFactory.TableRows.Item(dx).MemberName,
                    vbOKOnly, "Member " & CStr(dx) & "/" & CStr(mx))

                    dx = dx + 1
                    iCount = iCount + 1
                Loop While iCount < bk

                If dx <= mx Then 'iCount = 10
                    oDoc.Close
                    oDoc = InventorApp.Documents.Open(sFile)
                    oFactory = oDoc.ComponentDefinition.iPartFactory
                    iCount = 0
                End If
            Loop Until dx > mx 'Next
            'End With
        End If
    End Sub
End Module