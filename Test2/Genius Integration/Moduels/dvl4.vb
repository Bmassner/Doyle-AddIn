Module dvl4
    Dim inventorApp As Inventor.Application
    Public Function d4g2f1(
    Optional AiDoc As Inventor.Document = Nothing
) As Scripting.Dictionary
        '
        ' d4g2f1 -- not sure what
        '     but looks like something to do
        '     with grouping items by family
        '
        Dim rt As Scripting.Dictionary

        If AiDoc Is Nothing Then
            rt = d4g2f1(aiDocActive())
        Else
            With dcAiDocCompSetsByPtNum(AiDoc)
                'Stop
                If .Exists(1) Then
                    rt = DcFrom2Fields(CnGnsDoyle().Execute(
                    "select ls.it, ISNULL(i.Family, '') Fm from " _
                    & sqlValuesFromDc(dcOb(.Item(1)), "ls", "it") _
                    & " left join vgMfiItems i on ls.it = i.Item"
                ), "it", "fm")
                Else
                    rt = New Scripting.Dictionary
                    Stop
                End If
            End With
        End If

        d4g2f1 = rt
        'send2clipBdWin10 ConvertToJson(dcTransGrouped(d4g2f1(aiDocActive())), vbTab)
    End Function

    Public Function d4g0f1(
    AiDoc As Inventor.Document,
    Optional incTop As Long = 0
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        'Dim gn As Scripting.Dictionary
        Dim ky As Object
        Dim PartDoc As Inventor.Document '.PartDocument

        rt = New Scripting.Dictionary
        With dcOfDcOfDcByPlurality(dcAiDocsByPtNum(
        dcAiDocComponents(
        AiDoc, , incTop
    )))
            If .Exists(2) Then
                'ky = TypeName( _
                userChoiceFromDc(
                DcNewIfNone(
                dcOb(.Item(2))
            ))
                ky = nuSelFromDict(
                DcNewIfNone(
                dcOb(.Item(2))
            )).GetReply()
                'Stop
            End If

            If .Exists("") Then
                Stop
            End If

            Stop
            'Debug.Print(txDumpLs(dcNewIfNone(dcOb(obOf(.Item(1)))).Keys)
            gp = DcNewIfNone(dcOb(obOf(.Item(1))))
            With DcNewIfNone(dcOb(obOf(.Item(1)))) 'dcAiDocComponents(AiDoc, , incTop)
                For Each ky In .Keys
                    PartDoc = aiDocument(.Item(ky)) 'aiDocPart()
                    If PartDoc Is Nothing Then
                    Else
                        rt.Add(ky, dcGnsPtProps_Rev20220830_inProg(PartDoc))
                        'dcAiPropValsFromDc()
                        'dcOfGnsProps
                    End If
                Next
            End With
        End With

        ' gn = dcRecSetDcDx4json(dcFromAdoRS( _
        '    cnGnsDoyle().Execute(q1g1x2(AiDoc) _
        ')))
        'With gn
        'End With


        d4g0f1 = rt
        'send2clipBdWin10 ConvertToJson(d4g0f1(aiDocActive()), vbTab)
    End Function

    Public Function dcOfGnsProps(
    invDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcOfGnsProps
        '
        Dim rt As Scripting.Dictionary
        Dim dcPr As Scripting.Dictionary
        Dim ky As Object

        rt = DcNewIfNone(dc)

        dcPr = dcOfPropsInAiDoc(invDoc)
        With dcPr
            For Each ky In {
            pnPartNum, pnFamily, pnDesc,
            pnMaterial, pnStockNum, pnCatWebLink,
            pnMass, pnRawMaterial, pnRmQty, pnRmUnit,
            pnThickness, pnLength, pnWidth, pnArea}
                If .Exists(ky) Then
                    rt.Add(ky, .Item(ky))
                Else
                    rt.Add(ky, Nothing)
                End If
            Next

            ' NOTE[2022.09.16.1024]
            ' extraction of Categories XML text
            ' expected to move to content center
            ' processing in dcGnsValFromContentCtr
            'If .Exists("Categories") Then
            If Len(obAiProp(.Item("Categories")).Text) > 0 Then
                rt.Add("Categories", .Item("Categories"))
                'rt.Add("Parameters", dcAiDocParVals(invDoc)
                Debug.Print("") 'Breakpoint Landing: Content Center
            Else
                'Stop
            End If
            'Else
            '    Stop
            'End If
        End With

        dcOfGnsProps = rt
    End Function

    Public Function dcGnsValFromContentCtr(
    CpDef As Inventor.PartComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsValFromContentCtr
        '
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ky As Object

        Dim catXml As String
        Dim InvProperty As Inventor.Parameter

        rt = DcNewIfNone(dc)

        If CpDef Is Nothing Then
        Else
            With CpDef
                With aiDocPart(.Document).Propertys
                    catXml = .Item(
                gnDesign).Item(
                "Categories").Text

                    wk = dcAiPropsIn(
                    .Item(guidPrCLib)
                )

                End With

                ' NOTE[2022.09.16.1022]
                ' Categories XML processing
                ' with Parameter mapping
                ' will likely be addressed
                ' in a separate function
                ' to be called from here.
                'For Each Property In .Parameters
                'Next
            End With
        End If
        With wk
            For Each ky In {"Member FileName", "Family", "Standard", "Size Designation", "Categories"}
                If .Exists(ky) Then
                    rt.Add(ky, obAiProp(.Item(ky)).Text)
                End If
            Next
        End With

        dcGnsValFromContentCtr = rt
    End Function

    Public Function dcGnsValFromPartCompDef(
    CpDef As Inventor.PartComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsValFromPartCompDef
        '
        Dim rt As Scripting.Dictionary

        rt = DcNewIfNone(dc)

        If CpDef Is Nothing Then
        Else
            With CpDef
                'part of general ComponentDefinition
                'rt.Add("bomStruct", .BOMStructure

                If .IsContentMember Then
                    Debug.Print("") 'Breakpoint Landing: Content Member
                    'Stop 'and look at this one
                    'rt.Add("ContentMem", 1
                    rt.Add("ccPropVals", dcGnsValFromContentCtr(CpDef))
                    'want to develop this further
                    'possible to get library
                    'or collection?
                End If

                With .MassProperties
                    rt.Add(pnMass, System.Math.Round(cvMassKg2LbM * .Mass, 4))
                End With

                If .IsiPartMember Then
                    rt.Add("isIPartMem", 1)
                    rt.Add("iPartFactory", .iPartMember.ReferencedDocumentDescriptor.FullDocumentName)
                End If

                'part of general ComponentDefinition
                'With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                '    rt.Add("dimsModel", .Dictionary()
                'End With

                'part of general ComponentDefinition
                'MIGHT have some use
                'for this one in future
                'With .BOMQuantity
                '    '.BaseUnits
                'End With

                'also possible, but unsure
                With .Parameters
                End With
            End With

            With dcFlatPatVals(
            aiCompDefShtMetal(CpDef),
            DcDotted()
        )
                If .Count > 2 Then
                    rt.Add("flatPat",
                DcUnDotted(.Item(".")))
                Else
                    ''Debug.Print("KEYS = {" _
                    '& Join(.Keys, ", ") _
                    '& "}"
                    ''Stop 'and make sure
                    ''all that's in the Dictionary
                    ''are the self- and back-links
                    ''(check Immediate pane)
                End If
            End With
        End If

        dcGnsValFromPartCompDef = rt
    End Function

    Public Function dcGnsValFromAssyCompDef(
    CpDef As Inventor.AssemblyComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsValFromAssyCompDef
        '
        Dim rt As Scripting.Dictionary

        rt = DcNewIfNone(dc)

        If CpDef Is Nothing Then
        Else
            With CpDef
                'part of general ComponentDefinition
                'rt.Add("bomStruct", .BOMStructure

                With .MassProperties
                    rt.Add(pnMass, System.Math.Round(cvMassKg2LbM * .Mass, 4))
                End With

                If .IsiAssemblyMember Then
                    rt.Add("isIAssyMem", 1)
                    rt.Add("iAssyFactory", .iAssemblyMember.ReferencedDocumentDescriptor.FullDocumentName)
                End If

                'part of general ComponentDefinition
                'With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                '    rt.Add("dimsModel", .Dictionary()
                'End With

                'part of general ComponentDefinition
                'MIGHT have some use
                'for this one in future
                'With .BOMQuantity
                '    '.BaseUnits
                'End With

                'also possible, but unsure
                With .Parameters
                End With
            End With

            With dcFlatPatVals(
            aiCompDefShtMetal(CpDef),
            DcDotted()
            )
                If .Count > 2 Then
                    rt.Add("flatPat",
                DcUnDotted(.Item(".")))
                Else
                    'Debug.Print("KEYS = {" _
                    ' & Join(.Keys, ", ") _
                    ' & "}"
                    'Stop 'and make sure
                    'all that's in the Dictionary
                    'are the self- and back-links
                    '(check Immediate pane)
                End If
            End With
        End If

        dcGnsValFromAssyCompDef = rt
    End Function

    Public Function dcGnsValFromGenCompDef(
    CpDef As Inventor.ComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsValFromGenCompDef
        '
        Dim rt As Scripting.Dictionary

        rt = DcNewIfNone(dc)

        If CpDef Is Nothing Then
        Else
            With CpDef
                rt.Add("bomStruct", .BOMStructure)

                With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                    rt.Add("dimsModel", .Dictionary())
                End With

                'MIGHT have some use
                'for this one in future
                'With .BOMQuantity
                '    '.BaseUnits
                'End With
            End With
        End If

        dcGnsValFromGenCompDef =
        dcGnsValFromAssyCompDef(aiCompDefAssy(CpDef),
        dcGnsValFromPartCompDef(aiCompDefPart(CpDef), rt
    ))
    End Function

    Public Function dcGnsValGeneral(
    AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsValGeneral
        '
        Dim rt As Scripting.Dictionary
        'Dim dcPr As Scripting.Dictionary
        'Dim ky As Object

        rt = DcNewIfNone(dc)

        If AiDoc Is Nothing Then
        Else
            With AiDoc
                With .Propertys.Item(gnDesign)
                    rt.Add(pnPartNum, .Item(pnPartNum).Text)
                End With

                rt.Add("subType", .SubType)
            End With

            rt = dcGnsValFromGenCompDef(
            aiCompDefOf(AiDoc), rt
        )
        End If

        ' dcPr = dcOfPropsInAiDoc(AiDoc)
        'With dcPr
        '    For Each ky In { _
        '        pnPartNum, pnFamily, pnDesc, _
        '        pnMaterial, pnStockNum, pnCatWebLink, _
        '        pnMass, pnRawMaterial, pnRmQty, pnRmUnit, _
        '        pnThickness, pnLength, pnWidth, pnArea _
        '    )
        '        If .Exists(ky) Then
        '            rt.Add(ky, .Item(ky)
        '        Else
        '            rt.Add(ky, Nothing
        '        End If
        '    Next
        'End With

        dcGnsValGeneral = rt
    End Function

    Public Function dcGnsPtProps_Rev20220830_inProg(
    AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary '.PartDocument
        '
        ' dcGnsPtProps_Rev20220830_inProg
        '
        Dim rt As Scripting.Dictionary
        'Dim dcPr As Scripting.Dictionary
        'Dim dcVl As Scripting.Dictionary
        'Dim ky As Object

        ' rt = dcNewIfNone(dc)
        rt = dcGnsValGeneral(
        AiDoc, DcNewIfNone(dc)
    )

        With dcOfGnsProps(
    AiDoc, DcDotted())
            If .Count > 2 Then
                rt.Add("props",
            DcUnDotted(.Item(".")))

                rt.Add("propVals",
            dcPropVals(rt.Item("props")))

                If .Exists("Categories") Then
                    rt.Add("Parameters",
                dcAiDocParVals(AiDoc))
                End If
            End If
        End With

        If AiDoc Is Nothing Then
        Else
            ' dcPr = dcOfGnsProps(AiDoc) 'dcOfPropsInAiDoc
            ' dcVl = dcAiPropValsFromDc(dcPr)

            With AiDoc
                With .Propertys.Item(gnDesign)
                    'rt.Add(pnPartNum, .Item(pnPartNum).Text
                    'only want this one to tie
                    'back to other Dictionaries

                    'rt.Add(pnFamily, .Item(pnFamily).Text
                    'rt.Add(pnDesc, .Item(pnDesc).Text
                    'rt.Add(pnMaterial, .Item(pnMaterial).Text
                    'rt.Add(pnStockNum, .Item(pnStockNum).Text
                    'rt.Add(pnCatWebLink, .Item(pnCatWebLink).Text

                    'rt.Add(pnThickness, .Item(pnThickness).Text
                    'this one's a Custom property
                    'whether Inventor handles
                    'itself or not is uncertain.

                    'pnMass, pnRawMaterial, pnRmQty, pnRmUnit,
                    'pnLength, pnWidth, pnArea
                End With

                '.ActiveMaterial
                '.FullDocumentName
                '.UnitsOfMeasure
                '.NeedsMigrating
                '.RequiresUpdate
            End With
        End If 'AiDoc Is Nothing

        'Stop
        'Call iSyncPartFactory(AiDoc)
        ' rt = dcVl

        dcGnsPtProps_Rev20220830_inProg = rt
    End Function

    Public Function sqlValuesFromDc(
    dc As Scripting.Dictionary,
    Optional vw As String = "ls",
    Optional it As String = "it"
) As String
        '
        ' sqlValuesFromDc - generate SQL
        '     "VALUES" clause from Keys
        '     of supplied Dictionary.
        '     result is a relation
        '     of one attribute
        '     '
        '     VALUES clause must end
        '     with an AS phrase naming
        '     the relation and all
        '     attributes.
        '     '
        '     in this function, the
        '     names default to "ls"
        '     (list) for the relation,
        '     and "it" (item) for
        '     its one attribute
        '
        If dc Is Nothing Then
            sqlValuesFromDc = sqlValuesFromDc(
        New Scripting.Dictionary)
        Else
            sqlValuesFromDc = "(values ('" _
        & Join(dc.Keys, "'), ('") &
        "')) as ls(it)"
        End If
    End Function

    Public Function dcAiDocCompSetsByPtNum(
    AiDoc As Inventor.Document,
    Optional incTop As Long = 0
) As Scripting.Dictionary
        '
        ' dcAiDocCompsByPtNum -- formerly d4g3f1
        '
        '
        '
        Dim dc As Scripting.Dictionary
        'Dim ct As Long

        'ct = 1 'to include main assembly (for now)
        'now disabled in favor of input parameter incTop

        'dcAiDocsByPtNum replaces dcRemapByPtNum
        dc = dcOfDcOfDcByPlurality(
        dcAiDocsByPtNum(
        dcAiDocComponents(
            AiDoc, , incTop
        )
    )) 'incTop replaces ct
        'dcOfDcOfDcByItemCount() removed from original lineup
        'of dcOfDcOfDcByPlurality(dcOfDcOfDcByItemCount(dcAiDocsByPtNum(
        'dcOfDcOfDcByPlurality calls dcOfDcOfDcByItemCount internally
        'as part of its normal processing.

        'dcAiDocCompsByPtNum = dc
        'WAS going to simply replace assignment of
        'dc above with direct assignment of dcAiDocCompsByPtNum,
        'but further processing here may be desired.
    End Function

    Public Function dcAiDocsByPtNum(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcAiDocsByPtNum -- formerly d4g3f2
        '
        ' Returns Dictionary of Dictionaries
        ' of Inventor Documents keyed on
        ' associated Part Numbers.
        '
        ' Derived from dcRemapByPtNum, this
        ' variation collects all models of a
        ' given Part Number into a secondary
        ' Dictionary, under each Document's
        ' file name. Ideally, Part Numbers
        ' should map one-to-one to Documents,
        ' so each sub Dictionary should
        ' contain only one entry.
        '
        ' However, as it IS possible for more
        ' than one model to represent the same
        ' Part, more than one Document might
        ' in fact have the same Part Number.
        '
        ' Therefore, it may sometimes prove
        ' necessary to take additional steps
        ' in properly identifying which model
        ' (or models) to process in preparation
        ' for Genius.
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))
                pn = CStr(aiDocPropVal(
                PartDoc, pnPartNum, gnDesign
            ))

                ' REV[2022.05.17.1536]
                ' removing check/handling
                ' for blank/empty part number.
                ' client process can deal with that.
                With rt
                    If .Exists(pn) Then 'do nothing
                    Else
                        .Add(pn, New Scripting.Dictionary)
                    End If

                    gp = .Item(pn)
                End With

                With gp
                    If .Exists(ky) Then
                        Stop 'this should NOT happen
                    Else
                        .Add(ky, PartDoc)
                    End If
                End With
            Next
        End With

        dcAiDocsByPtNum = rt
    End Function

    Public Function dcOfDcOfDcByItemCount(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcOfDcOfDcByItemCount -- formerly d4g3f3
        '
        ' subdivide supplied Dictionary
        ' of Dictionaries into groups
        ' by Count of members.
        '
        ' result is a 3rd-order Dictionary,
        ' that is, a Dictionary (1)
        '     keyed by member count
        ' of Dictionaries (2)
        '     keyed by a shared key
        ' of yet more Dictionaries (3)
        '     keyed to some unique value
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim xp As Scripting.Dictionary
        Dim ky As Object
        Dim ct As Long

        rt = New Scripting.Dictionary

        With dc : For Each ky In .Keys
                gp = .Item(ky)
                ct = gp.Count

                With rt
                    If Not .Exists(ct) Then
                        .Add(ct, New Scripting.Dictionary)
                    End If

                    xp = .Item(ct)
                End With

                With xp
                    If .Exists(ky) Then
                        Stop 'this should NOT happen
                    Else
                        .Add(ky, gp)
                    End If : End With
            Next : End With

        dcOfDcOfDcByItemCount = rt
    End Function

    Public Function dcOfDcOfDcByPlurality(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcOfDcOfDcByPlurality -- formerly d4g3f4
        '
        ' given a 2nd-order Dictionary (NOT 3rd)
        ' as supplied by dcAiDocsByPtNum (NOT dcOfDcOfDcByItemCount),
        ' return a reorganized version as follows:
        '
        ' under key 1: a Dictionary of all
        '     Dictionaries having only one member.
        '     this should form the bulk of the
        '     supplied Dictionary's content.
        '
        ' under key 2: a Dictionary of Dictionaries
        '     having more than one member. these
        '     "plurals" might require additional
        '     review and/or processing to resolve
        '     ambiguities, conflicts, etc.
        '
        ' under key "" (blank string): the Dictionary,
        '     if present, of members with no assigned
        '     part or item number. this should almost
        '     NEVER arise, but again, might require
        '     special processing to resolve issues.
        '
        ' '
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim xp As Scripting.Dictionary
        Dim ky As Object
        Dim ct As Long

        rt = New Scripting.Dictionary

        ' to avoid modifying supplied Dictionary,
        ' generate a copy to work with directly.
        gp = DcCopy(dc)
        With gp
            If .Exists("") Then 'get the 
                'of blank "numbered"
                'items moved over...

                rt.Add("", .Item(""))
                .Remove("")
            End If : End With
        'before grouping by member counts:

        gp = dcOfDcOfDcByItemCount(gp)
        With gp
            'prep the "singles" return Dictionary
            xp = New Scripting.Dictionary
            rt.Add(1, xp)

            If .Exists(1) Then 'proceed to move
                'the (one, single) member of each
                'Dictionary under this one
                With dcOb(.Item(1))
                    For Each ky In .Keys
                        With dcOb(.Item(ky))
                            xp.Add(ky, .Items(0)) '.Item(.Keys(0))
                        End With
                    Next : End With
                .Remove(1)
            Else 'we've got nothing to move.
                'this is actually a problem,
                'but that's for the client
                'process to handle.
            End If
            'at this point, any remaining members
            'should be "plural" Dictionaries
            'containing more than one member.

            'THESE are to be combined into one
            '"plural" Dictionary to be returned.
            xp = New Scripting.Dictionary
            'DO NOT add to return Dictionary yet!

            For Each ky In .Keys
                'this step generates a NEW
                'Dictionary at each iteration.
                xp = DcKeysCombined(
                xp, dcOb(.Item(ky)), 1
            )
                'it does NOT update the original!
            Next

            'NOW, we can add the final result
            'to the return Dictionary...
            If xp.Count > 0 Then
                '...ASSUMING any are
                'left to add, of course!
                rt.Add(2, xp)
            End If
        End With

        'Stop 'because NOT sure this thing
        'is ready for prime time...

        'disabling the following section completely
        'it should not be needed, as all three parts
        'of the return Dictionary should be in place
        'at the end of the preceding With block.

        'this SHOULD add Dictionaries for all
        'Part Numbers with more than one Document,
        'in a single Dictionary of "plurals",
        'but need to check it out for sure yet.
        'Hence the preceding Stop

        'rt.Add(2, dcKeysMissing(gp, rt.Item(1))

        'With dc: For Each ky In .Keys
        '     gp = .Item(ky)
        '    ct = gp.Count
        '
        '    With rt
        '        If Not .Exists(ct) Then
        '        .Add(ct, New Scripting.Dictionary
        '        End If
        '
        '         xp = .Item(ct)
        '    End With
        '
        '    With xp
        '    If .Exists(ky) Then
        '        Stop 'this should NOT happen
        '    Else
        '        .Add(ky, gp
        '    End If: End With
        'Next: End With

        dcOfDcOfDcByPlurality = rt
    End Function

    Public Function d4g3f5from2(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' d4g3f5from2
        '
        ' Returns Dictionary of Dictionaries
        ' of Inventor Documents keyed on
        ' associated Part Numbers.
        '
        ' Derived from dcRemapByPtNum, this
        ' variation collects all models of a
        ' given Part Number into a secondary
        ' Dictionary, under each Document's
        ' file name. Ideally, Part Numbers
        ' should map one-to-one to Documents,
        ' so each sub Dictionary should
        ' contain only one entry.
        '
        ' However, as it IS possible for more
        ' than one model to represent the same
        ' Part, more than one Document might
        ' in fact have the same Part Number.
        '
        ' Therefore, it may sometimes prove
        ' necessary to take additional steps
        ' in properly identifying which model
        ' (or models) to process in preparation
        ' for Genius.
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))

                'PartDoc.co

                pn = CStr(aiDocPropVal(
                PartDoc, pnPartNum, gnDesign
            ))

                ' REV[2022.05.17.1536]
                ' removing check/handling
                ' for blank/empty part number.
                ' client process can deal with that.
                With rt
                    If .Exists(pn) Then 'do nothing
                    Else
                        .Add(pn, New Scripting.Dictionary)
                    End If

                    gp = .Item(pn)
                End With

                With gp
                    If .Exists(ky) Then
                        Stop 'this should NOT happen
                    Else
                        .Add(ky, PartDoc)
                    End If
                End With
            Next
        End With

        d4g3f5from2 = rt
    End Function

    Public Function d4g3f5b(
    cd As Inventor.ComponentDefinition
) As Object

    End Function

    Public Function d4g3f5a(
    cd As Inventor.ComponentDefinition
) As Object 'Inventor.iAssemblyTableCell
        'aiCompDefOf

        If cd Is Nothing Then
            d4g3f5a = Nothing
        ElseIf TypeOf cd Is Inventor.AssemblyComponentDefinition Then
            With aiCompDefAssy(cd)
                If .IsiAssemblyMember Then
                ElseIf .IsiAssemblyFactory Then
                ElseIf .IsModelStateMember Then
                ElseIf .IsModelStateFactory Then
                End If
            End With
        ElseIf TypeOf cd Is Inventor.PartComponentDefinition Then
        End If
    End Function

    Public Function gnsUpdtAll_iFact(
    cd As Inventor.ComponentDefinition
) As Scripting.Dictionary
        If TypeOf cd Is Inventor.PartComponentDefinition Then
            gnsUpdtAll_iFact = gnsUpdtAll_iPart(cd)
        ElseIf TypeOf cd Is Inventor.AssemblyComponentDefinition Then
            gnsUpdtAll_iFact = gnsUpdtAll_iAssy(cd)
        Else
            gnsUpdtAll_iFact = New Scripting.Dictionary
        End If
    End Function

    Public Function gnsUpdtAll_iAssy(
    cd As Inventor.AssemblyComponentDefinition
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim md As Inventor.AssemblyDocument
        Dim fc As Inventor.iAssemblyFactory
        Dim BillRow As Inventor.iAssemblyTableRow
        Dim r0 As Inventor.iAssemblyTableRow

        rt = New Scripting.Dictionary

        If cd Is Nothing Then
        ElseIf cd.IsiAssemblyFactory Then
            With cd.iAssemblyFactory
                md = .Parent.Parent

                With .TableColumns '.Item()
                    '"GeniusMass [Custom]"
                End With

                'note initial DefaultRow
                r0 = .DefaultRow

                For Each BillRow In .TableRows
                    .DefaultRow = BillRow
                    rt.Add(.DefaultRow.MemberName,
                dcOfDcAiPropVals(dcGeniusProps(md)))
                    'md.Save
                Next

                'restore initial DefaultRow
                .DefaultRow = r0
            End With
        Else
        End If

        gnsUpdtAll_iAssy = rt
        'Debug.Print(dcOfDcAiPropVals(dcGeniusProps(aiDocActive())).Count
        'send2clipBdWin10 ConvertToJson(gnsUpdtAll_iAssy(aiDocAssy(aiDocActive()).ComponentDefinition), vbTab)
    End Function

    Public Function gnsUpdtAll_iPart(
    cd As Inventor.PartComponentDefinition
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim md As Inventor.PartDocument
        Dim fc As Inventor.iPartFactory
        Dim BillRow As Inventor.iPartTableRow
        Dim r0 As Inventor.iPartTableRow

        rt = New Scripting.Dictionary

        If cd Is Nothing Then
        ElseIf cd.IsiPartFactory Then
            With cd.iPartFactory
                md = .Parent '.Parent

                With .TableColumns '.Item()
                    '"GeniusMass [Custom]"
                End With

                'note initial DefaultRow
                r0 = .DefaultRow

                For Each BillRow In .TableRows
                    .DefaultRow = BillRow
                    rt.Add(.DefaultRow.MemberName,
                dcAiPropValsFromDc(dcGeniusProps(md))) 'dcOfDcAiPropVals
                    System.Windows.Forms.Application.DoEvents()
                    md.Save()
                Next

                'restore initial DefaultRow
                .DefaultRow = r0
            End With
        Else
        End If

        gnsUpdtAll_iPart = rt
        'send2clipBdWin10 ConvertToJson(gnsUpdtAll_iPart(aiDocPart(aiDocActive()).ComponentDefinition), vbTab)
    End Function

    Public Function d4g3f6pt(
    cd As Inventor.PartComponentDefinition
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim mx As Long
        Dim dx As Long
        Dim ck As Inventor.iPartTableColumn

        rt = New Scripting.Dictionary
        If cd Is Nothing Then
        ElseIf cd.IsiPartFactory Then
            Stop
            With cd.iPartFactory.TableColumns
                mx = .Count
                Stop
                For dx = 1 To mx
                    ck = .Item(dx)
                    With ck
                        rt.Add(.Heading, NuDcPopulator(
                    ).Setting("dx", dx
                    ).Setting("dh", .DisplayHeading
                    ).Setting("fh", .FormattedHeading
                    ).Setting("dt", .ReferencedDataType
                    ).Setting("ob", .ReferencedObject
                    ).Setting("ot", TypeName(.ReferencedObject)
                    ).Dictionary)
                        '
                        ').Setting("hd", .Heading _
                        '
                        Stop
                    End With
                    'ck.ReferencedObject
                    'Debug.Print(obAiProp(ck.ReferencedObject).Name
                    'rt.Add(dx, nuDcPopulator().Setting("hd", ck.Heading).Setting("dh", ck.DisplayHeading).Setting("fh", ck.FormattedHeading).Dictionary
                    'If ck.ReferencedDataType = kMemberNameColumn Then
                    'End If
                Next
            End With
        ElseIf cd.IsiPartMember Then
            Stop
            rt = d4g3f6pt(aiDocPart(cd.iPartMember.ParentFactory.Parent).ComponentDefinition)
            'cd.iPartMember.ParentFactory.Parent
        Else
        End If

        d4g3f6pt = rt
    End Function

    Public Function d4g3f6as(
    cd As Inventor.AssemblyComponentDefinition
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim mx As Long
        Dim dx As Long
        Dim ck As Inventor.iAssemblyTableColumn

        rt = New Scripting.Dictionary
        If cd Is Nothing Then
        ElseIf cd.IsiAssemblyFactory Then
            Stop
            With cd.iAssemblyFactory.TableColumns
                mx = .Count
                Stop
                For dx = 1 To mx
                    ck = .Item(dx)
                    Stop
                Next
            End With
        ElseIf cd.IsiAssemblyMember Then
            Stop
            'cd.iAssemblyMember.ParentFactory
        Else
        End If

        d4g3f6as = rt
    End Function

    Public Function d4g3f7pt(
    cd As Inventor.PartComponentDefinition
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim hd As Scripting.Dictionary
        'Dim md As Inventor.PartDocument
        Dim fc As Inventor.iPartFactory
        Dim co As Inventor.iPartTableColumn
        Dim BillRow As Inventor.iPartTableRow
        'Dim r0 As Inventor.iPartTableRow
        'Dim df As Long

        rt = New Scripting.Dictionary
        rt.Add("", New Scripting.Dictionary)
        hd = rt.Item("")


        If cd Is Nothing Then
            fc = Nothing
        ElseIf cd.IsiPartFactory Then
            fc = cd.iPartFactory
        ElseIf cd.IsiPartMember Then
            fc = cd.iPartMember.ParentFactory
        Else
            fc = Nothing
        End If

        If Not fc Is Nothing Then
            With fc
                ' md = .Parent '.Parent

                hd.Add("", {
            "Index",
            "Key",
            "CustomColumn",
            "DisplayHeading",
            "FormattedHeading",
            "ReferencedDataType"
        })
                For Each co In .TableColumns : With co
                        '.Item()
                        With co
                            hd.Add(.Heading, {
                .Index,
                .Key,
                .CustomColumn,
                .DisplayHeading,
                .FormattedHeading,
                .ReferencedDataType})
                        End With
                        'Stop
                    End With : Next

                'note initial DefaultRow
                ' r0 = .DefaultRow

                For Each BillRow In .TableRows
                    With BillRow
                        'df = BillRow Is r0
                        rt.Add(.Index, { .Index, .MemberName, .PartName, BillRow}) ', df
                    End With
                    '.DefaultRow = BillRow
                    'rt.Add(.DefaultRow.MemberName, _
                    dcAiPropValsFromDc(dcGeniusProps(.md)) 'dcOfDcAiPropVals
                    'DoEvents
                    'md.Save
                Next

                'restore initial DefaultRow
                '.DefaultRow = r0
            End With : End If

        d4g3f7pt = rt
        'send2clipBdWin10 ConvertToJson(d4g3f7pt(aiDocPart(aiDocActive()).ComponentDefinition), vbTab)
    End Function

    Public Function d4g4f0(
    Optional AiDoc As Inventor.Document = Nothing
) As Scripting.Dictionary
        '
        ' d4g4f0 -- rebuilding Sub Update_Genius_Properties
        '           more or less from the ground up
        '
        'Dim invProgressBar As Inventor.ProgressBar

        'Dim fc As gnsIfcAiDoc
        Dim dc As Scripting.Dictionary
        Dim mt As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        'Dim goAhead As VbMsgBoxResult
        'Dim ActiveDoc As Document
        'Dim txOut As String
        Dim ky As Object
        Dim k2 As Object
        'Dim kyPt As Object
        Dim ct As Long

        'Dim dx As Long

        'Dim Fm As FmIfcTest05A

        ' NOTE[2022.06.01.1441]
        ' adding check for supplied Document
        ' call for selection if none
        If AiDoc Is Nothing Then
            ' rt = d4g4f0(userChoiceFromDc(dcAiDocsVisible(), aiDocActive()))
            rt = d4g4f0(aiDocActive())
        Else
            ' NOTE[2022.06.01.1442]
            ' disabling/skipping user checks for now
            ' this isn't the purpose of this whole mess.
            ' Confirm User Request
            ' to process active Document
            'goAhead = MsgBox( _
            '    Join({ _
            '        "Are you sure you want to process this document?", _
            '        "The process may require a few minutes depending on assembly size.", _
            '        "Suppressed and excluded parts will not be processed." _
            '    ), " "), _
            '    vbYesNo + vbQuestion, _
            '    "Process Document Custom iProperties" _
            ')
            'If goAhead = vbYes Then
            '
            'Else
            'End If

            '' NOTE[2022.06.01.1444]
            '' see Update_Genius_Properties REV[2022.05.24.0956]
            'With dcAiDocCompsByPtNum(AiDoc, ct) 'ActiveDoc
            '    If .Exists("") Then
            '        Stop 'for now
            '        'don't expect this situation
            '        'to occur frequently, so won't
            '        'worry about a handler just yet
            '    End If

            'If .Exists(2) Then
            '        'THIS situation IS known to occur,
            '        'if not TERRIBLY frequently, so a
            '        'handler here is a good idea.
            '        '
            '        With NuDcPopulator(.Item(2)) 'd4g4f4(dcOb(.Item(2)))
            '            Debug.Print(MsgBox(
            '            msg_2022_0603_1127(.Dictionary),
            '            vbOKOnly Or vbInformation,
            '            "Duplicate Part Numbers!"
            '        )) 'with just a slight modification,
            '        End With

            '        'With d4g4f4(dcOb(.Item(2)))
            '        '    Stop
            '        '
            '        '    'fortunately, we have one ready made
            '        '    'in the dcRemapByPtNum function this
            '        '    'section is replacing (see above).
            '        '    Debug.Print(MsgBox( _
            '        '        Join({ _
            '        '            "The following Part Numbers are", _
            '        '            "assigned to more than one Model:", _
            '        '            "", _
            '        '            vbTab & Join(.Keys, vbCrLf & vbTab), _
            '        '            "" _
            '        '        }, vbCrLf), _
            '        '        vbOKOnly Or vbInformation, _
            '        '        "Duplicate Part Numbers!" _
            '        '    ) 'with just a slight modification,
            '        '    'this serves to notify the user
            '        '    'just as dcRemapByPtNum did before.
            '        '    'a more sophisticated response may
            '        '    'eventually be called for, but for
            '        '    'now, this will do.
            '        'End With
            '    End If

            '    'and HERE is the step which ACTUALLY
            '    'replaces the prior version above.
            '    'Key 1 is guaranteed to be present
            '    'in the Dictionary returned, so no
            '    'need to check for it here.
            '    dc = dcOb(.Item(1))
            'End With
            ''
            ''
            '

            '
            ' NOTE[2022.06.01.1502]
            ' this section expected to be
            ' exported to its own function
            ' NOTE[2022.06.02.0906]
            ' (follow-up) original code
            ' extracted to functions
            '    ' dcOfKeys2match and d4g4f1
            '    mt = DcOfKeys2match({
            '    pnFamily, pnMass,
            '    pnRawMaterial,
            '    pnRmQty, pnRmUnit,
            '    pnWidth, pnLength,
            '    pnArea, pnThickness
            '})
            'pnFamily       replaces "Cost Center"
            'pnMass         replaces "GeniusMass"
            'pnRawMaterial  replaces "RM"
            'pnRmQty        replaces "RMQTY"
            'pnRmUnit       replaces "RMUNIT"
            'pnWidth        replaces "Extent_Width"
            'pnLength       replaces "Extent_Length"
            'pnArea         replaces "Extent_Area"
            'pnThickness    replaces "Thickness"

            ' rt = d4g4f1(dc, mt)

            'rt = New Scripting.Dictionary
            'With d4g4f1(dc, mt)
            '    For Each ky In .Keys
            '        With rt
            '            If Not .Exists(ky) Then
            '                .Add(ky, New Scripting.Dictionary)
            '            End If

            '            wk = .Item(ky)
            '        End With

            '        With dcOb(.Item(ky))
            '            For Each k2 In .Keys
            '                wk.Add(k2, obAiProp(.Item(k2)).Text)
            '            Next
            '        End With
            '    Next : End With

            '
            '
            '
            ' rt = dcCopy(dc)
        End If
        '
        d4g4f0 = rt
        'send2clipBdWin10 ConvertToJson(d4g4f0(), vbTab)
    End Function

    Public Function d4g4f1(
    dc As Scripting.Dictionary,
    rf As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' d4g4f1 -- returns a Dictionary of Dictionaries
        '     copied from supplied Dictionary dc,
        '     but with only those Keys matching those
        '     found in supplied 'reference' Dictionary rf
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                rt.Add(ky, DcKeysInCommon(
            dcOfPropsInAiDoc(
            aiDocument(.Item(ky))
            ), rf, 1
        ))
            Next : End With

        d4g4f1 = rt
    End Function

    Public Function d4g4f2(
    dc As Scripting.Dictionary
) As Inventor.PartDocument
        '
        ' d4g4f2 -- given a Dictionary of Part Documents
        '     return the first Content Center Member found
        '     (if none found, return Nothing)
        '
        Dim ky As Object
        Dim PartDoc As Inventor.PartDocument

        With dc : For Each ky In .Keys
                If PartDoc Is Nothing Then
                    PartDoc = aiDocPart(aiDocument(.Item(ky)))
                    If Not PartDoc Is Nothing Then
                        If Not PartDoc.ComponentDefinition(
            ).IsContentMember Then PartDoc = Nothing
                    End If
                End If
            Next : End With

        d4g4f2 = PartDoc
    End Function

    Public Function d4g4f3(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' d4g4f3 -- given a Dictionary of Part Document
        '     Dictionaries, return a sub containing
        '     only Content Center Members
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim PartDoc As Inventor.PartDocument

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                PartDoc = d4g4f2(.Item(ky))
                If Not PartDoc Is Nothing Then
                    rt.Add(ky, PartDoc)
                End If
            Next : End With

        d4g4f3 = rt
    End Function

    Public Function d4g4f4(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' d4g4f4 -- given a Dictionary of Part Document
        '     Dictionaries, return a sub dropping
        '     any with Content Center Members
        '
        d4g4f4 = DcKeysMissing(dc, d4g4f3(dc))
    End Function

    Public Function d4g4f5() As Scripting.Dictionary  '_
        ' Removed invalid line: dc As Scripting.Dictionary
        '
        '
        ' d4g4f5 -- given a Dictionary of Part Document
        '     Dictionaries, return a sub dropping
        '     any with Content Center Members
        '
        Dim rt As Scripting.Dictionary
        Dim dcRb As Scripting.Dictionary
        Dim dcTb As Scripting.Dictionary
        Dim dcRp As Scripting.Dictionary
        'Dim dcRb As Scripting.Dictionary
        Dim rb As Inventor.Ribbon
        Dim tb As Inventor.RibbonTab
        Dim rp As Inventor.RibbonPanel
        Dim ob As Object

        rt = New Scripting.Dictionary
        'InventorApp.UserInterfaceManager.RibbonState
        With InventorApp.UserInterfaceManager
            For Each rb In .Ribbons
                dcRb = New Scripting.Dictionary
                With rb
                    rt.Add(.InternalName, dcRb)
                    For Each tb In .RibbonTabs '.QuickAccessControls
                        dcTb = New Scripting.Dictionary
                        With tb
                            dcRb.Add(.InternalName, dcTb)
                            For Each rp In .RibbonPanels
                                dcRp = New Scripting.Dictionary
                                With rp
                                    dcTb.Add(.InternalName, dcRp)
                                    Stop
                                End With
                            Next
                        End With
                    Next
                End With
            Next : End With
        d4g4f5 = rt
    End Function

    Public Function compDefOfPart(
    AiDoc As Inventor.PartDocument
) As Inventor.ComponentDefinition
        If AiDoc Is Nothing Then
            compDefOfPart = Nothing
        Else
            compDefOfPart = AiDoc.ComponentDefinition
        End If
    End Function

    Public Function compDefOfAssy(
    AiDoc As Inventor.AssemblyDocument
) As Inventor.ComponentDefinition
        If AiDoc Is Nothing Then
            compDefOfAssy = Nothing
        Else
            compDefOfAssy = AiDoc.ComponentDefinition
        End If
    End Function

    Public Function compDefOf(
    AiDoc As Inventor.Document
) As Inventor.ComponentDefinition
        Dim rt As Inventor.ComponentDefinition

        rt = compDefOfPart(aiDocPart(AiDoc))
        If rt Is Nothing Then
            rt = compDefOfAssy(aiDocAssy(AiDoc))
        End If

        'AiDoc.FullFileName

        compDefOf = rt
    End Function

    Public Function famOfAiDoc(
    AiDoc As Inventor.Document
) As String
        '
        Dim itNum As String
        Dim mdFam As String
        Dim gnFam As String

        Dim pf As String
        Dim StockFam As String

        Dim ck As VbMsgBoxResult

        If AiDoc Is Nothing Then
            famOfAiDoc = ""
        Else
            With AiDoc
                ' NOTE!!! ONLY use this for Assemblies!
                ' will disable until better  up
                'With nuDcPopulator( _
                '    ).Setting("doyle", "D" _
                '    ).Setting("riverview", "R" _
                ').Matching(Split(.FullDocumentName, "\"))
                '    If .Count = 1 Then
                '        pf = .Item(.Keys(0))
                '    Else
                '        pf = ""
                '    End If
                'End With

                With .Propertys.Item(gnDesign)
                    mdFam = .Item(pnFamily).Text
                    itNum = .Item(pnPartNum).Text
                    gnFam = famInGenius(itNum)
                End With
            End With

            With compDefOf(AiDoc)
                If .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    StockFam = "PTS"
                End If
            End With
        End If
    End Function

    Public Function famInGenius(itNum As String) As String
        Dim gnFam As String

        With CnGnsDoyle().Execute(
        "select Family from vgMfiItems where Item = '" _
        & itNum & "';"
    )
            If .BOF Or .EOF Then
                gnFam = ""
            Else
                gnFam = Split(.GetString(
            .adClipString, , "", "", ""
            ), vbCr)(0)
            End If
        End With

        famInGenius = gnFam
    End Function

    Public Function famIfValid(mdFam As String) As String
        With CnGnsDoyle().Execute(Join({
        "select ISNULL(f.Family, '') Family",
        "from (values ('" & mdFam & "')) as i(f)",
        "left join vgMfiFamilies f on i.f = f.Family"}))
            If .BOF Or .EOF Then
                famIfValid = ""
            Else
                famIfValid = .Fields("Family").Value
            End If
        End With
    End Function

    Public Function famVsGenius(itNum As String,
    Optional mdFam As String = ""
) As String
        Dim ckFam As String
        Dim gnFam As String
        Dim ck As VbMsgBoxResult

        ' get current family from
        ' Genius, if it has one
        gnFam = famInGenius(itNum)

        If Len(gnFam) = 0 Then 'no family in Genius
            famVsGenius = mdFam 'so just use the model's
        Else 'need to check the model against it
            ' first, verify model family
            ckFam = famIfValid(mdFam)
            ' if not in Genius...

            If Len(ckFam) = 0 Then 'no need to ask
                famVsGenius = gnFam
            ElseIf gnFam = ckFam Then 'it's good
                famVsGenius = ckFam
            Else 'check with user
                ck = MsgBox(Join({
                "Item " & itNum,
                "Model Part Family " & ckFam & " differs",
                "from Genius Part Family " & gnFam, "",
                "Change Model to match Genius?", "",
                "(click [CANCEL] to debug)"
            }, vbCrLf),
                vbYesNoCancel + vbQuestion,
                "Use Genius Family?"
            )

                If ck = vbCancel Then
                    Stop 'to debug
                ElseIf ck = vbYes Then 'match Genius
                    famVsGenius = gnFam
                Else 'keep model Family
                    famVsGenius = ckFam
                End If
            End If
        End If
        '
    End Function


    Public Function msg_2022_0603_1127(
        dc As Scripting.Dictionary
    ) As String

        ' Function to handle a message and return a string.
        ' The Dictionary 'dc' likely contains message parameters.

        Dim result As String = "" ' Initialize the result string.

        Try
            ' Check if the dictionary is not nothing and contains the necessary keys.
            If Not dc Is Nothing AndAlso dc.ContainsKey("MessageType") Then

                Select Case dc("MessageType")

                    Case "TypeA" ' Example MessageType

                        ' Handle logic for MessageType "TypeA"
                        ' Example: Retrieve a value from the dictionary
                        If dc.ContainsKey("ValueA") Then
                            result = "TypeA: " & dc("ValueA").ToString()
                        Else
                            result = "TypeA: ValueA not found"
                        End If

                    Case "TypeB" ' Example MessageType

                        ' Handle logic for MessageType "TypeB"
                        If dc.ContainsKey("ValueB") Then
                            result = "TypeB: " & dc("ValueB").ToString()
                        Else
                            result = "TypeB: ValueB not found"
                        End If

                    Case Else
                        ' Handle unknown MessageTypes
                        result = "Unknown MessageType: " & dc("MessageType").ToString()
                End Select

            Else
                ' Handle cases where the dictionary is invalid or missing the MessageType key
                result = "Error: Invalid dictionary or missing MessageType."
            End If

        Catch ex As Exception
            ' Handle any exceptions that may occur
            result = "Error: " & ex.Message
        End Try

        Return result

    End Function

    Public Function askUserForPartMatl(
    AiDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' askUserForPartMatl -- Prompt User
        '     for Part Family and Material
        '     Selection, returning result
        '     in Dictionary
        '
        ' REV[2022.08.29.1621]
        ' add optional Dictionary parameter to which
        ' data from this function can be added.
        ' (see also askUserForMatlQty)
        '
        Dim rt As Scripting.Dictionary
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData
        Dim tx As String

        If dc Is Nothing Then
            rt = askUserForPartMatl(
        AiDoc, New Scripting.Dictionary)
        Else
            rt = dc 'New Scripting.Dictionary

            With rt
                If .Exists(pnFamily) Then .Remove(pnFamily)
                If .Exists(pnRawMaterial) Then .Remove(pnRawMaterial)

                If AiDoc Is Nothing Then
                    .Add(pnFamily, "")
                    .Add(pnRawMaterial, "")
                Else
                    With newFmTest1()
                        bd = nuAiBoxData().UsingInches.SortingDims(
                        AiDoc.ComponentDefinition.RangeBox
                    )
                        ck = .AskAbout(AiDoc,
                        "No Stock Found! Please Review" _
                        & vbCrLf & vbCrLf & bd.Dump(0)
                    )

                        If ck = vbYes Then
                            'Stop 'because this will
                            'override supplied Dictionary!
                            ' rt =
                            With .ItemData()
                                rt.Add(pnFamily, .Item(pnFamily))
                                rt.Add(pnRawMaterial, .Item(pnRawMaterial))
                            End With
                        Else
                            ' rt = New Scripting.Dictionary
                            Stop

                            With AiDoc.Propertys
                                tx = .Item(gnDesign
                                ).Item(pnFamily
                            ).Text
                                rt.Add(pnFamily, tx)

                                On Error Resume Next
                                Err.Clear()
                                tx = .Item(gnCustom
                                ).Item(pnRawMaterial
                            ).Text
                                If Err.Number Then tx = ""
                                On Error GoTo 0
                                rt.Add(pnRawMaterial, tx)
                            End With
                        End If
                    End With
                End If
            End With
        End If

        askUserForPartMatl = rt
    End Function

    Public Function askUserForMatlQty(
    AiDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' askUserForMatlQty -- Prompt User
        '     for Material Quantity and Units,
        '     returning result in Dictionary
        '
        ' REV[2022.08.29.1624]
        ' add optional Dictionary parameter to which
        ' data from this function can be added.
        ' (see also askUserForPartMatl)
        '
        Dim rt As Scripting.Dictionary
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData
        Dim tx As String

        If dc Is Nothing Then
            rt = askUserForMatlQty(
        AiDoc, New Scripting.Dictionary)
        Else
            rt = dc 'New Scripting.Dictionary
            With rt
                If .Exists(pnRmQty) Then .Remove(pnRmQty)
                If .Exists(pnRmUnit) Then .Remove(pnRmUnit)
            End With

            If AiDoc Is Nothing Then
                rt.Add(pnRmQty, 0)
                rt.Add(pnRmUnit, "")
            Else
                With nu_fmIfcMatlQty01().SeeUserWithPart(AiDoc)
                    ' following copied from dcGeniusPropsPartRev20180530 line 1632~?
                    If .Exists(pnRmQty) Then
                        rt.Add(pnRmQty, .Item(pnRmQty))
                        ' REV[2022.08.29.1459]
                        ' removing extraneous comments left over
                        ' from dcGeniusPropsPartRev20180530
                    Else
                        'Stop
                    End If

                    If .Exists(pnRmUnit) Then
                        rt.Add(pnRmUnit, .Item(pnRmUnit))
                    Else
                        'Stop
                    End If
                End With
            End If
        End If

        askUserForMatlQty = rt
    End Function

    Public Function askUserForPartMatlUpdate(
    AiDoc As Inventor.PartDocument
) As Scripting.Dictionary
        '
        ' askUserForPartMatlUpdate --
        '     Attempt to update Part Document
        '     material Properties from results
        '     of askUserForPartMatl
        '         (Family and Material Selection)
        '     and askUserForMatlQty
        '         (Material Quantity and Units)
        '     Return Dictionary of results
        '
        ' NOTE[2022.08.29.1627]
        ' want to separate user data collection
        ' from property updates in this function.
        ' further review/development called for.
        '
        Dim dcPr As Scripting.Dictionary
        Dim dcWk As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ck As VbMsgBoxResult
        Dim ky As Object

        dcPr = dcOfPropsInAiDoc(AiDoc)
        ck = vbOK
        If Not dcPr.Exists(pnRawMaterial) Then
            ck = MsgBox(Join({
            "Custom Property " & pnRawMaterial & ",",
            "used to identify Raw Material,",
            "is not yet present in this model.",
            "",
            "Go ahead and create it?"
        }, vbCrLf),
            vbYesNo + vbQuestion,
            "Required Property Missing!"
        )

            If ck = vbYes Then
                With AiDoc.Propertys.Item(gnCustom)
                    On Error Resume Next

                    For Each ky In {
                    pnRawMaterial, pnRmQty, pnRmUnit}

                    Next

                    Err.Clear()
                    InvProperty = .Add("", ky) 'pnRawMaterial
                    If Err.Number = 0 Then
                        dcPr.Add(ky, InvProperty) 'pnRawMaterial
                        ck = vbOK
                    Else
                        ck = vbAbort
                    End If
                    On Error GoTo 0
                End With
            Else
                ck = vbOK
            End If
        End If

        If ck <> vbOK Then
            ck = MsgBox(Join({
            "Custom Property " & pnRawMaterial & ",",
            "was not created! Raw Material",
            "will not be saved!"
        }, vbCrLf),
            vbOKCancel + vbExclamation,
            "Property Not Created!"
        )
        End If

        rt = New Scripting.Dictionary

        If ck = vbOK Then
            ' REV[2022.08.29.1616]
            ' condense two nearly identical With blocks
            ' into one, combining results of part material
            ' and material quantity data collections.
            ' NOTE: this required additional REVs
            ' to askUserForPartMatl (nee d4g1f1)
            ' and askUserForMatlQty (nee d4g1f3)
            ' to accept optional Dictionary to receive
            ' data points collected by each function.
            With askUserForMatlQty(AiDoc,
     askUserForPartMatl(AiDoc
))
                For Each ky In .Keys
                    With dcPr
                        If .Exists(ky) Then
                            InvProperty = .Item(ky)
                        Else
                            InvProperty = Nothing
                        End If
                    End With

                    If Not InvProperty Is Nothing Then
                        If Len(Trim(.Item(ky))) > 0 Then
                            If InvProperty.Text <> .Item(ky) Then
                                On Error Resume Next
                                Err.Clear()
                                'Stop 'so we can make sure this works
                                InvProperty.Text = .Item(ky)
                                Debug.Print("")  'Breakpoint Landing
                                'DON'T try to step at Property.Text
                                If Err.Number Then
                                    Stop
                                End If
                                On Error GoTo 0
                            End If
                        End If
                        rt.Add(ky, InvProperty.Text)
                    End If
                Next
            End With
        End If

        askUserForPartMatlUpdate = rt
    End Function

    Public Function dcGeniusPropsPartRev20180530_ck(
    invDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ' '
        ' NOTICE TO DEVELOPER [2021.11.12]
        ' '
        '
        ' This function definition was restored
        ' from a prior copy of this project
        ' (VB-000-1002_2021-1001.ipt)
        ' to restore current "normal" operation
        ' of the Genius Properties Update macro.
        ' The prior development version was
        ' retained for reference, renamed to
        ' dcGeniusPropsPartRev20180530_ck_broken
        '
        ' One minor revision was made to this
        ' restored version to retain improved
        ' generation of Genius Mass data.
        ' Additional changes should be kept
        ' to a MINIMUM to maintain correct
        ' operation going forward, and any
        ' desired changes implemented through
        ' some form of "shim"
        '
        ' '
        Dim rt As Scripting.Dictionary
        ' REV[2022.01.21.1351]
        ' Added following two Dictionaries
        Dim dcIn As Scripting.Dictionary
        ' to collect tings already in Genius
        Dim dcFP As Scripting.Dictionary
        ' to add a layer of separation
        ' to FlatPattern data collection
        ' (might not want to use for Properties
        '  so don't update immediately)

        ''
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        ''
        Dim prPartNum As Inventor.Property 'pnPartNum
        ' ADDED[2021.03.11] to simplify access
        ' to Part Number of Model, since it's
        ' requested several times in function
        Dim prFamily As Inventor.Property
        Dim prRawMatl As Inventor.Property 'pnRawMaterial
        Dim prRmUnit As Inventor.Property 'pnRmUnit
        Dim prRmQty As Inventor.Property 'pnRmQty
        ''
        Dim pnModel As String
        ' ADDED[2021.03.11] to further
        ' simplify access to Part Number
        Dim nmFamily As String
        Dim mtFamily As String
        ' UPDATE[2018.05.30]:
        '     Rename variable Family to nmFamily
        '     to minimize confusion between code
        '     and comment text in searches.
        '     Also add variable mtFamily
        '     for raw material Family name
        Dim pnStock As String
        Dim qtUnit As String
        Dim bomStruct As Inventor.BOMStructureEnum
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData

        If dc Is Nothing Then
            dcGeniusPropsPartRev20180530_ck =
        dcGeniusPropsPartRev20180530_ck(
            invDoc, New Scripting.Dictionary
        )
        Else
            rt = dc

            With invDoc
                ' Get Property s
                With .Propertys
                    aiPropsUser = .Item(gnCustom)
                    aiPropsDesign = .Item(gnDesign)
                End With

                ' Get Custom Properties
                prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)

                ' Part Number and Family properties
                ' are from Design, NOT Custom 
                prPartNum = aiGetProp(
                aiPropsDesign, pnPartNum)
                'ADDED 2021.03.11
                pnModel = prPartNum.Text
                prFamily = aiGetProp(
                aiPropsDesign, pnFamily)

                ' We should check HERE for possibly misidentified purchased parts
                ' UPDATE[2018.02.06]: Using new UserForm, see below
                With .ComponentDefinition
                    ' Request #1: Get the Mass in Pounds
                    ' and add to Custom Property GeniusMass
                    With .MassProperties
                        ' Update [2021.11.12]
                        '     System.Math.Round mass to nearest ten-thousandth
                        '     to try to match expected Genius value.
                        '     This should reduce or minimize reported
                        '     discrepancies during ETM process.
                        rt = dcWithProp(aiPropsUser, pnMass,
                        System.Math.Round(cvMassKg2LbM * .Mass, 4), rt
                    )
                    End With

                    '
                    ' Get BOM Structure type, correcting if appropriate,
                    ' and prepare Family value for part, if purchased.
                    '
                    ck = vbNo
                    ' UPDATE[2018.05.31]: Combined both InStr checks
                    ' by addition to generate a single test for > 0
                    ' If EITHER string match succeeds, the total
                    ' SHOULD exceed zero, so this SHOULD work.
                    If InStr(1, invDoc.FullFileName,
                    "\Doyle_Vault\Designs\purchased\"
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                    "|" & prFamily.Text & "|"
                ) > 0 Then
                        ' UPDATE[2018.02.06]: Using same
                        '     new UserForm as noted above.
                        ck = newFmTest2().AskAbout(invDoc, ,
                        "Is this a Purchased Part?" _
                        & vbCrLf & "(Cancel to debug)"
                    )
                    End If

                    ' Check process below replaces duplicate check/responses above.
                    If ck = vbCancel Then
                        Stop
                    ElseIf ck = vbYes Then
                        If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                            On Error Resume Next
                            .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                            If Err.Number = 0 Then
                                bomStruct = .BOMStructure
                            Else
                                bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                                ' WARNING: NOT a good way to go about this
                                '     but will go with it for now
                            End If
                            On Error GoTo 0
                        Else
                            bomStruct = .BOMStructure 'to make sure this is captured
                        End If
                    Else
                        bomStruct = .BOMStructure 'to make sure this is captured
                    End If

                    'Request #2: Change Cost Center iProperty.
                    'If BOMStructure = Purchased and not content center,
                    'then Family = D-PTS, else Family = D-HDWR.
                    '
                    'UPDATE[2018-05-30]: Value produced here
                    'will now be held for later processing,
                    'more toward the end of this function.
                    If bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                        If .IsContentMember Then
                            nmFamily = "D-HDWR"
                        Else
                            nmFamily = "D-PTS"
                            'NOTE: NON Content Center members
                            '       might still be D-HDWR
                            '       Additional checks might
                            '       be recommended
                        End If
                    Else
                        nmFamily = ""
                    End If
                End With
                ' At this point, nmFamily SHOULD be 
                ' to a non-blank value if Item is purchased.
                ' We should be able to check this later on,
                ' if Item BOMStructure is NOT Normal

                'Request #4: Change Cost Center iProperty.
                'If BOMStructure = Normal, then Family = D-MTO,
                'else if BOMStructure = Purchased then Family = D-PTS.
                If bomStruct = BOMStructureEnum.kNormalBOMStructure Then

                    ' REV[2022.01.28.1014]
                    ' Added initial raw material capture
                    ' to check against Genius
                    ' HOLD![2022.01.28.1046]
                    ' commenting out again
                    ' probably best below, still
                    pnStock = prRawMatl.Text
                    ' REV[2022.02.08.1304]
                    ' restored, to obtain any
                    ' value already defined.
                    ' MIGHT need moved further down,
                    ' but hold off on that for now.

                    ' REV[2022.01.17.1123]
                    ' Start adding code to capture
                    ' any raw material items for
                    ' part already in Genius.
                    ' REV[2022.01.21.1357]
                    ' Separated capture from With statement
                    ' into new Dictionary object in order
                    ' to check and use it further down,
                    ' as well as passing it to nuSelFromDict
                    ' to handle multiple line items
                    ' REV[2022.01.31.1008]
                    ' Restored assignment of dcFromAdoRS
                    ' result to Dictionary Object dcIn,
                    ' in order to pass it to other
                    ' functions, as needed.
                    '
                    dcIn = DcFromAdoRS(CnGnsDoyle().Execute(
                    sqlOf_GnsPartMatl(pnModel)
                ))
                    ''Debug.Print(ConvertToJson(dcRecSetDcDx4json(dcIn), vbTab)
                    'Stop
                    ' dcIn = dcOb(dcRecSetDcDx4json(dcIn).Item(pnRawMaterial))
                    If dcIn.Count > 0 Then 'Genius found something
                        With dcOb(dcRecSetDcDx4json(dcIn).Item(pnRawMaterial))
                            ' REV[2022.01.28.1336]
                            ' Added code to collect captured
                            ' dcIn = New Scripting.Dictionary


                            ' REV[2022.01.28.0857]
                            ' Added code to collect captured
                            ' material item number, asking user
                            ' to select from list if more than one.
                            If .Count > 0 Then 'Genius found something
                                If Len(pnStock) > 0 Then
                                    'some material already assigned
                                    If .Exists(pnStock) Then 'do nothing
                                        'it's a valid option, stick with it
                                        '
                                    Else 'probably going to need an update
                                        'so forget current value (for now)
                                        pnStock = ""
                                    End If
                                End If

                                If Len(pnStock) = 0 Then
                                    'grab first material item found
                                    'Stop
                                    'pnStock = dcOb(.Item(.Keys(0))).Item(pnRawMaterial)
                                    pnStock = .Keys(0)
                                    ' REV[2022.02.08.1336]
                                    ' switched back to pulling first Key.
                                    ' since the With statement pulls the
                                    ' Dictionary keyed on raw materials,
                                    ' this SHOULD be reliably correct.
                                    '
                                    ' that the nuSelector below is called
                                    ' with the list of Keys would seem
                                    ' to support this expectation.
                                End If

                                'and use it for the default...
                                If .Count > 1 Then
                                    Stop 'because selection is going
                                    'to be a lot more complicated.
                                    '(just look at that pnStock
                                    ' assignment up there!)

                                    pnStock = nuSelector(
                                ).GetReply(.Keys, pnStock)

                                    Stop 'to make sure things are okay
                                End If
                            Else 'do nothing
                                ' REV[2022.02.08.1353]
                                ' disabled pnStock assignment
                                ' from prRawMatl here, since
                                ' this is now done automatically
                                ' (see REV[2022.02.08.1304] above)
                                'pnStock = prRawMatl.Text
                            End If

                            ' REV[2022.01.28.0903]
                            ' Separated Dictionary capture
                            ' from Count check
                            If Len(pnStock) > 0 Then
                                If Len(CStr(prRawMatl.Text)) = 0 Then 'don't worry.
                                    'it'll be taken care of further down
                                ElseIf pnStock = prRawMatl.Text Then 'don't worry.
                                    'should only be minor quantity changes
                                    'Stop 'and make sure we want to do this.

                                    ' dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                    'Deactivated, moved down and out of this If-Then nest.
                                    'Search below for active copy

                                    'Debug.Print("") 'Breakpoint Landing
                                Else 'need to ask User what to go with
                                    Debug.Print("=== CURRENT GENIUS MATERIAL DATA ===")
                                    'Debug.Print(dumpLsKeyVal(dcIn, ":" & vbTab)
                                    ck = newFmTest2().AskAbout(invDoc,
                                    "Raw Material " & prRawMatl.Text _
                                    & vbCrLf & " for Item" _
                                    ,
                                    "does not match " & pnStock _
                                    & vbCrLf & "indicated in Genius." _
                                    & vbCrLf & vbCrLf _
                                    & "Change to match Genius?" _
                                    & vbCrLf & "(Cancel to debug)"
                                )
                                    If ck = vbCancel Then
                                        Stop 'to check things out
                                    ElseIf ck = vbNo Then
                                        ' NOTE[2022.02.08.1359]
                                        ' DO NOT DISABLE this instance
                                        ' of the pnStock assignment!
                                        pnStock = prRawMatl.Text
                                        ' this one implements the user's
                                        ' decision NOT to change the
                                        ' current material assignment,
                                        ' by forcing pnStock back to the
                                        ' original Value from prRawMatl
                                    End If

                                    'Stop 'to grab current raw material item
                                    ' NOTE: Since material data already
                                    ' in Genius is now captured in dcIn,
                                    ' the following assignments are NOT
                                    ' immediately necessary.
                                    'prRawMatl.Text = dcOb(.Item(.Keys(0))).Item(pnRawMaterial) 'pnStock
                                    'prRmQty.Text = CStr(dcOb(.Item(.Keys(0))).Item(pnRmQty))
                                    'prRmUnit.Text = dcOb(.Item(.Keys(0))).Item(pnRmUnit)
                                    '= dcOb(.Item(.Keys(0))).Item("MtFamily")
                                    ' In fact, these Properties should
                                    ' first be checked against Genius
                                    ' data, and the user prompted
                                    ' to go ahead with updates.
                                End If

                                ' REV[2022.01.28.1448]
                                ' Changed data extraction process here
                                ' to work with form returned from dcFromAdoRS
                                '
                                ' NOTE! This is !!!TEMPORARY!!!
                                ' Implemented during run time,
                                ' some truly insane acrobatics were required
                                ' to make it work without reting the run.
                                ' This code, including the With statement
                                ' above, MUST be rewritten as soon as feasible!
                                '
                                'Stop 'because we're doing to need to do something different
                                ''Debug.Print(ConvertToJson(dcRecSetDcDx4json(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial), vbTab)
                                ''Debug.Print(ConvertToJson(dcOb(dcRecSetDcDx4json(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock), vbTab)
                                ''Debug.Print(ConvertToJson(dcOb(.Item(dcOb(dcOb(dcRecSetDcDx4json(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0))), vbTab)
                                'dcOb(dcOb(dcRecSetDcDx4json(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)
                                'Stop

                                If .Exists(pnStock) Then
                                    dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                    'This is DEFINITELY going to need a rework!
                                    'But that will need a new function, most likely

                                    'deactivated the version below
                                    'to be superceded by the one above
                                    ' dcIn = dcOb(.Item(dcOb(dcOb(dcRecSetDcDx4json(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)))

                                    'original version, also deactivated
                                    'for obvious reasons
                                    ' dcIn = .Item(pnStock)

                                    Debug.Print("") 'Breakpoint Landing

                                    ''Debug.Print(ConvertToJson(dcIn, vbTab)
                                    'Stop
                                Else
                                    Stop 'because we've got a REAL problem here!
                                    ' THINK pnStock should ALWAYS be in the
                                    ' With contextual Dictionary at this point,
                                    ' so if it isn't, something probably went
                                    ' seriously wrong
                                End If
                            Else
                                dcIn = New Scripting.Dictionary
                            End If
                        End With
                    End If

                    With dcIn
                        If .Count = 0 Then
                            .Add("Ord", 0)
                            .Add("RM", "")
                            .Add("MtFamily", "")
                            .Add("RMQTY", 0)
                            .Add("RMUNIT", "")
                            '.Add("", ""
                        End If : End With

                    '----------------------------------------------------'
                    If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                        '----------------------------------------------------'
                        '
                        ' NOTE[2018-05-31]: At this point, we MAY wish
                        ' to check for a valid flat pattern,
                        ' and otherwise attempt to verify
                        ' an actual sheet metal design.
                        '

                        ' REV[2022.01.28.0903]
                        ' HERE is where things start to get interesting
                        ' Before processing Part as sheet metal,
                        ' want to make sure it's supposed to be.
                        '
                        ' FIRST, check what Genius had to say
                        With dcIn
                            If .Exists("MtFamily") Then
                                mtFamily = .Item("MtFamily")
                            Else
                                mtFamily = ""
                            End If
                        End With

                        If Len(mtFamily) = 0 Then 'need more info
                            ck = vbRetry
                        ElseIf mtFamily = "DSHEET" Then 'Genius says...
                            ck = vbYes
                        Else
                            ck = vbNo
                        End If

                        ' REV[2022.01.31.1335]
                        ' Move flat pattern collection out here
                        ' from inside the next If-Then block
                        If ck = vbNo Then 'we don't want/need flat pattern
                            dcFP = New Scripting.Dictionary
                            ' we're going to want to do
                            ' something different here
                        Else 'we MIGHT, so lets get it
                            dcFP = dcFlatPatVals(.ComponentDefinition)
                            ' try to get flat pattern data
                            ' WITHOUT mucking up Properties!
                            ' Want to avoid dirtying file with
                            ' changes until absolutely necessary)

                            If dcFP.Exists(pnThickness) Then
                                pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                                dcFP.Add(pnRawMaterial, pnStock)
                                'need to change this to use the following more directly
                                'sqlSheetMetal(.ActiveMaterial.DisplayName,CDbl(dcfp.Item(pnThickness)))
                            End If
                        End If
                        Debug.Print("") 'Breakpoint Landing
                        If False Then
                            ' 'Debug.Print(ConvertToJson({dcIn, dcFP}, vbTab))
                        End If

                        If ck = vbRetry Then 'Genius check is inconclusive
                            ' so let's see what the flat pattern can tell us

                            If dcFP.Exists("mtFamily") Then
                                If dcFP.Item("mtFamily") = "DSHEET" Then
                                    If dcFP.Exists("OFFTHK") Then
                                        Stop
                                        ck = newFmTest2().AskAbout(invDoc, "This Part: ", "might not be sheet metal. " & vbCrLf & vbCrLf & "Is it in fact sheet metal?")
                                        If ck = vbCancel Then
                                            ck = vbRetry
                                            Stop 'to debug
                                        End If
                                    Else
                                        ck = vbYes
                                    End If
                                ElseIf dcFP.Item("mtFamily") = "D-BAR" Then
                                    ck = vbNo
                                Else
                                    ck = vbRetry
                                End If
                            Else
                                ck = vbRetry
                            End If
                        End If

                        If ck = vbRetry Then
                            '  'Debug.Print(ConvertToJson({dcIn, dcFP}, vbTab))
                            Stop 'so we can figure out what to do next.
                            'for now, most likely just press [F5]
                            'to continue
                        End If

                        'Request #3:
                        '   Get sheet metal extent area
                        '   and add to custom property "RMQTY"

                        ' REV[2022.01.28.1556]
                        ' change if-then-else sequence
                        ' to check ck instead of dcIn
                        If ck = vbYes Then
                            rt = dcFlatPatProps(.ComponentDefinition, rt)
                            ' NOTE[2022.01.28.1551]:
                            ' THIS call should change to move
                            ' values from dcFP into Properties.
                            ' Probably a new function, which might
                            ' ALSO be called from dcFlatPatProps
                        ElseIf ck = vbRetry Then
                            rt = dcFlatPatProps(.ComponentDefinition, rt)
                        ElseIf ck = vbNo Then 'probably
                            'don't do anything here
                        Else 'we got a problem
                            'material type detection SHOULD produce
                            'one of the three preceding values

                            Stop 'and check it out
                        End If

                        ' NOTE[2018-05-30]:
                        '     Raw Material Quantity value
                        '     SHOULD be  upon return
                        '     We may need to review the process
                        '     to find an appropriate place
                        '     to  for NON sheet metal

                        'Moved to start of block to check for NON sheet metal

                        'NOTE: THIS call might best be combined somehow
                        '   with the flat pattern prop pickup above.
                        '   Note especially that if dcFlatPatProps
                        '   FINDS NO .FlatPattern, then there should
                        '   BE NO sheet metal part number!
                        If prRawMatl Is Nothing Then
                            If rt.Exists("OFFTHK") Then
                                ' NOTE[2021.12.10]:
                                '     Believe this OFFTHK property is meant
                                '     to capture "Sheet Metal" Parts that
                                '     aren't actually Sheet Metal.
                                '     This check might be needed further down.
                                ' UPDATE[2018.05.30]:
                                '     Restoring original key check
                                '     and adding code for debug
                                '     Previously changed to "~OFFTHK"
                                '     to avoid this block and its issues.
                                '     (Might re-revert if not prepped to fix now)
                                Debug.Print(aiProperty(rt.Item("OFFTHK")).Text)
                                Stop 'because we're going to need to do something with this.

                                pnStock = "" 'Originally the ONLY line in this block.
                                ' A more substantial response is required here.

                                If 0 Then Stop '(just a skipover)
                            Else
                                Stop 'because we don't know IF this is sheet metal yet
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If
                        Else
                            ''  ACTION ADVISED[2018.09.14]:
                            ''  pnStock can probably be 
                            ''  to prRawMatl.Text and THEN
                            ''  checked for length to see
                            ''  if lookup needed.
                            ''  This might also allow us to check
                            ''  for machined or other non-sheet
                            ''  metal parts.

                            ' REV[2021.12.17]: sanity check
                            '     Add sanity check to make sure
                            '     any existing sheet metal stock
                            '     number matches model specs
                            If Len(prRawMatl.Text) > 0 Then
                                ' we need to check it

                                If Len(pnStock) = 0 Then
                                    ' REV[2022.01.28.1445]:
                                    ' Placed this pnStock stock assignment
                                    ' inside this If-Then block to prevent
                                    ' overriding value from Genius
                                    pnStock = ptNumShtMetal(.ComponentDefinition)
                                End If
                                ' NOTE[2021.12.17@15:32]:
                                '     copied this up from
                                '     NOTE[2021.12.17@15:32]
                                '     for use in sanity check

                                ' NOTE[2021.12.17]:
                                '     This section simply warns the user
                                '     that the current raw material does
                                '     not match the recommended default,
                                '     and offers an opportunity to fix it.
                                '
                                '     This is yet another quick and dirty
                                '     "solution" that should be revised
                                ' NOTE[2022.01.05]:
                                '     Adding check for empty recommendation.
                                '     Do NOT believe user should be offered
                                '     opportunity to overwrite any current
                                '     part number with a BLANK one. Believe
                                '     the option to CLEAR is somewhere below.
                                If Len(pnStock) > 0 Then
                                    If pnStock <> prRawMatl.Text Then
                                        'Stop

                                        ' NOTE[2022.01.03]:
                                        '     Following text SHOULD no longer
                                        '     be needed. Verify function of
                                        '     FmTest2 following, and when good,
                                        '     disable and/or remove this block.
                                        Debug.Print("!!! NOTICE !!!")
                                        Debug.Print("Recommended Sheet Metal Stock (" & pnStock & ")")
                                        Debug.Print("does not match current Stock (" & prRawMatl.Text & ")")
                                        Debug.Print("")
                                        Debug.Print("To continue with no change, just press [F5]. Otherwise,")
                                        Debug.Print("press [ENTER] on the following line first to change:")
                                        Debug.Print("prRawMatl.Text = """ & pnStock & """")
                                        Debug.Print("")

                                        ' NOTE[2022.01.03]:
                                        '     Now using FmTest2(?) to prompt
                                        '     user as in other checks (above?)
                                        ck = newFmTest2().AskAbout(invDoc,
                                        "Suggest Sheet Metal change from" _
                                        & vbCrLf & prRawMatl.Text & " to" _
                                        & vbCrLf & pnStock & " for",
                                        "Change it?"
                                    )
                                        If ck = vbCancel Then
                                            Stop 'to check things out
                                        ElseIf ck = vbYes Then
                                            'Stop
                                            prRawMatl.Text = pnStock
                                        End If
                                        'Stop
                                    End If
                                End If
                            ElseIf Len(pnStock) > 0 Then
                                'go ahead and assign material
                                prRawMatl.Text = pnStock
                                ' REV[2022.02.08.1406]
                                ' added new branch to assign pnStock,
                                ' if not blank, to prRawMatl, if it is.
                            End If

                            If Len(prRawMatl.Text) > 0 Then
                                If rt.Exists("OFFTHK") Then
                                    'Stop 'and verify raw material item
                                    ' NOTE[2021.12.13]:
                                    '     OFFTHK property check added
                                    '     to catch sheet metal already
                                    '     assigned by accident.
                                    ck = newFmTest2().AskAbout(invDoc,
                                    "Assigned Raw Material " & prRawMatl.Text _
                                    & vbCrLf & " might be incorrect for ",
                                    "Clear it?"
                                )
                                    If ck = vbCancel Then
                                        Stop 'to check things out
                                    ElseIf ck = vbYes Then
                                        prRawMatl.Text = ""
                                    End If
                                    'Stop
                                End If


                                If pnStock = prRawMatl.Text Then
                                    'no need to assign it again
                                    Debug.Print("") 'Breakpoint Landing
                                Else 'need to check things out...
                                    ' 'Debug.Print(ConvertToJson({pnStock, prRawMatl.Text})) 'and...
                                    Stop 'before we do something stupid!
                                    pnStock = prRawMatl.Text
                                End If

                                ' The following With block copied and modified [2021.03.11]
                                ' from elsewhere in this function as a temporary measure
                                ' to address a stopping situation later in the function.
                                ' See comment below for details.
                                '
                                With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                    If .BOF Or .EOF Then
                                        If pnStock <> "0" Then
                                            Stop 'because Material value likely invalid
                                            ' REV[2022.03.01.1553]
                                            ' embedded in check
                                            ' for string value "0"
                                            ' as this seems to come
                                            ' up as a legacy issue,
                                            ' and is readily remedied
                                            ' in this section. No stop
                                            ' is needed in that case.
                                        End If
                                        ' REV[2022.02.08.1413]
                                        ' reinstated interruption here
                                        ' because at this point, pnStock
                                        ' has likely already been assigned
                                        ' to prRawMatl, so changing it here
                                        ' is NOT likely to be productive.
                                        ' this section will likely need
                                        ' reconsideration, revision,
                                        ' and/or possibly removal.
                                        ' UPDATE[2021.12.10]:
                                        '     added this check for OFFTHK
                                        '     to avoid blindly adding sheet
                                        '     metal stock to a "sheet metal"
                                        '     part that isn't actually meant
                                        '     to be made of sheet metal.
                                        If rt.Exists("OFFTHK") Then 'likely NOT
                                            'actual Sheet Metal, so just clear this:
                                            pnStock = ""
                                        Else
                                            pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                                            Debug.Print("") 'Breakpoint Landing
                                            ' UPDATE[2021.12.10]:
                                            '     embedded this call in the OFFTHK
                                            '     check noted above. see that note
                                            '     for details
                                            '
                                            ''  ACTION TAKEN[2021.03.11]:
                                            ''  temporary measure to try to ensure
                                            ''  recovered material Item number is valid,
                                            ''  and if not, to fix it automatically.
                                            ''  This seeks to address a stop situation
                                            ''  later in this function, encountered
                                            ''  when the Part Number property is neither
                                            ''  blank NOR valid (typically "0"), likely
                                            ''  as a result of an uninitialized iPart property.
                                            ''  (see ACTION ADVISED[2018.09.14] elsewhere)
                                        End If
                                    Else
                                        ''  This section retained from source,
                                        ''  but disabled to avoid potential issues
                                        ''  with subsequent operations, just in case
                                        ''  anything depends on mtFamily remaining
                                        ''  uninitialized up to that point.
                                        ''
                                        ''  With .Fields
                                        ''      mtFamily = .Item("Family").Value
                                        ''  End With
                                    End If
                                End With
                                '
                                ' This section likely should be removed when primary issue resolved.
                                '
                            ElseIf rt.Exists("OFFTHK") Then
                                ' UPDATE[2021.12.10]:
                                '     another OFFTHK check added to avoid
                                '     adding sheet metal stock by mistake.
                                pnStock = ""
                                ' NOTE[2021.12.10]:
                                ' by keeping this value blank, it is hoped
                                ' to force the User to select the appropriate
                                ' raw material item, rather than assigning
                                ' a sheet metal item by mistake, just because
                                ' it matches the Part's defined Thickness.
                                '
                                ' if the measured height of the flat pattern
                                ' doesn't closely match the defined Thickness,
                                ' the part is most likely NOT sheet metal.
                                '
                                ' it might be most appropriate to move the
                                ' OFFTHK check outside and above the others
                                ' in this sequence, as it likely determines
                                ' whether ANY so-called sheet metal part
                                ' should actually be treated as such.
                                '
                            Else
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                                ' UPDATE[2021.12.10]:
                                '     just as before, embedded this call
                                '     in another OFFTHK check for the same
                                '     reason noted above
                                ' NOTE[2021.12.17@15:32]:
                                '     copying this up to ...
                            End If

                            If Len(pnStock) = 0 Then
                                ' UPDATE[2018.05.30]:
                                '     Pulling ALL code/text from this section
                                '     to get rid of excessive cruft.
                                '
                                '     In fact, reversing logic to go directly
                                '     to User Prompt if no stock identified
                                '
                                '     IN DOUBLE FACT, hauling this WHOLE MESS
                                '     RIGHT UP after initial pnStock assignment
                                '     to prompt user IMMEDIATELY if no stock found
                                With newFmTest1()
                                    If Not invDoc.ComponentDefinition.Document Is invDoc Then Stop

                                    bd = nuAiBoxData().UsingInches.SortingDims(
                                    invDoc.ComponentDefinition.RangeBox
                                )
                                    ck = .AskAbout(invDoc,
                                    "No Stock Found! Please Review" _
                                    & vbCrLf & vbCrLf & bd.Dump(0)
                                )

                                    If ck = vbYes Then
                                        ' UPDATE[2018.05.30]:
                                        '     Pulling some extraneous commented code
                                        '     from here and beginning of block
                                        With .ItemData
                                            If .Exists(pnFamily) Then
                                                nmFamily = .Item(pnFamily)
                                                Debug.Print(pnFamily & "=" & nmFamily)
                                            End If

                                            If .Exists(pnRawMaterial) Then
                                                pnStock = .Item(pnRawMaterial)
                                                Debug.Print(pnRawMaterial & "=" & pnStock)
                                            End If
                                        End With
                                        If 0 Then Stop 'Use this for a debugging shim
                                    End If
                                End With
                            ElseIf Left$(pnStock, 2) = "LG" Then 'it's probably lagging
                                Debug.Print(pnModel & ": PROBABLE LAGGING [" & pnStock & "]")
                                Debug.Print("  TRY TO VERIFY. IF CHANGE REQUIRED,")
                                Debug.Print("  FILL IN NEW VALUE FOR pnStock BELOW, ")
                                Debug.Print("  AND PRESS ENTER ON THE LINE. WHEN ")
                                Debug.Print("  READY, PRESS [F5] TO CONTINUE.")
                                Debug.Print("  pnStock = """ & pnStock & """")
                                Stop
                            End If

                            If Len(pnStock) > 0 Then 'and ONLY then
                                'do we look for a Raw Material Family!

                                With CnGnsDoyle().Execute(
                                "select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                    If .BOF Or .EOF Then
                                        Stop 'because Material value likely invalid
                                        ''  ACTION ADVISED[2018.09.14]:
                                        ''  Will need to address this situation
                                        ''  in a more robust manner.
                                        ''  A more thorough query above
                                        ''  might also be called for.
                                    Else
                                        With .Fields
                                            mtFamily = .Item("Family").Value
                                        End With

                                        ' UPDATE[2021.06.18]:
                                        '     New pre-check for Material Item
                                        '     in Purchased Parts Family.
                                        '     VERY basic handler simply
                                        '     maps Material Family to D-BAR
                                        '     to force extra processing below.
                                        '     Further refinement VERY much needed!
                                        If mtFamily Like "?-MT*" Then
                                            'Debug.Print(pnModel & " [" & prRawMatl.Text & "]: " & aiPropsDesign(pnDesc).Text
                                            Debug.Print(pnModel & "[" & prRmQty.Text & qtUnit & "*" & pnStock & ": " & aiPropsDesign.item(pnDesc).Text & "]") ' prRawMatl.Text
                                            Stop 'FULL Stop!
                                        ElseIf mtFamily = "D-PTS" Then
                                            nmFamily = "D-RMT"
                                            Stop 'NOT SO FAST!
                                            mtFamily = "D-BAR"
                                        ElseIf mtFamily = "R-PTS" Then
                                            nmFamily = "R-RMT"
                                            Stop 'NOT SO FAST!
                                            mtFamily = "D-BAR"
                                        End If

                                        If mtFamily = "DSHEET" Then
                                            'We should be okay. This is sheet metal stock
                                            nmFamily = "D-RMT"
                                            qtUnit = "FT2"
                                            ' UPDATE[2018.05.30]:
                                            '     Moving part family assignment
                                            '     to this section for better mapping
                                            '     and updating to new Family names
                                            '     as well as pulling up qtUnit assignment
                                        ElseIf mtFamily = "D-BAR" Then
                                            ' UPDATE[2021.06.18]:
                                            '     Added check for Part Family already 
                                            '     to more properly handle new situation (above)
                                            If Len(nmFamily) = 0 Then
                                                nmFamily = "R-RMT"
                                            Else
                                                Debug.Print("") 'Breakpoint Landing
                                                'Stop
                                            End If

                                            ' UPDATE[2022.01.11]:
                                            '     Adding Do..Loop Until to following section
                                            '     to allow user to retry ting material
                                            '     quantity and units. This change made in
                                            '     conjunction with new prompt form (below).
                                            ' NOTE! This is FIRST instance of revision
                                            '     Search on UPDATE text above to locate
                                            '     the other in this function
                                            qtUnit = prRmUnit.Text '"IN"
                                            ck = vbCancel
                                            Do

                                                ''may want function here
                                                ' UPDATE[2018.05.30]: As noted above
                                                '     Will keep Stop for now
                                                '     pending further review,
                                                '     hopefully soon
                                                'Debug.Print(pnModel & " [" & prRawMatl.Text & "]: " & aiPropsDesign(pnDesc).Text
                                                'Debug.Print(CDbl(dcIn.Item(pnRmQty))
                                                ' UPDATE[2021.03.11]: Replaced
                                                ' aiPropsDesign.Item(pnPartNum)
                                                ' with prPartNum (and now pnModel)
                                                ' since it's used in several places

                                                'Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,"
                                                'Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                                'Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW."

                                                ' REV[2022.02.08.1511]
                                                ' replaced boilerplate above with new version below
                                                ' in hopes of better presenting change options
                                                ' in a more compact and accessible form.

                                                Debug.Print("===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====")
                                                Debug.Print("  If change required, place new values at end")
                                                Debug.Print("  of lines below for prRmQty.Text and qtUnit.")
                                                Debug.Print("  Press [ENTER] on each line to be changed.")
                                                Debug.Print("  Press [F5] when ready to continue.")
                                                Debug.Print("----- " & pnModel & " [" & prRawMatl.Text & "]: " & aiPropsDesign.item(pnDesc).Text & " -----")
                                                'Debug.Print(""

                                                ' REV[2022.02.09.0923]
                                                ' replication of REV[2022.02.09.0919]
                                                ' from section below: prep to replace
                                                ' old dimension dump operation with more
                                                ' compact call to aiBoxData's Dump method
                                                If True Then 'go ahead and run old dump
                                                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                                    With invDoc.ComponentDefinition.RangeBox
                                                        Debug.Print("")
                                                        System.Math.Round((.MaxPoint.X - .MinPoint.X) / cvLenIn2cm, 4)
                                                        System.Math.Round((.MaxPoint.Y - .MinPoint.Y) / cvLenIn2cm, 4)
                                                        System.Math.Round((.MaxPoint.Z - .MinPoint.Z) / cvLenIn2cm, 4)
                                                    End With
                                                End If

                                                With nuAiBoxData().UsingInches().UsingBox(
                                            invDoc.ComponentDefinition.RangeBox
                                        )
                                                    Debug.Print(.Dump(0))
                                                End With
                                                'Stop 'and check output against prior version

                                                ' REV[2022.02.08.1446]
                                                ' removed block of Debug.Print(lines
                                                ' disabled now for some time, as they
                                                ' do not seem to have been missed.
                                                Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text), " 'in model. ")
                                                If dcIn.Exists(pnRmQty) Then Debug.Print("In Genius: ", CStr(dcIn.Item(pnRmQty)))
                                                Debug.Print("")
                                                Debug.Print("qtUnit = """, qtUnit, """ 'in model. ")
                                                If dcIn.Exists(pnRmUnit) Then Debug.Print("In Genius: ", CStr(dcIn.Item(pnRmUnit)))
                                                If dcIn.Item(pnRmUnit) <> "IN" Then Debug.Print(" ( or try IN )")
                                                Debug.Print("")
                                                'Debug.Print("qtUnit = ""IN"""
                                                'Debug.Print(""
                                                'Debug.Print(""
                                                'Debug.Print(""
                                                Stop 'because we might want a D-BAR handler
                                                ' Actually, we might NOT need to stop here
                                                ' if bar stock is already selected,
                                                ' because quantities would presumably
                                                ' have been established already.
                                                ' Any D-BAR handler probably needs
                                                ' to be implemented in prior section(s)
                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                                ck = newFmTest2().AskAbout(invDoc,
                                            "Raw Material Quantity is now " _
                                            & CStr(prRmQty.Text) & qtUnit & " for",
                                            "If this is okay, click [YES]. Otherwise," _
                                            & vbCrLf & "click [NO] or [CANCEL] to fix."
                                        )
                                                'Stop
                                            Loop Until ck = vbYes
                                            ' UPDATE[2022.01.11]:
                                            '     This is the terminal end of the
                                            '     Do..Loop Until block noted above

                                            rt = dcAddProp(prRmQty, rt)
                                            Debug.Print("") 'Landing line for debugging. Do not disable.
                                        Else
                                            nmFamily = ""
                                            qtUnit = "" 'may want function here
                                            ' UPDATE[2018.05.30]: As noted above
                                            '     However, might need more handling here.
                                            Stop 'because we don't know WHAT to do with it
                                        End If
                                    End If
                                End With
                            Else
                                If 0 Then Stop 'and regroup
                                ' Things are looking a right royal mess
                                ' at the moment I'm writing this comment.
                            End If
                        End If

                        With prRawMatl
                            If Len(Trim$(.Text)) > 0 Then
                                If pnStock <> .Text Then
                                    'Debug.Print("Raw Stock Selection"
                                    'Debug.Print("  Current : " & prRawMatl.Text
                                    'Debug.Print("  Proposed: " & pnStock
                                    'Stop 'because we might not want to change existing stock ting
                                    'if
                                    ck = MsgBox(
                                    Join({
                                        "Raw Stock Change Suggested",
                                        "  for Item " & pnModel,
                                        "",
                                        "  Current : " & prRawMatl.Text,
                                        "  Proposed: " & pnStock,
                                        "", "Change It?", ""
                                    }, vbCrLf),
                                    vbYesNo, pnModel & " Stock"
                                )
                                    '"Change Raw Material?"
                                    '"Suggested Sheet Metal"
                                    If ck = vbYes Then .Text = pnStock
                                End If
                            Else
                                .Text = pnStock
                            End If
                        End With
                        rt = dcAddProp(prRawMatl, rt)

                        With prRmUnit
                            If Len(.Text) > 0 Then
                                If Len(qtUnit) > 0 Then
                                    If .Text <> qtUnit Then
                                        'Stop 'and check both so we DON'T
                                        'automatically "fix" the RMUNIT value

                                        ck = newFmTest2().AskAbout(invDoc, ,
                                        "Raw Material " & prRawMatl.Text _
                                        & vbCrLf & "Unit of Measure currently " _
                                        & .Text & vbCrLf & vbCrLf _
                                        & "Change to " & qtUnit & "?" _
                                        & vbCrLf & " "
                                    )
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            .Text = qtUnit
                                        End If
                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        ' rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                        ' Plan to remove commented line above,
                        ' superceded by the one above that
                        Debug.Print("") 'Another landing line

                        '--------------------------------------------'
                    Else 'for standard Part (NOT Sheet Metal) ---'
                        '--------------------------------------------'
                        ' [2018.07.31 by AT]
                        ' Duped following block from above
                        ' to mod for material assignment
                        ' to non-sheet metal part.
                        '
                        ' Except, this isn't enough.
                        ' Also need the code to add
                        ' Stock PN to Attribute RM.
                        ' That's a whole 'nother
                        ' block of code, and likely
                        ' best consolidated.
                        With newFmTest1()
                            If Not invDoc.ComponentDefinition.Document Is invDoc Then Stop

                            ' [2018.07.31 by AT]
                            ' Added the following to try to
                            ' preselect non-sheet metal stock
                            '.dbFamily.Text = "D-BAR"
                            '.LbxFamily.Text = "D-BAR"
                            ' Doesn't quite do it.
                            'With New aiBoxData
                            'bd = nuAiBoxData().UsingInches.UsingBox(
                            'invDoc.ComponentDefinition.RangeBox(
                            '    ))
                            'bd = nuAiBoxData().UsingInches.SortingDims(
                            '        invDoc.ComponentDefinition.RangeBox
                            '    )
                            ''End With

                            ck = .AskAbout(invDoc,
                                    "Please Select Stock for Machined Part" _
                                    & vbCrLf & vbCrLf & bd.Dump(0)
                                )

                            If ck = vbYes Then
                                ' UPDATE[2018.05.30]:
                                '     Pulling some extraneous commented code
                                '     from here and beginning of block
                                With .ItemData
                                    If .Exists(pnFamily) Then
                                        nmFamily = .Item(pnFamily)
                                        Debug.Print(pnFamily & "=" & nmFamily)
                                    End If

                                    If .Exists(pnRawMaterial) Then
                                        pnStock = .Item(pnRawMaterial)
                                        Debug.Print(pnRawMaterial & "=" & pnStock)
                                    End If
                                End With
                                If 0 Then Stop 'Use this for a debugging shim
                                ''  We're going to need something here
                                ''  to make sure raw material gets added
                                ''  for non sheet metal parts, as well
                                ''  What we're going to need to do
                                ''  is refactor this whole bloody thing.
                            End If
                        End With
                        '
                        '
                        '
                        '
                        ' The following If block is copied
                        ' wholesale from sheet metal section above.
                        ' Some changes (to be) made to accommodate
                        ' machined or other non-sheet metal stock.
                        '
                        ' Ultimately, whole mess to require refactor.
                        '
                        If Len(pnStock) > 0 Then 'and ONLY then
                            'do we look for a Raw Material Family!

                            ' This enclosing With block should NOT be necessary
                            ' since the newFmTest1 above takes care of collecting
                            ' the Stock Family along with the Stock itself
                            With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    '
                                    ' Content formerly here moved BELOW and OUT of this section
                                    ' as it should only require results of newFmTest1 exchange above
                                End If
                            End With
                            ' These closing statements moved up from below following If block
                            '

                            'mtFamily = nmFamily 'to force "correct" behavior of following section
                            If mtFamily = "DSHEET" Then
                                Stop 'because we should NOT be doing Sheet Metal in this section.
                                ' This might require further investigation and/or development, if encountered.
                                'We should be okay. This is sheet metal stock
                                nmFamily = "D-RMT"
                                qtUnit = "FT2"
                                ' UPDATE[2018.05.30]:
                                '     Moving part family assignment
                                '     to this section for better mapping
                                '     and updating to new Family names
                                '     as well as pulling up qtUnit assignment
                            ElseIf mtFamily = "D-BAR" Then
                                ' UPDATE[2022.01.11]:
                                '     Adding Do..Loop Until to following section
                                '     to allow user to retry ting material
                                '     quantity and units. This change made in
                                '     conjunction with new prompt form (below).
                                ' NOTE! This is SECOND instance of revision
                                '     Search on UPDATE text above to locate
                                '     the other in this function
                                nmFamily = "R-RMT"
                                qtUnit = prRmUnit.Text '"IN"
                                ck = vbCancel
                                Do
                                    'Debug.Print(pnModel, " [", prRawMatl.Text, "]: ", aiPropsDesign(pnDesc).Text
                                    ' UPDATE[2021.03.11]: Replaced
                                    ' aiPropsDesign.Item(pnPartNum)
                                    ' as noted above
                                    'Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,"
                                    'Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                    'Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW."

                                    ' REV[2022.02.08.1521]
                                    ' replaced boilerplate above with new version below
                                    ' as per REV[2022.02.08.1511]

                                    Debug.Print("===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====")
                                    Debug.Print("  If change required, place new values at end")
                                    Debug.Print("  of lines below for prRmQty.Text and qtUnit.")
                                    Debug.Print("  Press [ENTER] on each line to be changed.")
                                    Debug.Print("  Press [F5] when ready to continue.")
                                    Debug.Print("----- " & pnModel & " [" & prRawMatl.Text & "]: " & aiPropsDesign.item(pnDesc).Text & " -----")
                                    'Debug.Print(""

                                    ' REV[2022.02.09.0919]
                                    ' prep to replace old dimension dump
                                    ' operation with more compact call
                                    ' to aiBoxData's Dump method
                                    If True Then 'go ahead and run old dump
                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                        ' REV[2022.02.09.0904]
                                        ' replicated With block from other section
                                        ' to replace original "sprawled out" version
                                        ' of Print statement hastily generated
                                        ' during run time.
                                        With invDoc.ComponentDefinition.RangeBox
                                            Debug.Print("")
                                            System.Math.Round((.MaxPoint.X - .MinPoint.X) / cvLenIn2cm, 4)
                                            System.Math.Round((.MaxPoint.Y - .MinPoint.Y) / cvLenIn2cm, 4)
                                            System.Math.Round((.MaxPoint.Z - .MinPoint.Z) / cvLenIn2cm, 4)
                                        End With
                                    End If

                                    With nuAiBoxData().UsingInches().UsingBox(
                                            invDoc.ComponentDefinition.RangeBox
                                        )
                                        Debug.Print(.Dump(0))
                                    End With
                                    'Stop 'and check output against prior version

                                    ' REV[2022.02.08.1446]
                                    ' removed block of Debug.Print(lines
                                    ' disabled now for some time, as they
                                    ' do not seem to have been missed.
                                    Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text), " 'in model. ")
                                    If dcIn.Exists(pnRmQty) Then Debug.Print("In Genius: ", CStr(dcIn.Item(pnRmQty)))
                                    Debug.Print("")
                                    Debug.Print("qtUnit = """, qtUnit, """ 'in model.")
                                    If dcIn.Exists(pnRmUnit) Then Debug.Print("In Genius: ", CStr(dcIn.Item(pnRmUnit)))
                                    Debug.Print(" ( or try IN )")

                                    ' REV[2022.02.08.1525]
                                    ' replaced boilerplate below with new version
                                    ' above in like manner to REV[2022.02.08.1446]
                                    ' and also per REV[2022.02.08.1511]

                                    'Debug.Print("qtUnit = ""IN"""
                                    'Debug.Print(""
                                    'Debug.Print(""
                                    'Debug.Print(""
                                    'Debug.Print(""
                                    'Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
                                    'Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE."
                                    'Debug.Print(""
                                    'Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text)
                                    'Debug.Print("qtUnit = ""IN"""
                                    Debug.Print("")
                                    Stop 'because we might want a D-BAR handler
                                    ' Actually, we might NOT need to stop here
                                    ' if bar stock is already selected,
                                    ' because quantities would presumably
                                    ' have been established already.
                                    ' Any D-BAR handler probably needs
                                    ' to be implemented in prior section(s)
                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                    ck = newFmTest2().AskAbout(invDoc,
                                            "Raw Material Quantity is now " _
                                            & CStr(prRmQty.Text) & qtUnit & " for",
                                            "If this is okay, click [YES]. Otherwise," _
                                            & vbCrLf & "click [NO] or [CANCEL] to fix."
                                        )
                                    'Stop
                                Loop Until ck = vbYes
                                ' UPDATE[2022.01.11]:
                                '     This is the terminal end of the
                                '     Do..Loop Until block noted above

                                rt = dcAddProp(prRmQty, rt)
                                Debug.Print("") 'Landing line for debugging. Do not disable.
                            Else
                                nmFamily = ""
                                qtUnit = "" 'may want function here
                                ' UPDATE[2018.05.30]: As noted above
                                '     However, might need more handling here.
                                Stop 'because we don't know WHAT to do with it
                            End If
                        Else
                            If 0 Then Stop 'and regroup
                            ' Things are looking a right royal mess
                            ' at the moment I'm writing this comment.
                        End If


                        ' NOTE[2022.01.07.1004]:
                        '     Another check for empty recommendation.
                        '     (SEE NOTE[2022.01.05] elsewhere in this function)
                        '     Again, don't want user accidentally
                        '     clearing an existing part number.
                        If Len(pnStock) > 0 Then
                            With prRawMatl
                                If Len(Trim$(.Text)) > 0 Then
                                    If pnStock <> .Text Then
                                        'Debug.Print("Raw Stock Selection"
                                        'Debug.Print("  Current : " & prRawMatl.Text
                                        'Debug.Print("  Proposed: " & pnStock
                                        'Stop 'because we might not want to change existing stock ting
                                        'if
                                        ck = MsgBox(
                                    Join({
                                        "Raw Stock Change Suggested",
                                        "  Current : " & prRawMatl.Text,
                                        "  Proposed: " & pnStock,
                                        "", "Change It?", ""
                                    }, vbCrLf),
                                    vbYesNo, "Change Raw Material?"
                                )
                                        '"Suggested Sheet Metal"
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            .Text = pnStock
                                        End If
                                    End If
                                Else
                                    .Text = pnStock
                                End If
                            End With
                        End If
                        rt = dcAddProp(prRawMatl, rt)

                        With prRmUnit
                            If Len(.Text) > 0 Then
                                If Len(qtUnit) > 0 Then
                                    If .Text <> qtUnit Then
                                        'Stop 'and check both so we DON'T
                                        'automatically "fix" the RMUNIT value

                                        ck = newFmTest2().AskAbout(invDoc, ,
                                        "Raw Material " & prRawMatl.Text _
                                        & vbCrLf & "Unit of Measure currently " _
                                        & .Text & vbCrLf & vbCrLf _
                                        & "Change to " & qtUnit & "?" _
                                        & vbCrLf & " "
                                    )
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            .Text = qtUnit
                                        End If
                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        '
                        '
                        '
                    End If 'Sheetmetal vs Part
                ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                    ' As mentioned above, nmFamily
                    ' SHOULD be  at this point
                    If Len(nmFamily) = 0 Then
                        If 1 Then Stop 'because we might
                        'need to check out the situation
                        nmFamily = "D-PTS" 'by default
                    End If
                ElseIf bomStruct = BOMStructureEnum.kPhantomBOMStructure Then
                    ' REV[2022.01.17.1135]
                    '     Adding a crude handler for Phantom
                    '     Part Documents. Since they shouldn't
                    '     have subcomponents to promote, they
                    '     shouldn't have that BOM structure.
                    '     User intervention might be required.
                    ck = newFmTest2().AskAbout(invDoc,
                    "For some reason, THIS Item is marked Phantom:",
                    "Is this okay? (Click [NO] OR [CANCEL] if not)"
                )
                    If ck = vbYes Then
                        'just let it go
                    Else
                        Stop
                    End If
                Else
                    ' REV[2022.01.17.1138]
                    '     Adding another handler to catch Part
                    '     Documents with an unexpected BOM Structure. Since they shouldn't
                    '     have subcomponents to promote, they
                    '     shouldn't have that BOM structure.
                    '     User intervention might be required.
                    ck = newFmTest2().AskAbout(invDoc,
                    "The following Item has an unhandled BOM Structure:",
                    "Skip it? (Click [NO] OR [CANCEL] to review)"
                )
                    If ck = vbYes Then
                        'just let it go
                    Else
                        Stop 'and let User decide what to do with it.
                        'NOTE to USER: See 'bomStruct' in the 'Locals'
                        'window ('Locals Window' under View menu)
                        'to see name of current BOM structure.
                    End If
                    Stop '(extraneous, disable/remove whenever)
                End If

                ' Get the design tracking property ,
                ' and update the Cost Center Property
                If invDoc.ComponentDefinition.IsContentMember Then
                    ' Don't muck around with the Family!
                Else
                    If Len(nmFamily) > 0 Then
                        On Error Resume Next
                        prFamily.Text = nmFamily
                        If Err.Number Then
                            Debug.Print("CHGFAIL[FAMILY]{'" _
                            & prFamily.Text & "' -> '" & nmFamily & "'}: " _
                            & invDoc.DisplayName & " (" & invDoc.FullDocumentName & ")")
                            If MsgBox(
                            "Couldn't Change Family" & vbCrLf _
                            & "for Item " & invDoc.DisplayName & vbCrLf _
                            & vbCrLf & "(" & invDoc.FullDocumentName & ")" _
                            & vbCrLf & vbCrLf & "Stop to Review?",
                            vbYesNo Or vbDefaultButton2,
                            invDoc.DisplayName
                        ) = vbYes Then
                                Stop
                            End If
                        Else
                        End If
                        On Error GoTo 0
                        rt = dcAddProp(prFamily, rt)
                        ' rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                    End If
                End If
            End With

            Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
            dcGeniusPropsPartRev20180530_ck = rt
        End If
    End Function
End Module