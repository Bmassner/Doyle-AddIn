Module dvl2
    Public Function dcGnsMatlSpecPairings(ThisApplication As Inventor.Application) As Scripting.Dictionary
        '
        ' dcGnsMatlSpecPairings -- Genius Raw Material Spec Relations
        '     Return a Dictionary of Dictionaries
        '     keyed to each Specification value
        '     found in ANY Spec field of any
        '     Raw Material Item, each listing
        '     all OTHER Spec values found in
        '     conjunction with each value.
        '
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim dcVl As Scripting.Dictionary
        Dim dcAl As Scripting.Dictionary
        Dim kyVl As Object
        Dim dxAl As Object
        Dim kyAl As String

        rt = New Scripting.Dictionary

        wk = dcDxFromRecSetDc(DcFromAdoRS(
    CnGnsDoyle().Execute(SqlOf_MatlSpecXref(ThisApplication))
))
        If wk Is Nothing Then
        ElseIf wk.Exists("val") Then
            dcVl = wk.Item("val")
            With dcVl
                For Each kyVl In .Keys
                    rt.Add(kyVl, New Scripting.Dictionary)
                Next

                For Each kyVl In .Keys
                    dcAl = rt.Item(kyVl)
                    With dcOb(.Item(kyVl))
                        For Each dxAl In .Keys
                            kyAl = dcOb(.Item(dxAl)).Item("also")
                            With rt
                                If .Exists(kyAl) Then
                                    dcAl.Add(kyAl, .Item(kyAl))
                                Else
                                    Stop 'because something went wrong
                                End If : End With
                        Next : End With
                Next
            End With
        Else
        End If

        dcGnsMatlSpecPairings = rt
    End Function

    Public Function dcOfDcWithXrefsDep1st(
    dc As Scripting.Dictionary,
    Optional wk As Scripting.Dictionary = Nothing,
    Optional PartDoc As String = "#"
) As Scripting.Dictionary
        '
        ' dcOfDcWithXrefsDep1st
        '     Replace rudundant / recursive Dictionary
        '     Objects in hierarchical Dictionary structure
        '     '
        '     This is a depth-first implementation, which
        '     might locate an initial Dictionary reference
        '     deep inside an early branch before finding a
        '     shallower instance that might be preferable.
        '     '
        '     A breadth-first implementation might be preferred.
        '
        Dim rt As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim ar As Object
        Dim ky As Object
        Dim sp As String

        If wk Is Nothing Then
            rt = dcOfDcWithXrefsDep1st(dc,
        New Scripting.Dictionary)
        Else
            rt = New Scripting.Dictionary
            With dc
                For Each ky In .Keys
                    ar = { .Item(ky)}
                    ck = dcOb(obOf(ar(0)))

                    If ck Is Nothing Then
                    Else
                        If wk.Exists(ck) Then
                            ar = {wk.Item(ck)}
                        Else
                            ' prep new $ref path
                            sp = PartDoc & "/" & CStr(ky)

                            ' add new $ref to wk
                            With NuDcPopulator(
                        ).Setting("$ref", sp)
                                wk.Add(ck, .Dictionary)
                            End With

                            ' go ahead and process
                            ' subdictionary
                            ar = {dcOfDcWithXrefsDep1st(ck, wk, sp)}
                            ' with new $ref in wk BEFORE the call,
                            ' it should be picked up for any new
                            ' references to the same directory
                        End If
                    End If

                    rt.Add(ky, ar(0))
                Next
            End With

            dcOfDcWithXrefsDep1st = rt
        End If

        dcOfDcWithXrefsDep1st = rt
        'send2clipBdWin10 ConvertToJson(dcOfDcWithXrefsDep1st(dcGnsMatlSpecPairings()), vbTab)
    End Function

    Public Function dcOfDcWithXrefsBrd1st(dc As Scripting.Dictionary,
    Optional wk As Scripting.Dictionary = Nothing,
    Optional PartDoc As String = "#"
) As Scripting.Dictionary
        '
        ' dcOfDcWithXrefsBrd1st
        '     Replace rudundant / recursive Dictionary
        '     Objects in hierarchical Dictionary structure
        '     '
        '     This is a depth-first implementation, which
        '     might locate an initial Dictionary reference
        '     deep inside an early branch before finding a
        '     shallower instance that might be preferable.
        '     '
        '     A breadth-first implementation might be preferred.
        '
        Dim rt As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim ls As Scripting.Dictionary
        Dim ar As Object
        Dim ky As Object
        Dim sp As String

        If wk Is Nothing Then
            rt = dcOfDcWithXrefsBrd1st(dc,
        New Scripting.Dictionary)
        Else
            ' create returned Dictionary
            rt = New Scripting.Dictionary

            ' create local working
            ' Dictionary of Dictionaries
            ls = New Scripting.Dictionary

            ' being processing
            ' supplied Dictionary
            With dc
                ' first pass: collect and process
                ' all sub Dictionary Objects
                For Each ky In .Keys
                    ck = dcOb(obOf(.Item(ky)))

                    If Not ck Is Nothing Then
                        If wk.Exists(ck) Then
                            ' add existing $ref Dictionary
                            ' to Dictionary list. thinking
                            ' recursion should NOT be an issue
                            ls.Add(ky, wk.Item(ck))
                        Else
                            ' add new Dictionary to list
                            ' for subsequent recursion
                            ls.Add(ky, .Item(ky))

                            ' prep new $ref path
                            sp = PartDoc & "/" & CStr(ky)

                            With wk
                                ' add new $ref Dictionary
                                .Add(ck, New Scripting.Dictionary)

                                ' add path to Dictionary
                                dcOb(.Item(ck)).Add("$ref", sp)
                            End With
                        End If
                    End If
                Next

                For Each ky In .Keys
                    If ls.Exists(ky) Then
                        rt.Add(ky, dcOfDcWithXrefsBrd1st(ls.Item(ky),
                        wk, PartDoc & "/" & CStr(ky)
                    ))
                    Else
                        rt.Add(ky, .Item(ky))
                    End If
                Next
            End With
        End If

        dcOfDcWithXrefsBrd1st = rt
        'send2clipBdWin10 ConvertToJson(dcOfDcWithXrefsBrd1st(dcGnsMatlSpecPairings()), vbTab)
    End Function

    Public Function dcGnsMatlSpecPairings4json(ThisApplication As Inventor.Application) As Scripting.Dictionary
        '
        ' dcGnsMatlSpecPairings4json -- check on dcGnsMatlSpecPairings
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim k2 As Object

        rt = dcGnsMatlSpecPairings(ThisApplication)

        With rt : For Each ky In .Keys()
                '.Item(ky) = Join(dcOb(.Item(ky)).Keys)
                With dcOb(.Item(ky))
                    For Each k2 In .Keys()
                        .Item(k2) = Join(dcOb(.Item(k2)).Keys)
                    Next : End With
            Next : End With

        dcGnsMatlSpecPairings4json = rt
        'send2clipBdWin10 ConvertToJson(dcGnsMatlSpecPairings4json(), vbTab)
    End Function

    Public Function dcSpecSubWith(
    txSpec As String,
    inDc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        If inDc.Exists(txSpec) Then
            dcSpecSubWith = DcKeysInCommon(inDc,
            dcOb(inDc.Item(txSpec)), 1
        )
        Else
            dcSpecSubWith = New Scripting.Dictionary
        End If
        'Debug.Print(Join(dcSpecSubWith("ROUND", dcSpecSubWith("BAR", dcGnsMatlSpecPairings())).Keys)
    End Function

    Public Function dcSpecSubWithAll(
    dcSpec As Scripting.Dictionary,
    inDc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = inDc
        For Each ky In dcSpec.Keys
            rt = dcSpecSubWith(CStr(ky), rt)
        Next
        dcSpecSubWithAll = rt
    End Function

    Public Function dcSpecFromUser(ThisApplication As Inventor.Application
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim dc As Scripting.Dictionary
        Dim Fm As FmSelectorList
        Dim nx As String

        rt = New Scripting.Dictionary
        dc = dcGnsMatlSpecPairings(ThisApplication)
        'Debug.Print(Join(dc.Keys)

        Do
            Fm = nuSelFromDict(dc)
            nx = Fm.GetReply(, "")

            If Len(nx) > 0 Then
                rt.Add(nx, nx)
                dc = dcSpecSubWith(nx, dc)
                If dc.Count = 0 Then nx = ""
            End If
        Loop While Len(nx) > 0

        'Stop
        dcSpecFromUser = rt
        'Debug.Print(Join(dcSpecFromUser().Keys)
    End Function

    Public Function d2g3f1(
    Part As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' d2g3f1 -- Return Dictionary
        '     of relevant Part Properties
        '     and information for use in
        '     Genius data extraction
        '
        Dim rt As Scripting.Dictionary
        Dim txPartNum As String

        If dc Is Nothing Then
            rt = d2g3f1(Part,
        New Scripting.Dictionary)
        Else
            rt = dc

            With Part
                With .Propertys.Item(GnDesign)
                    rt.Add(PnPartNum, .Item(PnPartNum))
                    rt.Add(PnFamily, .Item(PnFamily))
                End With

                'rt.Add("subType", .SubType 'aiSubType

                'With .ComponentDefinition
                '    rt.Add("bomStr", .BOMStructure 'aiBomType
                '
                '    With nuAiBoxData().SortingDims(.RangeBox)
                '        With .UsingInches()
                '            rt.Add("Width", .SpanX
                '            rt.Add("Length", .SpanY
                '            rt.Add("Height", .SpanZ
                '        End With
                '    End With
                'End With
            End With

            rt = dcGnsInfoCompDef(aiCompDefOf(Part), rt)
            'aiCompDefOf replaces obAiCompDefAny
        End If

        d2g3f1 = rt
    End Function

    Public Function dcGnsInfoAiDocBase(
    AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoAiDocBase (formerly d2g3f1a)
        '     Return Dictionary of Document Properties
        '     and information relevant to Genius
        '     for data extraction
        '
        Dim rt As Scripting.Dictionary
        Dim txPartNum As String

        If dc Is Nothing Then
            rt = dcGnsInfoAiDocBase(AiDoc,
        New Scripting.Dictionary)
        Else
            rt = dc

            With AiDoc
                With .Propertys.Item(GnDesign)
                    rt.Add(PnPartNum, .Item(PnPartNum))
                    rt.Add(PnFamily, .Item(PnFamily))
                End With

                If False Then
                    rt.Add("subType", .SubType) 'aiSubType
                    rt.Add("docType", .DocumentType)
                    rt.Add("dsbType", .DocumentSubType.DocumentSubTypeID)
                End If
            End With

            rt = dcGnsInfoCompDef(aiCompDefOf(AiDoc), rt)
            'aiCompDefOf replaces obAiCompDefAny
        End If

        dcGnsInfoAiDocBase = rt
    End Function

    Public Function dcGnsInfoCompDef(
    CpDef As Inventor.ComponentDefinition,
    Optional dcWkg As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoCompDef -- Generate and/or populate
        '     Dictionary (new or supplied) with data for
        '     Genius from supplied ComponentDefinition.
        ' This is the "generic" Object, which dispatches
        '     the supplied ComponentDefinition to a Object
        '     more specific to its Class. (some Class
        '     Objects remain to be implemented)
        ' Note that this function follows the convention
        '     of a recursive call with a new Dictionary
        '     object when none is supplied. Duplication
        '     of the basic function structure should ensure
        '     this pattern is followed by all specialized
        '     Objects. While this should not usually be
        '     necessary under normal usage (dispatch to
        '     specialized Objects from here), it should
        '     help accommodate the possibility of direct
        ''      calls from other client functions.
        '
        Dim rt As Scripting.Dictionary

        rt = dcWkg
        If rt Is Nothing Then
            rt = dcGnsInfoCompDef(CpDef,
        New Scripting.Dictionary)
        ElseIf CpDef Is Nothing Then 'Do Nothing
            'cuz we got Nothing to Do With Nothing
        Else
            With CpDef '.ComponentDefinition
                rt.Add("bomStr", .BOMStructure) 'aiBomType
                If .BOMStructure = BOMStructureEnum.kNormalBOMStructure Then rt.Add("Type", "M")
                If .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then rt.Add("Type", "R")

                With nuAiBoxData().UsingBox(.RangeBox) '.SortingDims
                    With .UsingInches() 'WARNING[2021.12.15]
                        ' Forcing inch conversion MAY lead
                        ' to issues in future development.
                        ' It is absolutely ESSENTIAL that
                        ' unit measurement be tracked and
                        ' kept consistent throughout the
                        ' entire management process.
                        rt.Add(PnLength, System.Math.Round(.SpanX, 6))
                        rt.Add(PnWidth, System.Math.Round(.SpanY, 6))
                        rt.Add("Height", System.Math.Round(.SpanZ, 6))
                    End With
                End With
            End With

            If TypeOf CpDef Is Inventor.SheetMetalComponentDefinition Then
                rt = dcGnsInfoCompDefShtMtl(CpDef, rt)
            ElseIf TypeOf CpDef Is Inventor.WeldmentComponentDefinition Then
                Stop 'using general Assembly handler
                rt = dcGnsInfoCompDefAssy(CpDef, rt)
            ElseIf TypeOf CpDef Is Inventor.WeldsComponentDefinition Then
                Stop 'using general Assembly handler
                rt = dcGnsInfoCompDefAssy(CpDef, rt)
            ElseIf TypeOf CpDef Is Inventor.PartComponentDefinition Then
                rt = dcGnsInfoCompDefPart(CpDef, rt)
            ElseIf TypeOf CpDef Is Inventor.AssemblyComponentDefinition Then
                rt = dcGnsInfoCompDefAssy(CpDef, rt)
            Else
            End If
        End If

        dcGnsInfoCompDef = rt
    End Function

    Public Function dcGnsInfoCompDefShtMtl(
    CpDef As Inventor.SheetMetalComponentDefinition,
    Optional dcWkg As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoCompDefShtMtl -- Generate and/or populate Dictionary
        '     (new or supplied) with data for Genius
        '     from supplied ComponentDefinition.
        ' This is the Assembly Object.
        '     '
        '
        Dim rt As Scripting.Dictionary
        Dim wd As Double 'width
        Dim lg As Double 'length
        Dim ht As Double 'height
        Dim tk As Double 'thickness
        Dim ar As Double 'area
        Dim ck As Double 'check height vs thickness

        Dim rm As Scripting.Dictionary
        Dim s6 As String
        Dim ky As Object

        rt = dcWkg
        If rt Is Nothing Then
            rt = dcGnsInfoCompDefShtMtl(
        CpDef, New Scripting.Dictionary)
        Else
            rt = dcGnsInfoCompDefPart(CpDef, rt)
            With rt
                If Not .Exists("SPEC06") Then
                    Stop
                    .Add("SPEC06", steelSpec6("", 1))
                End If
                s6 = .Item("SPEC06")

                If Not .Exists("RMLIST") Then
                    .Add("RMLIST", New Scripting.Dictionary)
                End If
                rm = dcOb(.Item("RMLIST"))
            End With

            With CpDef
                tk = .Thickness.Text / CvLenIn2cm
                ' NOTE conversion to Inches from Centimeters.
                ' keep in mind we're grabbing Thickness HERE
                ' and will use Height (below) in an effort
                ' to validate the Flat Pattern, and determine
                ' this Part is MEANT to be Sheet Metal.

                If .HasFlatPattern Then
                    With .FlatPattern
                        With nuAiBoxData().UsingBox(.RangeBox)
                            With .UsingInches()
                                ht = System.Math.Round(.SpanZ, 6)
                                ' remember, Height here is meant
                                ' to verify Sheet Metal Part

                                lg = System.Math.Round(.SpanX, 6)
                                wd = System.Math.Round(.SpanY, 6)
                            End With
                        End With
                    End With

                    ck = System.Math.Round(Math.Abs(ht - tk), 6)
                    If ck > 0.002 Then
                        With dcFlatPatSpansByVertices(.FlatPattern)
                            If ht > .Item("Z") Then
                                ht = .Item("Z")

                                Debug.Print(.Item("X") - lg)
                                Debug.Print(.Item("Y") - wd)
                                Debug.Print("") 'Breakpoint Landing
                            End If
                        End With
                    Else
                    End If

                    If System.Math.Round(Math.Abs(ht - tk), 6) > 0 Then
                    End If

                    ar = lg * wd '.SpanX * .SpanY
                    'does this need to be divided by 144?
                    'to get to ft^2? or do we stick to in^2?
                Else
                    With rt
                        If .Exists(PnLength) Then
                            lg = .Item(PnLength)
                            '.Remove(pnLength
                        End If

                        If .Exists(PnWidth) Then
                            wd = .Item(PnWidth)
                            '.Remove(pnWidth
                        End If

                        If .Exists("Height") Then
                            ht = .Item("Height")
                            '.Remove("Height"
                        End If
                    End With

                    ar = 0 'STOPGAP[2021.12.08]
                    ' might want to consider using this
                    ' to store whatever material quantity
                    ' might be obtained, regardless
                    ' of stock type
                End If
                '
                ' At this point, we should have either
                ' likely dimensions of the flat pattern, OR
                ' the original dimensions of the part itself.
                '
                ' The next step is to determine whether they
                ' are consistent with a valid sheet metal part.
                ' If not, it's likely a structural one.
                '
                ' The key criterion is how closely the height
                ' dimension matches the given thickness.
                '
                ck = System.Math.Round(Math.Abs(ht - tk), 6)
                If ck > 0.002 Then
                Else
                End If

                ' REV[2021.12.15]:
                '     add material option collection
                '     specific to sheet metal
                With dcGnsMatlOps(dcCtOfEach(
                {tk, lg, wd, ht}
            ), s6)
                    For Each ky In .Keys
                        If Not rm.Exists(ky) Then
                            rm.Add(ky, .Item(ky))
                        End If : Next
                End With

                '
                '
                '
                With rt
                    ' first, remove any previous
                    ' dimensional values
                    If .Exists(PnLength) Then .Remove(PnLength)
                    If .Exists(PnWidth) Then .Remove(PnWidth)
                    If .Exists("Height") Then .Remove("Height")
                    '(not sure this is the best way
                    ' but going to try it for now)

                    .Add(PnThickness, tk)
                    .Add(PnLength, lg)
                    .Add(PnWidth, wd)
                    .Add(PnArea, ar)
                    .Add("Height", ht)
                End With
            End With
        End If

        dcGnsInfoCompDefShtMtl = rt
    End Function

    Public Function dcFlatPatSpansByVertices(
    smFlat As Inventor.FlatPattern
) As Scripting.Dictionary
        '
        ' dcFlatPatSpansByVertices -- get extents of
        '     Sheet Metal Flat Pattern
        '     from a scan of its Vertices.
        '     this is a last resort,
        '     in case an erroneous Z span
        '     reported from the Range Box
        '     fails to match Thickness.
        '
        Dim rt As Scripting.Dictionary
        Dim vx As Inventor.Vertex
        Dim xmn As Double
        Dim xmx As Double
        Dim ymn As Double
        Dim ymx As Double
        Dim zmn As Double
        Dim zmx As Double

        rt = New Scripting.Dictionary

        If Not smFlat.Body Is Nothing Then
            With smFlat.Body '.Vertices'.RangeBox
                For Each vx In .Vertices
                    With vx.Point
                        If .X < xmn Then xmn = .X
                        If .X > xmx Then xmx = .X
                        If .Y < ymn Then ymn = .Y
                        If .Y > ymx Then ymx = .Y
                        If .Z < zmn Then zmn = .Z
                        If .Z > zmx Then zmx = .Z
                    End With
                Next
            End With
            '

            '
        End If

        With rt
            .Add("X", xmx - xmn)
            .Add("Y", ymx - ymn)
            .Add("Z", zmx - zmn)
        End With

        dcFlatPatSpansByVertices = rt
    End Function

    Public Function dcGnsInfoCompDefPart(ThisApplication As Inventor.Application,
    CpDef As Inventor.PartComponentDefinition,
    Optional dcWkg As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoCompDefPart -- Generate and/or populate Dictionary
        '     (new or supplied) with data for Genius
        '     from supplied ComponentDefinition.
        ' This is the general Part Object.
        '     It's grown somewhat in complexity
        '     since development's begun.
        ' Here is rough flow map:
        '     (inside component definition)
        '     - stage supplied Dictionary for return
        '       (start a new one, if none supplied
        '        one USUALLY should be)
        '     - get Mass -- don't add to Dictionary yet
        '       (no data should be added to Dictionary
        '        until all data are collected and verified)
        '     - get Active Material, and its name
        '     - use this to  target Spec 6
        '     (inside returning Dictionary)
        '     - collect length, width, and height
        '       dimensions from Dictionary
        '       (this is why one should be supplied)
        '     - collect raw material candidate items
        '       from Genius into a Record, using
        '       an SQL query generated from collected
        '       dimensions, and material Spec 6
        '     - generate Dictionary of candidates
        '       from the Record, keyed on item names
        '     - add data to Dictionary:
        '       - mass
        '       - material name
        '       - spec 6
        '       - Dictionary of raw material
        '         item candidates
        '
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary 'may be temporary
        Dim d2 As Scripting.Dictionary 'may be temporary
        Dim mt As Inventor.MaterialAsset
        Dim mtName As String
        Dim s6 As String
        Dim ms As Double 'mass
        Dim ky As Object
        Dim ck As Double

        rt = dcWkg
        If rt Is Nothing Then
            rt = dcGnsInfoCompDefPart(CpDef,
        New Scripting.Dictionary)
        Else
            With CpDef
                With .MassProperties
                    ms = System.Math.Round(.Mass * CvMassKg2LbM, 4)
                    'System.Math.Round( _
                    ThisApplication.kUnitsOfMeasureObject.ConvertUnits.Mass(, "kg", "lb", 4)
                    ' Apparently, empty parentheses may be
                    ' placed after both ThisApplication
                    ' and UnitsOfMeasure without error.
                    ' This makes it rather easy to lay out
                    ' code in a nice, compact, and maybe
                    ' even more readable form.
                End With

                'ptNumShtMetal
                mt = aiDocPart(.Document).ActiveMaterial
                If mt Is Nothing Then
                    mtName = ""
                    'Stop
                Else
                    mtName = mt.DisplayName
                    ' NOTE[2021.12.03]:
                    ' mt also has a .CategoryName that might
                    ' want included in full material designator.
                    ' also not sure if constant pnMaterial is
                    ' best choice for Dictionary Key, though
                    ' probably so. Keep this point in mind.
                End If
                s6 = steelSpec6(mtName)
            End With

            With rt
                wk = New Scripting.Dictionary
                For Each ky In {PnLength, PnWidth, "Height"}
                    If .Exists(ky) Then
                        wk.Add(ky, System.Math.Round(CDbl(.Item(ky)), 6))
                        '
                        ' REV[2021.12.15]:
                        '     disabling everything from here
                        '     to end of Then block. Instead,
                        '     will collect target dimension
                        '     values, then submit them to
                        '     function dcCtOfEach to generate
                        '     the "histogram".
                        '
                        '     that way, the same function may
                        '     be used by other callers, like
                        '     dcGnsInfoCompDefShtMtl
                        '
                        '    'ck = System.Math.Round(CDbl(.Item(ky)), 6)
                        '    ' NOTE[2021.12.08]:
                        '    '     The conversion kludge here
                        '    '     might NOT be reliable for
                        '    '     long term use. Be prepared
                        '    '     to deal with issues here.
                        '    With wk
                        '    If .Exists(ck) Then
                        '        .Item(ck) = _
                        '        .Item(ck) + 1
                        '    Else
                        '        .Add(ck, 1 'ky
                        '    End If: End With
                        ' NOTE[2021.12.08]:
                        '     Dictionary wk counts occurrences
                        '     of each dimension value. While not
                        '     presently used, the count might
                        '     prove helpful in prioritizing raw
                        '     material candidate items.
                        ' REV[2021.12.15]:
                        '     occurrence count has been moved
                        '     below, outside the For-Next loop,
                        '     and into a call to new function
                        '     dcCtOfEach.
                    End If
                Next
                wk = dcCtOfEach(wk.Items)
                If wk.Count = 0 Then wk.Add(0.075, 1)
                ' another kludge to trap an error
                ' which should NOT occur as long as
                ' a prepared Dictionary is supplied.

                '
                ' Here is where we'll attempt to collect
                ' raw material Item candidates from Genius
                '

                'present up -- plan to change
                '"select d.v "
                'Debug.Print("from (values (" & txDumpLs(wk.Keys, "), (") & "))"
                '" as d(v)"

                'future proposal -- counts occurrences
                '"select d.v, d.c "
                'Debug.Print("from (values (" & dumpLsKeyVal(wk, ", ", "), (") & "))"
                '" as d(v, c)"

                wk = dcGnsMatlOps(wk, s6)
                ' REV[2021.12.15]:
                '     preceding line replaces With block below,
                '     moving Genius material options request
                '     to function dcGnsMatlOps, so it can be
                '     called from other functions, like,
                '     again, dcGnsInfoCompDefShtMtl
                'With cnGnsDoyle()
                '    Dim rs As ADODB.Record
                '
                '    'wk.RemoveAll
                '    On Error Resume Next
                '
                '    Err.Clear
                '     rs = .Execute( _
                '    sqlOf_GnsMatlOptions( _
                '        s6, wk.Keys _
                '    ))
                '
                '    If Err.Number = 0 Then
                '        With dcFromAdoRS(rs, "") ' wk =
                '        For Each ky In .Keys
                '             d2 = dcOb(.Item(ky))
                '            If d2 Is Nothing Then
                '                Stop
                '            Else
                '                wk.Add(d2.Item("Item"), d2
                '            End If
                '            'Stop
                '            ' ENDOFDAY[2021.12.08]:
                '            '     Need to up process of remapping
                '            '     raw material Items from Genius
                '            '     to their Item names
                '        Next: End With
                '
                '        rs.Close
                '    Else
                '        Stop
                '        Err.Clear
                '    End If
                '    On Error GoTo 0
                '
                '    .Close
                'End With

                '.Add(pnRawMaterial, wk

                .Add(PnMass, ms)
                .Add(PnMaterial, mtName)
                .Add("SPEC06", s6)

                'If False Then
                'not quite ready for this one yet
                .Add("RMLIST", wk)
                'End If
                Debug.Print("") 'Breakpoint Landing
            End With
        End If

        dcGnsInfoCompDefPart = rt
    End Function

    Public Function dcGnsInfoCompDefAssy(
    CpDef As Inventor.AssemblyComponentDefinition,
    Optional dcWkg As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoCompDefAssy -- Generate and/or populate Dictionary
        '     (new or supplied) with data for Genius
        '     from supplied ComponentDefinition.
        ' This is the general Assembly Object.
        '     '
        '
        Dim rt As Scripting.Dictionary

        rt = dcWkg
        If rt Is Nothing Then
            rt = dcGnsInfoCompDefAssy(CpDef,
        New Scripting.Dictionary)
        Else
            With CpDef
                With .MassProperties
                    rt.Add(PnMass, System.Math.Round(
                .Mass * CvMassKg2LbM, 4))
                    '
                    ' see dcGnsInfoCompDefPart
                    ' for alternate
                    ' implementation
                    '
                End With
            End With
        End If

        dcGnsInfoCompDefAssy = rt
    End Function

    Public Function dcGnsInfoCompDefTBD(
    CpDef As Inventor.ComponentDefinition,
    Optional dcWkg As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGnsInfoCompDefTBD (formerly d2g4f2zz)
        '     Generate and/or populate Dictionary
        '     (new or supplied) with data for Genius
        '     from supplied ComponentDefinition.
        ' This is the <TBD> Object. (formerly <zz>)
        '     Use it as a template for others.
        '     (be sure to modify comments accordingly)
        '
        Dim rt As Scripting.Dictionary

        rt = dcWkg
        If rt Is Nothing Then
            rt = dcGnsInfoCompDefTBD(CpDef,
        New Scripting.Dictionary)
        Else
            With CpDef
            End With
        End If

        dcGnsInfoCompDefTBD = rt
    End Function

    Public Function dcGnsInfoSQLitem(ThisApplication As Inventor.Application,
    Item As String
) As Scripting.Dictionary
        '
        ' dcGnsInfoSQLitem -- Return a Dictionary
        '     of Part data from Genius
        '     for the indicated Item
        '
        Dim rt As Scripting.Dictionary
        Dim mt As Scripting.Dictionary
        Dim rs As ADODB.Record
        Dim ky As Object

        With CnGnsDoyle()
            On Error Resume Next
            Err.Clear()
            rs = .Execute(SqlOf_GnsPartInfo(ThisApplication, Item)) 'sqlOf_ASDF
            If Err.Number = 0 Then
                rt = DcFromAdoRSrow(rs, "")
                'With rs
                '    If Not .EOF Then
                '        .MoveNext
                '        If Not .EOF Then
                '            Stop 'to handle multiple raw materials
                '            Debug.Print("") 'Breakpoint Landing
                '        End If
                '    End If
                '
                '    .Close
                'End With
            Else
                Debug.Print(Err.Number)
                Debug.Print(Err.Description)
                Stop
            End If

            rs = .Execute(SqlOf_GnsPartMatl(ThisApplication, Item)) 'sqlOf_ASDF
            If Err.Number = 0 Then
                mt = DcFromAdoRS(rs, "")
                With mt
                    If .Count > 0 Then
                        If .Count > 1 Then
                            Stop 'to handle multiple raw materials
                            Debug.Print("") 'Breakpoint Landing
                        Else
                            With dcOb(.Item(.Keys(0)))
                                For Each ky In .Keys
                                    If rt.Exists(ky) Then
                                        Stop 'to deal with collision
                                        'which should NOT happen here
                                        'because field s returned
                                        'by each query should have
                                        'NO names in common.
                                    Else
                                        rt.Add(ky, .Item(ky))
                                    End If
                                Next : End With
                        End If
                    End If
                End With
            Else
                Debug.Print(Err.Number)
                Debug.Print(Err.Description)
                Stop
            End If

            On Error GoTo 0
            .Close()
        End With

        dcGnsInfoSQLitem = rt
        'Debug.Print(dumpLsKeyVal(dcGnsInfoSQLitem(aiProperty(d2g3f1(aiDocPart(userChoiceFromDc())).Item(pnPartNum)).Text), " = ")
    End Function

    Public Function d2g3f4(Part As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' d2g3f4 -- Return a Dictionary
        '     of Properties and info from
        '     Inventor Part Document for
        '     Genius Interface.
        '
        Dim rt As Scripting.Dictionary
        'Dim rs As ADODB.Record

        rt = dcProps4genius(Part, d2g3f1(Part, dc), 0)

        d2g3f4 = rt
        'Debug.Print(dumpLsKeyVal(d2g3f4(aiDocPart(userChoiceFromDc())), " = ")
    End Function

    Public Function d2g3f5(ThisApplication As Inventor.Application,
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        '
        ' d2g3f5 -- Gather Dictionaries of Inventor
        '     Properties and Genius info from supplied
        '     Document for correlation and potential
        '     revision.
        ' REV[2021.12.15]:
        '     Parameter Part renamed to AiDoc, with Class
        '     changed from PartDocument to the more general
        '     Document, as it would appear all supporting
        '     functions will accept and work with it.
        '
        Dim dcPt As Scripting.Dictionary
        '   base Genius info and inherent Properties
        Dim dcPr As Scripting.Dictionary
        '   base + custom Genius Properties
        Dim dcVlAi As Scripting.Dictionary
        '   values of all collected Properties
        Dim dcVlPr As Scripting.Dictionary
        '   values of all collected Properties
        Dim dcGn As Scripting.Dictionary
        '   information from Genius database
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        dcPt = dcGnsInfoAiDocBase(AiDoc)
        dcVlAi = dcMapAiProps2vals(dcPt)
        ' REV[2021.12.16]:
        '     additional value Dictionary
        '     collects values ONLY from
        '     the inherent document data
        '     and Properties, which gets
        '     overridden at the next step:

        dcPr = dcProps4genius(AiDoc, DcCopy(dcPt), 2)
        ' REV[2021.12.15] argument 2 replaces 0 in order
        ' to generate references to missing Properties
        ' (see dcGnsPropsListed) without trying to
        ' create them. That way, client functions may
        ' be made aware of Properties that need created.
        '
        ' Modifications to those functions might be needed
        ' to be prepared for missing Properties, whose
        ' names will map to Nothing (void references)
        ' UPDATE[2021.12.16]:
        '     just happened today. created new function
        '     blankIfNoValElseSelf to address this issue.
        '     see dcMapAiProps2vals for application
        '
        '= dcProps4genius(AiDoc, d2g3f1(AiDoc, dc), 0)
        '= d2g3f4(AiDoc)

        dcVlPr = dcMapAiProps2vals(dcPr) 'dcPt
        With dcVlPr
            For Each ky In {
            PnThickness, PnWidth, PnLength, PnArea, PnRmQty}
                ' THIS IS A KLUDGE
                ' to temporarily "fix" an issue with Width,
                ' Length, and Area values that don't match
                ' up between Inventor Properties and Genius,
                ' even when their numeric values are equal.
                '
                ' A more thorough review and revision
                ' will probably be needed eventually.
                '
                ' REV[2021.12.07]:
                ' Thickness is also affected
                ' and has been added to the list.
                '
                ' REV[2021.12.16]:
                '     added pnRmQty to list as quick and
                '     dirty method to force a blank value
                '     to zero, and prevent an error in
                '     the correction code after the loop.
                '
                '     This must be the sort of cruft Joel
                '     Spolsky was talking about in that
                '     essay of his. Still not a justified
                '     defense for crud programming, which
                '     this is, let's face it! Right there
                '     with you, Brando!
                '
                If .Exists(ky) Then
                    .Item(ky) = Val(Split(
                    "0" + CStr(.Item(ky)), " "
                )(0))
                    ' REV[2021.12.16]:
                    '     minor switch in order of operations:
                    '     now prepending "0" to string BEFORE
                    '     splitting, to avoid the emtpy array
                    '     returned by splitting an empty string
                End If
            Next

            If .Exists(PnRmQty) Then
                .Item(PnRmQty) = System.Math.Round(.Item(PnRmQty), 8)
                ' THIS is intended to fix a precision discrepancy
                ' between Inventor, which seems to store material
                ' quantity with at least twelve digits of precision,
                ' and Genius, which keeps only eight.
            End If
        End With

        dcGn = dcGnsInfoSQLitem(ThisApplication,
        dcVlPr.Item(PnPartNum)
    )

        rt = New Scripting.Dictionary
        With rt
            .Add("aiVal", dcVlAi)
            .Add("inv", dcVlPr)
            .Add("gns", dcGn)
            .Add("prp", dcPr)

            d2g3f5 = rt
            'send2clipBdWin10 ConvertToJson(d2g3f5(aiDocPart(userChoiceFromDc())), vbTab)
        End With
    End Function

    Public Function d2g3f5as(
    Assy As Inventor.AssemblyDocument,
    Optional ThisToo As Long = 0
) As Scripting.Dictionary
        '
        ' d2g3f5as -- Assembly counterpart to d2g3f5
        '     not sure what's actually to be done with it yet.
        '     probably just remove it, d2g3f5 can handle both.
        '
        Dim dc As Scripting.Dictionary

        dc = dcRemapByPtNum(
        dcAiDocComponents(
            Assy, , ThisToo
        )
    )
    End Function

    Public Function dcMapAiProps2vals(
    dc As Scripting.Dictionary,
    Optional Flags As Long = 0
) As Object
        '
        ' dcMapAiProps2vals --
        '     Return a Dictionary
        '     containing the Values of
        '     any Inventor Properties
        '     in supplied Dictionary,
        '     with all other members
        '     returned as they are.
        '
        ' related functions:
        '     dcOfDcAiPropVals
        '     dcAiPropValsFromDc
        '     dcOfPropsInAiDoc
        '
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        With DcNewIfNone(dc) : For Each ky In .Keys
                rt.Add(ky, blankIfNoValElseSelf(
        valIfAiPropElseSelf(.Item(ky))))
                ' REV[2021.12.16]:
                '     add call to blankIfNoValElseSelf
                '     against RESULT of valIfAiPropElseSelf
                '     so blind checks against expected key
                '     values don't fail when a null Object
                '     (AKA Nothing) is encountered.
            Next : End With
        dcMapAiProps2vals = rt
    End Function

    Public Function valIfAiPropElseSelf(
    vl As Object
) As Object
        '
        ' valIfAiPropElseSelf --
        '     Return the Value of any
        '     supplied Inventor Property.
        '     Any other type of argument
        '     should be returned directly.
        '
        Dim InvProperty As Inventor.Property

        If TypeOf vl Is Object Then
            InvProperty = aiProperty(obOf(vl))
            If InvProperty Is Nothing Then
                valIfAiPropElseSelf = vl
            Else
                valIfAiPropElseSelf = InvProperty.Text
            End If
        Else
            valIfAiPropElseSelf = vl
        End If
    End Function

    Public Function blankIfNoValElseSelf(
    vl As Object
) As Object
        '
        ' blankIfNoValElseSelf --
        '     Return the Value of any
        '     supplied Inventor InvProperty.
        '     Any other type of argument
        '     should be returned directly.
        '
        Dim InvProperty As Inventor.Property

        If TypeOf vl Is Inventor.Property Then
            If obOf(vl) Is Nothing Then
                blankIfNoValElseSelf = ""
            Else
                blankIfNoValElseSelf = vl
            End If
        ElseIf vl Is Nothing Then
            blankIfNoValElseSelf = ""
        Else
            blankIfNoValElseSelf = vl
        End If
    End Function

    Public Function d2g3f7(ThisApplication As Inventor.Application,
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        '
        ' d2g3f7 --
        '
        Dim rt As Scripting.Dictionary
        Dim Ky As Object

        ' rt = New Scripting.Dictionary
        With d2g3f5(ThisApplication, AiDoc)
            Debug.Print("") 'Breakpoint Landing
            rt =
               dcTreeReKeyedInPlc("src1", "gns",
               dcTreeReKeyedInPlc("src0", "inv",
               DcWBQbyCmpResult(
               DcCmpTextOf2dc(
                   .Item("inv"),
                   .Item("gns")
               )
           )))

            With rt
                'Stop
            End With
            rt.Add("prp", .Item("prp"))
            rt.Add("doc", AiDoc)
            ' might want this in order
            ' to grab custom Property

            'With dcKeysInCommon(.Item("inv"), .Item("gns"))
            'End With
        End With
        d2g3f7 = rt
        'send2clipBdWin10 ConvertToJson(d2g3f7(aiDocPart(userChoiceFromDc())), vbTab)
        'send2clipBdWin10 ConvertToJson(nuDcPopulator().Setting(Format$(Now, "\[YYYY.MM.DD@HH.NN.SS\]"), d2g3f7(aiDocPart(userChoiceFromDc()))).Dictionary(), vbTab)
    End Function

    Public Function d2g3f8(ThisApplication As Inventor.Application,
    Optional AiDoc As Inventor.Document = Nothing
) As Scripting.Dictionary
        '
        ' d2g3f8 --
        '
        Dim rt As Scripting.Dictionary
        'Dim ck As Inventor.Document
        Dim ky As Object

        If AiDoc Is Nothing Then
            With ThisApplication
                If .ActiveDocument Is Nothing Then
                    Stop
                Else
                    rt = d2g3f8(.ActiveDocument)
                End If
            End With
        Else
            rt = New Scripting.Dictionary

            With nuPicker(New KyPickAiPartVsAssy
        ).AfterScanning(dcAiDocComponents(AiDoc))
                '
                With .DcIn() 'Parts
                    For Each ky In .Keys
                        rt.Add(ky, d2g3f7(ThisApplication, aiDocPart(obOf(.Item(ky)))))
                    Next
                End With

                With .DcOut() 'Assemblies
                End With
            End With
        End If

        d2g3f8 = rt
        'send2clipBdWin10 ConvertToJson(nuDcPopulator().Setting(Format$(Now, "\[YYYY.MM.DD@HH.NN.SS\]"), d2g3f8(aiDocument(obOf(userChoiceFromDc())))).Dictionary(), vbTab)
    End Function

    Public Function dcTreeMembersWithKey(
    tg As Object, dc As Scripting.Dictionary,
    Optional wk As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcTreeMembersWithKey (formerly d2g5f1)
        '     Given a Dictionary that might contain
        '     other Dictionaries, check it and any
        '     sub Dictionaries for target key (tg)
        '     and return a Dictionary of those
        '     Dictionaries containing it, each
        '     keyed to the number already found.
        '     This should ensure a unique key
        '     for each match found, with no
        '     need to track any other keys.
        '
        ' The ultimate goal of this function is to
        '     support a Key Find/Replace operation
        '     across a hierarchy of Dictionaries.
        '
        ' This is initially and specifically to map
        '     comparison keys "src0" and "src1" to
        '     the names of sources they represent.
        '
        ' This is of course the 'Find' component
        '     of the ultimate product
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If wk Is Nothing Then
            rt = dcTreeMembersWithKey(tg, dc,
        New Scripting.Dictionary)
        Else
            rt = wk
            If Not dc Is Nothing Then
                With dc
                    If .Exists(tg) Then
                        With rt
                            .Add(.Count, dc)
                        End With
                    End If

                    For Each ky In .Keys
                        rt = dcTreeMembersWithKey(tg,
                    dcOb(obOf(.Item(ky))) _
                    , rt)
                    Next
                End With
            End If
        End If
        dcTreeMembersWithKey = rt
    End Function

    Public Function dcTreeMemWithReplcmt(
    rp As Object, dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcTreeMemWithReplcmt (formerly d2g5f2)
        '     Given a Dictionary of Dictionaries,
        '     check for any Dictionary containing target
        '     replacement Key rp, and return a Dictionary
        '     containing any results.
        '
        ' This is a check for potential Key collisions.
        '     The Dictionary returned should be empty.
        '
        ' This is presently accomplished by first calling
        '     dcTreeMembersWithKey against the supplied Dictionary,
        '     which is normally expected to be the result
        '     of a PRIOR call to dcTreeMembersWithKey using the target
        '     key to be replaced.
        '
        ' It is therefore possible that the supplied
        '     Dictionary might contain replacement key rp,
        '     and thus be included in the local result.
        '     That Dictionary should NOT be included
        '     in the FINAL result returned.
        '
        ' It is therefore necessary to scan the result
        '     of the local dcTreeMembersWithKey call, and remove it,
        '     if found.
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = dcTreeMembersWithKey(rp, dc)
        With rt : For Each ky In .Keys
                If dcOb(obOf(.Item(ky))) Is dc Then
                    Stop
                    .Remove(ky)
                End If
            Next : End With
        dcTreeMemWithReplcmt = rt
    End Function

    Public Function dcTreeReKeyedInPlc(
    tg As Object, rp As Object,
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcTreeReKeyedInPlc (formerly d2g5f3)
        '     Given a target Key tg (to be replaced),
        '     a replacement Key rp, and a Dictionary
        '     that includes other Dictionaries,
        '     attempt to replace all instances of the
        '     target Key with the replacement Key in
        '     all Dictionaries within the hierarchy.
        '
        ' Note that this is a DESTRUCTIVE replacement
        '     operation. A preferable option might be
        '     to generate a NEW hierarchical Dictionary
        '     replicating the original, with the desired
        '     key substitution. Will consider that for
        '     a later implementation.
        '
        ' Note also that error checking/handling
        '     in this implementation is presently minimal.
        '     A more robust process should also be considered.
        '
        Dim wk As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim ky As Object

        wk = dcTreeMembersWithKey(tg, dc)
        ck = dcTreeMemWithReplcmt(rp, wk)

        If ck.Count > 0 Then
            Stop
        Else
            With wk : For Each ky In .Keys
                    With dcOb(obOf(.Item(ky)))
                        ' A Dictionary object is assumed, here.
                        ' Though typically risky in a With block,
                        ' it SHOULD be guaranteed here,
                        ' so no error should occur.
                        ' Don't be surprised if it does, though.
                        If .Exists(rp) Then
                            Stop 'because this
                            'should NOT be happening!
                            'dcTreeMemWithReplcmt
                            'should have caught
                            'any replacement key
                            'collisions already.
                        Else
                            'a proper error handler might
                            'be desired here in future
                            'On Error Resume Next
                            'for now, keep disabled

                            'note order of operations here
                            .Add(rp, .Item(tg)) 'FIRST

                            .Remove(tg) 'ONLY AFTER
                            'associated Item is added
                            'under replacement Key

                            'this ensures the associated Item
                            'is retained under AT LEAST ONE Key,
                            'and not lost in the event of some
                            'fault or error, which really
                            'shouldn't occur, BUT...

                            'On Error GoTo 0
                            'potential error handler
                            'to end here, unless moved
                        End If : End With
                Next : End With
        End If

        dcTreeReKeyedInPlc = dc
    End Function

    Public Function userChoiceFromDc(Optional dc As Scripting.Dictionary = Nothing, Optional ifNone As Object = Nothing) As Object
        '
        ' userChoiceFromDc (formerly d2g3f2)
        '     Request User Selection from
        '     a Dictionary of options.
        '
        '     A list of Dictionary Keys is
        '     presented to the user. After
        '     User selects a Key, matching
        '     Item is returned for use.
        '
        Dim ck As VbMsgBoxResult
        Dim msNoSel As String
        Dim rp As Object
        Dim rt As Object

        ' REV[2023.05.17.1304]
        ' add ifNone processing to present
        ' User with information on default
        ' option(s), if supplied
        On Error Resume Next
        Err.Clear()
        msNoSel = CStr(ifNone)

        If Err.Number = 0 Then
            If Len(msNoSel) > 0 Then
                msNoSel = "Use default value (" & msNoSel & ")?"
            End If
        Else
            msNoSel = ""
            Err.Clear()

            If TypeOf ifNone Is Object Then
                If Not ifNone Is Nothing Then
                    msNoSel = Join({
                    "Use default " & TypeName(ifNone) & " Object?",
                    "(Object details not available)"
}, vbCrLf)
                End If
            Else
                Stop
            End If
        End If
        On Error GoTo 0

        If Len(msNoSel) > 0 Then
            msNoSel = vbCrLf & msNoSel
        End If
        msNoSel = Join({(
        "User selection was requested" _
        , "with no available options!" _
        , msNoSel
    ), vbCrLf})

        If dc Is Nothing Then
            rt = {userChoiceFromDc(dcAiDocsVisible())}
        Else
            If dc.Count > 0 Then
                rp = nuSelFromDict(dc
            ).GetReply()
                ' , , , , , _
                Join({
                    "No option selected!",
                    msNoSel
                }, vbCrLf)
                If dc.Exists(rp) Then
                    rt = {dc.Item(rp)}
                Else
                    rt = {ifNone}
                End If
            Else
                ck = MsgBox(Join({
                "User selection was requested",
                "with no available options!"
                }, vbCrLf), vbOKOnly, "No Options!"
                )
                If ck = vbNo Then
                    rt = {Nothing}
                    'not the best option, but
                    'not sure what else to do
                Else
                    rt = {ifNone}
                End If
            End If
        End If

        If Not IsNothing(rt(0)) Then
            userChoiceFromDc = rt(0)
        Else
            userChoiceFromDc = rt(0)
        End If
    End Function

    Public Function dcGnsPrpPtDvl_2021_1112(ThisApplication As Inventor.Application,
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim dc01 As Scripting.Dictionary
        Dim dcVlGn As Scripting.Dictionary
        Dim rs As ADODB.Record
        ''
        Dim aiPartNum As String
        Dim aiFamily As String
        Dim aiSubType As String
        ''
        'Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        ''
        Dim prPartNum As Inventor.Property
        Dim prFamily As Inventor.Property
        ''
        'Dim aiPartNum   As String 'will be same as gnPartNum
        'Dim aiPartFam   As String
        'Dim aiMatlNum   As String
        'Dim aiMatlFam   As String
        'Dim aiMatlQty   As Double
        'Dim aiQtyUnit   As String
        Dim aiBomType As Inventor.BOMStructureEnum
        ''
        ''
        rt = New Scripting.Dictionary
        dc01 = New Scripting.Dictionary

        With oDoc
            ' Get Property s
            With .Propertys
                ' aiPropsUser = .Item(gnCustom)
                aiPropsDesign = .Item(GnDesign)
            End With
            aiBomType = .ComponentDefinition.BOMStructure
            aiSubType = .SubType
        End With

        ' Get Part Number and Family
        ' Properties from Design 
        With aiPropsDesign
            prPartNum = .Item(PnPartNum)
            prFamily = .Item(PnFamily)
        End With

        ' Get Values of Part Number
        ' and Family Properties
        aiPartNum = prPartNum.Text
        aiFamily = prFamily.Text
        ' NOTE[2021.11.12]
        '     The preceding three sections
        '     can PROBABLY be consolidated
        '     into one, using fewer variables
        '     and probably just one With block

        'dc01
        With CnGnsDoyle()
            On Error Resume Next
            Err.Clear()
            rs = .Execute(
            SqlOf_ASDF(ThisApplication, aiPartNum)
        ) '
            If Err.Number = 0 Then
                dcVlGn = DcFromAdoRSrow(rs, "")
                With dcVlGn
                    'gnPartNum = .Item("Item")
                    'gnPartFam = .Item("Family")
                    'gnBomType = .Item("bomStr")
                    '' fdOrder = .Item("Ord")
                    'gnMatlNum = .Item("Material")
                    'gnMatlFam = .Item("MtFamily")
                    'gnMatlQty = .Item("Qty")
                    'gnQtyUnit = .Item("Unit")
                End With

                With rs
                    If .BOF And .EOF Then
                    Else
                        With .Fields
                            ' fdItem = .Item("Item")
                            'gnPartNum = .Item("Item")
                            'should ALWAYS match aiPartNum
                            'IF it's found in Genius
                            'otherwise, always BLANK

                            ' fdFamly = .Item("Family")
                            'gnPartFam = .Item("Family").Value

                            'gnBomType = .Item("bomStr").Text

                            ' fdOrder = .Item("Ord")

                            ' fdMatrl = .Item("Material")
                            'gnMatlNum = .Item("Material").Text

                            ' fdMtFam = .Item("MtFamily")
                            'gnMatlFam = .Item("MtFamily").Text

                            ' fdQty = .Item("Qty")
                            'gnMatlQty = .Item("Qty").Text

                            ' fdUnit = .Item("Unit")
                            'gnQtyUnit = .Item("Unit").Text

                            'Stop 'to check things out
                            'not doing anything else with this yet
                            'but want to start matching against model
                        End With

                        .MoveNext
                        If Not .EOF Then
                            Stop 'to handle multiple raw materials
                            Debug.Print("") 'Breakpoint Landing
                        End If
                    End If

                    .Close()
                End With
            Else
                Debug.Print(Err.Number)
                Debug.Print(Err.Description)
                Stop
            End If
            On Error GoTo 0
            .Close()
        End With

        If aiBomType = BOMStructureEnum.kNormalBOMStructure Then
            If aiSubType = GuidSheetMetal Then
                'try to get flat pattern data here
            End If
        ElseIf aiBomType = BOMStructureEnum.kPurchasedBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kPhantomBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kInseparableBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kReferenceBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kNormalBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kPhantomBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kDefaultBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kVariesBOMStructure Then
            Stop
        ElseIf aiBomType = BOMStructureEnum.kDefaultBOMStructure Then
            Stop
        Else
            Stop
        End If

        ''
        dcGnsPrpPtDvl_2021_1112 = rt
    End Function

    Public Function dcGeniusPropsPartRev20180530_broken2(ThisApplication As Inventor.Application,
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        'Dim dcPr As Scripting.Dictionary
        Dim dcVlGn As Scripting.Dictionary
        Dim dcProp As Scripting.Dictionary
        Dim dcVlPr As Scripting.Dictionary
        Dim dcVlAi As Scripting.Dictionary
        Dim dcVlFP As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ky As Object
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
        ' UPDATE[2021.11.08] MAJOR CHANGE
        '     Overhaul variable names to better
        '     reflect TWO distinct value s
        '         one from Genius
        '         another from Inventor
        '     in order to better compare
        '     and synchronize them.
        '
        ' First  are the Genius variables:
        ' the original , renamed en masse:
        '
        Dim gnPartNum As String 'was pnModel
        Dim gnPartFam As String 'was ptFamily
        Dim gnMatlNum As String 'was pnStock
        Dim gnMatlFam As String 'was mtFamily
        Dim gnMatlQty As Double 'was qtRawMatl
        Dim gnQtyUnit As String 'was qtUnit
        Dim gnBomType As Inventor.BOMStructureEnum
        '
        ' Second  are the new Inventor variables.
        ' These should replace the Genius instances
        ' anywhere their values are taken
        ' from the model.
        '
        Dim aiPartNum As String 'will be same as gnPartNum
        Dim aiPartFam As String
        Dim aiMatlNum As String
        Dim aiMatlFam As String
        Dim aiMatlQty As Double
        Dim aiQtyUnit As String
        Dim aiBomType As Inventor.BOMStructureEnum
        '
        '
        ''
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData
        ' UPDATE[2021.11.03]:
        '
        '
        Dim rs As ADODB.Record
        Dim fdItem As ADODB.Field
        Dim fdFamly As ADODB.Field
        Dim fdOrder As ADODB.Field
        Dim fdMatrl As ADODB.Field
        Dim fdMtFam As ADODB.Field
        Dim fdQty As ADODB.Field
        Dim fdUnit As ADODB.Field

        If dc Is Nothing Then
            dcGeniusPropsPartRev20180530_broken2 =
        dcGeniusPropsPartRev20180530_broken2(
            oDoc, New Scripting.Dictionary
        )
        Else
            aiBomType = oDoc.ComponentDefinition.BOMStructure
            'UPDATE[2021.11.11]
            '     Moved Property  collection
            '     to top of program to permit
            '     collection of Design Properties
            '     in second step. Also pulled up
            '     BOM Structure capture (above)
            '     along with the Values of
            '     of Design Properties.

            With oDoc
                With .Propertys
                    aiPropsDesign = .Item(GnDesign)
                    aiPropsUser = .Item(GnCustom)
                End With

                aiBomType = .ComponentDefinition.BOMStructure

                If aiBomType = BOMStructureEnum.kNormalBOMStructure Then
                    If .SubType = GuidSheetMetal Then
                    End If
                End If
            End With

            ' Part Number and Family properties
            ' are from Design, NOT Custom 
            With aiPropsDesign 'we know they're present
                'so we can grab them directly
                prPartNum = .Item(PnPartNum)
                prFamily = .Item(PnFamily)
            End With
            aiPartNum = prPartNum.Text
            aiPartFam = prFamily.Text

            dcProp = dcGnsPropsPart(oDoc, , 0) 'dcAiPropsIn
            dcVlPr = New Scripting.Dictionary
            With dcProp
                .Add(PnPartNum, prPartNum)
                .Add(PnFamily, prFamily)

                For Each ky In .Keys
                    InvProperty = aiProperty(.Item(ky))
                    If InvProperty Is Nothing Then
                        Stop
                    Else
                        dcVlPr.Add(ky, InvProperty.Text)
                    End If
                Next

                If .Exists(PnRawMaterial) Then
                    prRawMatl = .Item(PnRawMaterial)
                    aiMatlNum = prRawMatl.Text
                Else
                    aiMatlNum = ""
                End If

                If .Exists(PnRmUnit) Then
                    prRmUnit = .Item(PnRmUnit)
                    aiQtyUnit = prRmUnit.Text
                Else
                    aiQtyUnit = ""
                End If

                If .Exists(PnRmUnit) Then
                    prRmQty = .Item(PnRmQty)
                    aiMatlQty = prRmQty.Text
                Else
                    aiMatlQty = 0
                End If
            End With
            Debug.Print("=== Check Existing Model Genius Properties ===")
            Debug.Print(DumpLsKeyVal(dcVlPr, "="))
            Debug.Print("")
            Stop

            ' NOTE[2021.11.11]
            '     Assignment of initial rt Dictionary
            '     now essentially duplicates the new
            '     process now preceding this section.
            '     The only difference is, that version
            '     does NOT apply Genius Property col-
            '     lection to the supplied Dictionary dc.
            rt = dcGnsPropsPart(oDoc, dc, 0) 'dcAiPropsIn
            dcVlAi = New Scripting.Dictionary
            With rt
                .Add(PnPartNum, prPartNum)
                .Add(PnFamily, prFamily)

                For Each ky In .Keys
                    InvProperty = aiProperty(.Item(ky))
                    If InvProperty Is Nothing Then
                        Stop
                    Else
                        dcVlAi.Add(ky, InvProperty.Text)
                    End If
                Next
                InvProperty = Nothing
            End With
            '     Ultimately, processes which populate
            '     returned Dictionary rt, and  the
            '     Properties it should receive, should
            '     be moved toward the end of the function.

            With CnGnsDoyle()
                'Pre-clear all relevant variables
                'to be  from query results,
                'if available.

                'gnPartNum = aiPartFam
                'gnPartFam = ""
                ' fdOrder = .Item("Ord")
                'gnMatlNum = ""
                'gnMatlFam = ""
                'gnMatlQty = 0
                'gnQtyUnit = ""
                'gnBomType =BOMStructureEnum.kDefaultBOMStructure
                'use this to indicate no BOM type
                'or structure returned from Genius

                On Error Resume Next
                Err.Clear()
                rs = .Execute(
                SqlOf_ASDF(ThisApplication, aiPartNum)
            ) '
                If Err.Number = 0 Then
                    dcVlGn = DcFromAdoRSrow(rs, "")
                    With dcVlGn
                        gnPartNum = .Item("Item")
                        gnPartFam = .Item("Family")
                        gnBomType = .Item("bomStr")
                        ' fdOrder = .Item("Ord")
                        gnMatlNum = .Item("Material")
                        gnMatlFam = .Item("MtFamily")
                        gnMatlQty = .Item("Qty")
                        gnQtyUnit = .Item("Unit")
                    End With

                    With rs
                        If .BOF And .EOF Then
                        Else
                            With .Fields
                                ' fdItem = .Item("Item")
                                'gnPartNum = .Item("Item")
                                'should ALWAYS match aiPartNum
                                'IF it's found in Genius
                                'otherwise, always BLANK

                                ' fdFamly = .Item("Family")
                                'gnPartFam = .Item("Family").Value

                                'gnBomType = .Item("bomStr").Text

                                ' fdOrder = .Item("Ord")

                                ' fdMatrl = .Item("Material")
                                'gnMatlNum = .Item("Material").Text

                                ' fdMtFam = .Item("MtFamily")
                                'gnMatlFam = .Item("MtFamily").Text

                                ' fdQty = .Item("Qty")
                                'gnMatlQty = .Item("Qty").Text

                                ' fdUnit = .Item("Unit")
                                'gnQtyUnit = .Item("Unit").Text

                                'Stop 'to check things out
                                'not doing anything else with this yet
                                'but want to start matching against model
                            End With

                            .MoveNext
                            If Not .EOF Then
                                Stop 'to handle multiple raw materials
                                Debug.Print("") 'Breakpoint Landing
                            End If
                        End If

                        .Close()
                    End With
                Else
                    Debug.Print(Err.Number)
                    Debug.Print(Err.Description)
                    Stop
                End If
                On Error GoTo 0
                .Close()
            End With

            Debug.Print("== Prop Check ==")
            Debug.Print("---- Genius ----")
            Debug.Print(DumpLsKeyVal(dcVlGn, "="))
            Debug.Print("--- Inventor ---")
            Debug.Print(DumpLsKeyVal(dcVlPr, "=")) 'dcVlAi
            Debug.Print("================")
            Stop

            With oDoc
                'UPDATE[2021.11.11]
                '     Moved Property  collection
                '     to top of program, along with
                '     collection of Design Properties
                '     and their values. BOM Structure
                '     as well.

                ' We should check HERE for possibly misidentified purchased parts
                ' UPDATE[2021.11.08]
                '     Another MAJOR overhaul, here:
                '     change Purchased Parts identification
                '     to defer to Genius. Only attempt to guess
                '     when no value comes back from Genius.
                'Stop 'BKPT-2021-1108-1608
                ' CHANGE NEEDED[2021.11.08]:
                '     indeterminate -- stopping work @endOfDay
                '     effort here is to separate collection
                '     and potential reassignment of
                '     based on Part's family, file location,
                '     and whatever other criteria, if any.
                '
                '     Likely need a counterpart variable
                '     which takes its value from the Model.
                '     The most likely Genius equivalent is
                '     probably the ItemType field in view
                '     table vgMfiItems, which will need
                '     translation.
                '
                If gnBomType = BOMStructureEnum.kDefaultBOMStructure Then
                    'Genius didn't return an Item type
                    'or BOM structure. We need to get it here.

                    ' BKPT-2021-1109-1042
                    '     Checkpoint here. Verify desired
                    '     behavior here prior to removal.
                    Stop

                    ' Get BOM Structure type, correcting if appropriate,
                    ' and prepare Family value for part, if purchased.
                    '
                    ' UPDATE[2018.02.06]
                    '     Using new UserForm, see below
                    ' UPDATE[2018.05.31]
                    '     Combined both InStr checks by addition
                    '     to generate a single test for > 0
                    '     If EITHER string match succeeds, the total
                    '     SHOULD exceed zero, so this SHOULD work.
                    ' UPDATE[2021.11.08]
                    '     Removed extraneoous code previously
                    '     disabled under preceding update[2018.05.31]
                    '     Also reseparated InStr checks previously combined
                    If aiBomType = BOMStructureEnum.kPurchasedBOMStructure Then 'it's purchased.
                        'Just assume that's what it's supposed to be.
                        gnBomType = aiBomType
                    ElseIf InStr(1,
                    "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                    "|" & aiPartFam & "|"
                ) > 0 Then
                        'it needs to be  purchased.
                        gnBomType = BOMStructureEnum.kPurchasedBOMStructure
                    Else
                        'might need to ask User
                        If InStr(1, oDoc.FullFileName,
                        "\Doyle_Vault\Designs\purchased\"
                    ) > 0 Then 'it'a LIKELY purchased.
                            'Double check with User.
                            ck = newFmTest2().AskAbout(oDoc, ,
                            "Is this a Purchased Part?"
                        )
                        Else
                            ck = vbNo
                        End If

                        'Stop 'BKPT-2021-1105-0942
                        ' CHANGE NEEDED[2021.11.05]:
                        '     ONLY COLLECT desired BOMStructure here
                        '     while keeping track of current value.
                        '     Reassignment should take place along
                        '     with collective property changes
                        ' UPDATE[2021.11.09]
                        '     This section now reduced to ting
                        '     gnBomType from User response, if any.
                        '     Code to assign Model BOM structure
                        '     moved toward bottom for further work.
                        '
                        ' Check process below replaces duplicate check/responses above.
                        If ck = vbYes Then 'User said it IS purchased
                            gnBomType = BOMStructureEnum.kPurchasedBOMStructure
                        Else
                            gnBomType = aiBomType
                            dcVlGn.Item("bomStr") = gnBomType
                        End If

                        'Request #2: Change Cost Center iProperty.
                        'If BOMStructure = Purchased and not content center,
                        'then Family = D-PTS, else Family = D-HDWR.
                        '
                        ' UPDATE[2018.05.30]: Value produced here
                        '     will now be held for later processing,
                        '     more toward the end of this function.
                        ' UPDATE[2021.11.09]
                        '     Changed to  target (Genius)
                        '     Family, and ONLY if not  already.
                        '     '
                        '     MIGHT want to  up a more robust check
                        '     system, but see how this holds up, first.
                        '     '
                        If Len(gnPartFam) = 0 Then
                            If gnBomType = BOMStructureEnum.kPurchasedBOMStructure Then
                                If .IsContentMember Then
                                    Stop 'BKPT-2021-1105-0946
                                    gnPartFam = "D-HDWR"
                                Else
                                    Stop 'BKPT-2021-1105-0947
                                    gnPartFam = "D-PTS"
                                    'NOTE: NON Content Center members
                                    '       might still be D-HDWR
                                    '       Additional checks might
                                    '       be recommended
                                End If
                            Else
                            End If
                        End If
                    End If
                    'Else 'disabled unless/until needed, or removed
                    'aiPartFam = gnPartFam 'no, DON'T!
                    ' UPDATE[2021.11.09]: Disabled this
                    '     Keep the Model Part Family value AS IS
                    '     so it may be used to check for equality
                    '     and the need to update at the end.
                    '     '
                End If

                With .ComponentDefinition
                    ' Request #1: Get the Mass in Pounds
                    ' and add to Custom Property GeniusMass
                    With .MassProperties
                        'Stop 'BKPT-2021-1110-1551
                        ' CHANGE NEEDED[2021.11.10]
                        '     '
                        '     '
                        '     '
                        If dcVlPr.Exists(PnMass) Then
                            If System.Math.Round(CvMassKg2LbM * .Mass, 4) - CDbl(dcVlPr.Item(PnMass)) = 0 Then
                            Else
                                'Stop
                                dcVlPr.Item(PnMass) = System.Math.Round(CvMassKg2LbM * .Mass, 4)
                            End If
                        Else
                            dcVlPr.Add(PnMass, System.Math.Round(
                            CvMassKg2LbM * .Mass, 4
                        ))
                        End If
                        ' UPDATE[2021.11.09]
                        '     Part Mass Value now assigned
                        '     to new Values Dictionary, instead
                        '     of directly to Genius Mass Property.
                        '     That assignment now moved toward the
                        '     end of this function, where it can
                        '     be  alongside other Properties
                        '     in one straight process.
                        '     '
                        '     Note that value in Genius is not
                        '     yet collected for comparison. This
                        '     will likely require modification
                        '     of an SQL query, and a new variable
                        '     or Dictionary to keep track of it.
                    End With
                End With
                ' At this point, gnPartFam SHOULD be 
                ' to a non-blank value if Item is purchased.
                ' We should be able to check this later on,
                ' if Item BOMStructure is NOT Normal

                'Stop 'BKPT-2021-1109-1053
                ' HERE is where it starts to get interesting
                ' Actually, just a little further down, where
                ' Part SubType is checked for Sheet Metal.
                ' At that point, the function divides into two
                ' LONG, and possibly nearly identical branches.
                ' Ideally, these should be refactored, with as
                ' much of their processes as possible combined
                ' into a single path.

                'Request #4: Change Cost Center iProperty.
                'If BOMStructure = Normal, then Family = D-MTO,
                'else if BOMStructure = Purchased then Family = D-PTS.
                'If aiBomType = BOMStructureEnum.kNormalBOMStructure Then
                ' Get Custom Properties
                'Stop 'BKPT-2021-1105-1144
                ' CHANGE NEEDED[2021.11.05]:
                '     these properties should NOT
                '     be added immediately, but only
                '     when it's time to  them,
                '     towards the END of this function.
                ' UPDATE[2021.11.09]
                '     Custom Property collection/generation
                '     moved into Normal BOM Part handling, as
                '     no earlier usage appears to take place.
                '     '
                '     If possible, may wish to move even further.
                '     Plan to review later, as time permits.
                ' UPDATE[2021.11.10]
                '     Disabled Genius Property collection here
                '     since a Dictionary of ALL Genius Properties
                '     is generated towards the beginning.
                '     '
                '     '
                'With rt
                '    If .Exists(pnRawMaterial) Then  prRawMatl = .Item(pnRawMaterial)
                '     prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                '    If .Exists(pnRmUnit) Then  prRmUnit = .Item(pnRmUnit)
                '     prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                '    If .Exists(pnRmQty) Then  prRmQty = .Item(pnRmQty)
                '     prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                'End With
                '     Collecting them at this point
                '     might still be appropriate,
                '     or it may be more desirable
                '     to hold off until later.
                '     '
                ' UPDATE[2021.11.08]
                '     Design Properties have been moved
                '     toward the top, as proposed.
                '     Commentary recommending this
                '     has been removed as extraneous.
                '     '
                ' UPDATE[2021.11.09]
                '     BOM Structure collection has also been
                '     moved up, alongside Design Properties,
                '     using renamed variable aiBomType
                '     (formerly bomStruct)
                '     '

                ''----------------------------------------------------'
                'If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                '----------------------------------------------------'
                'Request #3:
                '   Get sheet metal extent area
                '   and add to custom property "RMQTY"
                ' UPDATE[2021.11.10]
                '     Now collecting Flat Pattern Values
                '     instead of the Properties for them.
                '     If necessary, Properties should be
                '     assigned in a separate function
                ' CHANGE NEEDED[2021.11.05]:
                '     not quite sure on this one yet,
                '     but dcFlatPatProps might need its
                '     own  of revisions to generate
                '     assignment recommendations WITHOUT
                '     performing them itself
                ' UPDATE[2021.11.10]
                '     Embedded Flat Pattern Property collection
                '     in bypassed If branch. Preceding Stop,
                '     when enabled, offers user/developer
                '     an opportunity to run it, if desired.
                Stop 'BKPT-2021-1105-1105
                If True Then
                    dcVlFP = dcFlatPatVals(
                    .ComponentDefinition
                ) 'dcVlAi
                    With dcVlFP
                        aiMatlFam = .Item("mtFamily")
                        .Remove("mtFamily")
                        For Each ky In .Keys
                            If dcVlPr.Exists(ky) Then
                                If CStr(dcVlPr.Item(ky)) = CStr(.Item(ky)) Then
                                Else
                                    Stop
                                    ' need to add value to a new 'change' Dictionary
                                End If
                            Else
                                Stop
                                ' need to add value to a new 'change' Dictionary
                            End If
                        Next
                    End With
                    Stop
                Else
                    rt = dcFlatPatProps(.ComponentDefinition, rt)
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
                ''   FINDS NO .FlatPattern, then there should
                ''   BE NO sheet metal part number!
                'If prRawMatl Is Nothing Then
                '    If rt.Exists("OFFTHK") Then
                '        ' UPDATE[2018.05.30]:
                '        '     Restoring original key check
                '        '     and adding code for debug
                '        '     Previously changed to "~OFFTHK"
                '        '     to avoid this block and its issues.
                '        '     (Might re-revert if not prepped to fix now)
                '        Debug.Print(aiProperty(rt.Item("OFFTHK")).Text)
                '        Stop 'because we're going to need to do something with this.

                '        gnMatlNum = "" 'Originally the ONLY line in this block.
                '        ' A more substantial response is required here.

                '        If 0 Then Stop '(just a skipover)
                '    Else
                '        If Len(gnMatlNum) = 0 Then
                '            Stop 'because we don't know IF this is sheet metal yet
                '            gnMatlNum = ptNumShtMetal(.ComponentDefinition)
                '        End If
                '    End If
                'Else
                ''  ACTION ADVISED[2018.09.14]:
                ''  gnMatlNum can probably be 
                ''  to prRawMatl.Text and THEN
                ''  checked for length to see
                ''  if lookup needed.
                ''  This might also allow us to check
                ''  for machined or other non-sheet
                ''  metal parts.

                Stop
                ' !!!WARNING!!![2021.11.04]:
                ' Following section has been shuffled
                ' and should be considered HIGHLY
                ' UNSTABLE until verified functional
                ' and SAFE! TWO Stop commands are
                ' placed to emphasize the need for
                ' EXTREME CAUTION at this point
                Stop
                ' UPDATE[2021.11.04]:
                '     This section is being adjusted
                '     in an attempt to improve the raw
                '     material determination process.
                '
                '     This particular segment should
                '     ONLY be invoked if gnMatlNum is not
                '     successfully retrieved from Genius
                '
                If Len(gnMatlNum) = 0 Then
                    'no stock retrieved from Genius
                    'attempt to retrieve from Model
                    'gnMatlNum = aiMatlNum

                    If Len(aiMatlNum) > 0 Then 'gnMatlNum
                        'need to verify it against Genius
                        'by retrieving its Family there
                        '
                        ' This With block copied and modified [2021.03.11]
                        ' from elsewhere in this function as a temporary measure
                        ' to address a stopping situation later in the function.
                        ' See comment below for details.
                        '
                        ' UPDATE[2021.11.04]:
                        '     This section MIGHT be removed in future,
                        '
                        With CnGnsDoyle().Execute(
                    "select Family " &
                    "from vgMfiItems " &
                    "where Item='" & gnMatlNum & "';"
                )
                            If .BOF Or .EOF Then
                                'Stop 'because Material value likely invalid
                                Stop 'because we do NOT want to  gnMatlNum!
                                ' want to assign it to a separate RETURN variable
                                ' or most likely, the return Dictionary.
                                gnMatlNum = ptNumShtMetal(oDoc.ComponentDefinition)
                                Debug.Print("") 'Breakpoint Landing
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
                            Else
                                ''  This section retained from source,
                                ''  but disabled to avoid potential issues
                                ''  with subsequent operations, just in case
                                ''  anything depends on gnMatlFam remaining
                                ''  uninitialized up to that point.
                                ' UPDATE[2021.11.09]
                                '     Re-enabling Genius Material
                                '     Family assignment, as it SHOULD
                                '     be  to match what Genius
                                '     returns from this query.
                                '     '
                                '     Might not be the best place
                                '     to do this, though. If ptNumShtMetal
                                '     returns a valid Material Item above,
                                '     a Family is still needed.
                                '     '
                                '     NOTE: Fix disabled With block between runs
                                ''  With .Fields
                                Stop 'because we do not want to  gnMatlFam
                                '   for same reasons as above
                                gnMatlFam = .Fields.Item("Family").Value
                                ''  End With
                            End If
                        End With
                    End If
                End If

                If Len(gnMatlNum) = 0 Then
                    ' UPDATE[2018.05.30]:
                    '     Pulling ALL code/text from this section
                    '     to get rid of excessive cruft.
                    '
                    '     In fact, reversing logic to go directly
                    '     to User Prompt if no stock identified
                    '
                    '     IN DOUBLE FACT, hauling this WHOLE MESS
                    '     RIGHT UP after initial gnMatlNum assignment
                    '     to prompt user IMMEDIATELY if no stock found
                    With newFmTest1()
                        If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop

                        bd = nuAiBoxData().UsingInches.SortingDims(
                        oDoc.ComponentDefinition.RangeBox
                    )
                        ck = .AskAbout(oDoc,
                        "No Stock Found! Please Review" _
                        & vbCrLf & vbCrLf & bd.Dump(0)
                    )

                        If ck = vbYes Then
                            ' UPDATE[2018.05.30]:
                            '     Pulling some extraneous commented code
                            '     from here and beginning of block
                            With .ItemData
                                If .Exists(PnFamily) Then
                                    gnPartFam = .Item(PnFamily)
                                    Debug.Print(PnFamily & "=" & gnPartFam)
                                End If

                                If .Exists(PnRawMaterial) Then
                                    gnMatlNum = .Item(PnRawMaterial)
                                    Debug.Print(PnRawMaterial & "=" & gnMatlNum)
                                End If
                            End With
                            If 0 Then Stop 'Use this for a debugging shim
                        End If
                    End With
                ElseIf Left$(gnMatlNum, 2) = "LG" Then 'it's probably lagging
                    Debug.Print(aiPartNum & ": PROBABLE LAGGING")
                    Debug.Print("  TRY TO IDENTIFY, AND FILL IN BELOW.")
                    Debug.Print("  PRESS ENTER ON gnMatlNum LINE WHEN")
                    Debug.Print("  COMPLETED, THEN F5 TO CONTINUE.")
                    Debug.Print("  gnMatlNum = """ & gnMatlNum & """")
                    Stop
                End If

                'If Len(gnMatlNum) > 0 Then 'and ONLY then
                'do we look for a Raw Material Family!

                ' NOTE[2021.11.10]
                '     This query is probably WAY more than needed here.
                '     Spec fields are probably not needed at all,
                '     and it's not clear which of the others might be.
                '     '
                '     It might also be possible to REMOVE this query
                '     based on the earlier one, which would return
                '     Material Family along with Part Family,
                '     providing the Part/Item were found in Genius.
                '    '     '
                '    With CnGnsDoyle().Execute(
                '    "select Family " &
                '    "from vgMfiItems " &
                '    "where Item='" & gnMatlNum & "';"
                ') '", Description1, Unit, " & _
                '    "Specification1, Specification2, Specification3, " & _
                '    "Specification4, Specification5, Specification6, " & _
                '    "Specification7, Specification8, Specification9, " & _
                '    "Specification15, Specification16 " & _
                ''
                ' UPDATE[2021.11.10]
                '     Removed (likely) unneeded fields from query text.
                '     Will keep a lookout for any resulting errors.
                If .BOF Or .EOF Then
                    Stop 'because Material value likely invalid
                    ''  ACTION ADVISED[2018.09.14]:
                    ''  Will need to address this situation
                    ''  in a more robust manner.
                    ''  A more thorough query above
                    ''  might also be called for.
                Else
                    With .Fields
                        If Len(gnMatlFam) > 0 Then
                            If gnMatlFam = .Item("Family").Value Then
                            Else
                                Stop
                                'as with gnMatlNum, want to be careful
                                'about changing gnMatlFam, although
                                'in this case, it SHOULD be okay
                                'since it should just align with
                                'the selected material's Family
                            End If
                        Else
                            gnMatlFam = .Item("Family").Value
                        End If
                    End With
                    ' NOTE[2021.11.10]
                    '     Else branch should PROBABLY end here
                    '     to permit Record to be closed,
                    '     and probably a new If/Then block
                    '     proceed based on results.
                    '     '

                    ' UPDATE[2021.06.18]:
                    '     New pre-check for Material Item
                    '     in Purchased Parts Family.
                    '     VERY basic handler simply
                    '     maps Material Family to D-BAR
                    '     to force extra processing below.
                    '     Further refinement VERY much needed!
                    If gnMatlFam Like "?-MT*" Then
                        'Debug.Print(aiPartNum & " [" & aiMatlNum & "]: " & aiPropsDesign(pnDesc).Text
                        Debug.Print(aiPartNum & "[" & prRmQty.Text & gnQtyUnit & "*" & gnMatlNum & ": " & aiPropsDesign.item(PnDesc).Text & "]") ' aiMatlNum
                        Stop 'FULL Stop!
                    ElseIf gnMatlFam = "D-PTS" Then
                        gnPartFam = "D-RMT"
                        Stop 'NOT SO FAST!
                        gnMatlFam = "D-BAR"
                    ElseIf gnMatlFam = "R-PTS" Then
                        gnPartFam = "R-RMT"
                        Stop 'NOT SO FAST!
                        gnMatlFam = "D-BAR"
                    End If

                    If gnMatlFam = "DSHEET" Then
                        'We should be okay. This is sheet metal stock

                        ' UPDATE[2021.11.04]:
                        '     Expanding gnPartFam and gnQtyUnit
                        '     assignments to check for pre-
                        '     existing values, and validate
                        '     them if found.
                        If Len(gnPartFam) = 0 Then
                            gnPartFam = "D-RMT"
                        Else
                            If gnPartFam = "D-RMT" Then
                            Else
                                Stop 'because we have
                                'an unexpected situation
                            End If
                        End If

                        If Len(gnQtyUnit) = 0 Then
                            gnQtyUnit = "FT2"
                        Else
                            If gnQtyUnit = "FT2" Then
                            Else
                                Stop 'because we have
                                'an unexpected situation
                            End If
                        End If

                        ' UPDATE[2018.05.30]:
                        '     Moving part family assignment
                        '     to this section for better mapping
                        '     and updating to new Family names
                        '     as well as pulling up gnQtyUnit assignment
                        Stop 'BKPT-2021-1105-1120
                        ' CHANGE NEEDED[2021.11.05]:
                        '     probably want to demote this ElseIf
                        '     along with the subsequent Else into
                        '     a replacement Else clause, and thus
                        '     allow for on post-interactive check
                        '     following whichever branch is taken.
                        '     '
                        '     As is, a separate check is required
                        '     within but this ElseIf and the Else
                        '     '
                        ' UPDAGE[SAME_DAY]:
                        '     Change completed. All that remains
                        '     is to indent the embedded With block
                        '     when safe to do so.
                    Else
                        If gnMatlFam = "D-BAR" Then
                            ' UPDATE[2021.06.18]:
                            '     Added check for Part Family already 
                            '     to more properly handle new situation (above)
                            If Len(gnPartFam) = 0 Then
                                gnPartFam = "R-RMT"
                                'mignt not want to use
                                'fixed constant like this
                                'see gnQtyUnit below
                            Else
                                If gnPartFam = "R-RMT" Then
                                Else
                                    Stop
                                End If
                                Debug.Print("") 'Breakpoint Landing
                                'Stop
                            End If

                            If Len(gnQtyUnit) = 0 Then
                                gnQtyUnit = "IN" 'prRmUnit.Text '
                            Else
                                If gnQtyUnit = "IN" Then 'prRmUnit.Text '
                                Else
                                    Stop
                                End If
                                Debug.Print("") 'Breakpoint Landing
                                'Stop
                            End If
                            ''may want function here
                            ' UPDATE[2018.05.30]: As noted above
                            '     Will keep Stop for now
                            '     pending further review,
                            '     hopefully soon
                            Debug.Print(aiPartNum & " [" & gnMatlNum & "]: " & aiPropsDesign.item(PnDesc).Text) 'aiMatlNum
                            ' UPDATE[2021.03.11]: Replaced
                            ' aiPropsDesign.Item(pnPartNum)
                            ' with prPartNum (and now aiPartNum)
                            ' since it's used in several places
                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), gnQtyUnit, ". IF CHANGE NEEDED,")
                            Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                            Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                            'Stop 'BKPT-2021-1105-1137
                            ' CHANGE NEEDED[2021.11.05]:
                            '     indent the following With,
                            '     when possible to do so
                            '     without reting project
                            With oDoc.ComponentDefinition.RangeBox
                                Debug.Print(
                                                        (.MaxPoint.X - .MinPoint.X) / 2.54,
                                                        (.MaxPoint.Y - .MinPoint.Y) / 2.54,
                                                        (.MaxPoint.Z - .MinPoint.Z) / 2.54
                                                    )
                            End With
                            'Debug.Print("CURRENT RAW MATERIAL QUANTITY (";
                            'Debug.Print(CStr(prRmQty.Text), ") IS SHOWN BELOW."
                            'Debug.Print("IF NOT CORRECT, YOU MAY TYPE A NEW VALUE"
                            'Debug.Print("IN ITS PLACE, AND PRESS ENTER TO CHANGE IT."
                            'Debug.Print("SOME SUGGESTED VALUES INCLUDE X, Y, AND Z"
                            'Debug.Print("EXTENTS (ABOVE) OR YOU MAY SUPPLY YOUR OWN."
                            'Debug.Print(""
                            'Debug.Print(""
                            'Debug.Print("YOU MAY ALSO CHANGE THE UNIT OF MEASURE BELOW,"
                            'Debug.Print("IF DESIRED. BE SURE TO PRESS ENTER/RETURN"
                            'Debug.Print("AFTER CHANGING EITHER LINE. WHEN FINISHED, "
                            'Debug.Print("PRESS [F5] TO CONTINUE."
                            Debug.Print("")
                            Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text))
                            'Debug.Print("gnQtyUnit = """, gnQtyUnit, """"
                            Debug.Print("gnQtyUnit = ""IN""")
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
                        Else
                            Debug.Print("NON-STANDARD MATERIAL FAMILY (" & gnMatlFam & ")")
                            Debug.Print("PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW")
                            Debug.Print("PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED")
                            Debug.Print("PRESS [F5] WHEN READY TO CONTINUE")
                            Debug.Print("")
                            Debug.Print("gnPartFam = """ & gnPartFam & """ 'PART FAMILY")
                            Debug.Print("gnQtyUnit = """ & gnQtyUnit & """ 'UNIT OF MEASURE")
                            Stop 'because we don't know WHAT to do with it
                            'but might NOT want to clear variables
                            'gnPartFam = ""
                            'gnQtyUnit = "" 'may want function here
                            ' UPDATE[2018.05.30]: As noted above
                            '     However, might need more handling here.
                        End If

                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), gnQtyUnit, ". IF OKAY, CONTINUE.")
                        Stop

                        Stop 'BKPT-2021-1105-1117
                        ' CHANGE NEEDED[2021.11.05]:
                        '     Property assignment needs moved
                        '     to collective assignment sequence
                        rt = dcAddProp(prRmQty, rt)
                        Debug.Print("") 'Landing line for debugging. Do not disable.
                    End If
                End If
            End With
            '    Else
            '        If 0 Then Stop 'and regroup
            '        ' Things are looking a right royal mess
            '        ' at the moment I'm writing this comment.
            '    End If
            'End If

            ' UPDATE[2021.11.10]
            '     Transported prRawMatl and prRmUnit assignments
            '     to point following this If/Else block
            '     to consolidate duplicate processes.
            ''--------------------------------------------'
            'Else 'for standard Part (NOT Sheet Metal) ---'
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
                If oDoc.ComponentDefinition.Document Is oDoc Then
                    'following needs indented if not already

                    ' [2018.07.31 by AT]
                    ' Added the following to try to
                    ' preselect non-sheet metal stock
                    '.dbFamily.Text = "D-BAR"
                    '.LbxFamily.Text = "D-BAR"
                    ' Doesn't quite do it.
                    'With New aiBoxData
                    ' bd = nuAiBoxData().UsingInches.UsingBox( _
                    'oDoc.ComponentDefinition.RangeBox(_
                    ')
                    Stop 'BKPT-2021-1105-0955
                    ' CHANGE NEEDED[2021.11.05]:
                    '     Probably want to move this
                    '     outside of this With block,
                    '     and closer to the beginning
                    '     of this function, as it could
                    '     prove helpful at other points.
                    bd = nuAiBoxData().UsingInches.SortingDims(
                oDoc.ComponentDefinition.RangeBox
            )
                    'End With

                    ck = .AskAbout(oDoc,
                "Please Select Stock for Machined Part" _
                & vbCrLf & vbCrLf & bd.Dump(0)
            )

                    If ck = vbYes Then
                        ' UPDATE[2018.05.30]:
                        '     Pulling some extraneous commented code
                        '     from here and beginning of block
                        With .ItemData
                            If .Exists(PnFamily) Then
                                If Len(gnPartFam) = 0 Then
                                    gnPartFam = .Item(PnFamily)
                                Else
                                    If gnPartFam = .Item(PnFamily) Then
                                    Else
                                        Debug.Print("=====")
                                        Debug.Print("Model Family differs from Genius")
                                        Debug.Print("Genius: " & PnFamily)
                                        Debug.Print("Model:  " & .Item(PnFamily))
                                        Debug.Print("gnPartFam = .Item(pnFamily) 'Press [ENTER] on this line to fix, and/or [F5] to continue'")
                                        Stop
                                    End If
                                End If
                                Debug.Print(PnFamily & "=" & gnPartFam)
                            End If

                            If .Exists(PnRawMaterial) Then
                                If Len(gnMatlNum) = 0 Then
                                    gnMatlNum = .Item(PnRawMaterial)
                                Else
                                    If gnMatlNum = .Item(PnRawMaterial) Then
                                    Else
                                        Debug.Print("=====")
                                        Debug.Print("Model Raw Material differs from Genius")
                                        Debug.Print("Genius: " & gnMatlNum)
                                        Debug.Print("Model:  " & .Item(PnRawMaterial))
                                        Debug.Print("gnMatlNum = .Item(pnRawMaterial) 'Press [ENTER] on this line to fix, and/or [F5] to continue'")
                                        Stop
                                    End If
                                End If
                                Debug.Print(PnRawMaterial & "=" & gnMatlNum)
                            End If
                        End With
                        If 0 Then Stop 'Use this for a debugging shim
                        ''  We're going to need something here
                        ''  to make sure raw material gets added
                        ''  for non sheet metal parts, as well
                        ''  What we're going to need to do
                        ''  is refactor this whole bloody thing.
                    Else
                        Stop 'shouldn't actually hit this line
                        'as the condition checked should always
                        'be true, at least for now.
                    End If
                Else
                    Stop 'because we've got a serious mismatch
                End If
            End With
            '
            '
            '

            If Len(gnMatlNum) > 0 Then 'and ONLY then
                'do we look for a Raw Material Family!

                ' This enclosing With block should NOT be necessary
                ' since the newFmTest1 above takes care of collecting
                ' the Stock Family along with the Stock itself
                With CnGnsDoyle().Execute(
                "select Family " &
                "from vgMfiItems " &
                "where Item='" & gnMatlNum & "';"
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
                            If Len(gnMatlFam) = 0 Then
                                gnMatlFam = .Item("Family").Value
                            ElseIf gnMatlFam = .Item("Family").Value Then
                            Else
                                Stop
                            End If
                        End With
                    End If
                End With

                If gnMatlFam = "DSHEET" Then
                    Stop 'because we should NOT be doing Sheet Metal in this section.
                    ' This might require further investigation and/or development, if encountered.
                    'We should be okay. This is sheet metal stock
                    ' UPDATE[2021.11.04]:
                    '     Expanding gnPartFam and gnQtyUnit
                    '     assignments to check for pre-
                    '     existing values, and validate
                    '     them if found.
                    If Len(gnPartFam) = 0 Then
                        gnPartFam = "D-RMT"
                    Else
                        If gnPartFam = "D-RMT" Then
                        Else
                            Stop 'because we have
                            'an unexpected situation
                        End If
                    End If

                    If Len(gnQtyUnit) = 0 Then
                        gnQtyUnit = "FT2"
                    Else
                        If gnQtyUnit = "FT2" Then
                        Else
                            Stop 'because we have
                            'an unexpected situation
                        End If
                    End If
                    ' UPDATE[2018.05.30]:
                    '     Moving part family assignment
                    '     to this section for better mapping
                    '     and updating to new Family names
                    '     as well as pulling up gnQtyUnit assignment
                ElseIf gnMatlFam = "D-BAR" Then
                    gnPartFam = "R-RMT"
                    If Len(gnQtyUnit) = 0 Then
                        'this might have to change
                        'to better handle case
                        'of missing prRmUnit
                        gnQtyUnit = prRmUnit.Text '"IN"
                    End If
                    ''may want function here
                    ' UPDATE[2018.05.30]: As noted above
                    '     Will keep Stop for now
                    '     pending further review,
                    '     hopefully soon
                    Debug.Print(aiPartNum & " [" & gnMatlNum & "]: " & CStr(aiPropsDesign.item(PnDesc).Text)) 'prRawMatl.Text
                    ' UPDATE[2021.03.11]: Replaced
                    ' aiPropsDesign.Item(pnPartNum)
                    ' as noted above
                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(gnMatlQty), gnQtyUnit, ". IF CHANGE NEEDED,") 'prRmQty.Text
                    Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                    Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                    'Debug.Print(oDoc.ComponentDefinition.RangeBox.MaxPoint.X - oDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Y - oDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Z - oDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
                    Debug.Print("")
                    Debug.Print("PLACE CURSOR ON gnQtyUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.")
                    Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.")
                    Debug.Print("")
                    Debug.Print("gnMatlQty = ", CStr(gnMatlQty)) 'prRmQty.Text
                    Debug.Print("gnQtyUnit = ""IN""")
                    Debug.Print("")
                    Stop 'because we might want a D-BAR handler
                    ' Actually, we might NOT need to stop here
                    ' if bar stock is already selected,
                    ' because quantities would presumably
                    ' have been established already.
                    ' Any D-BAR handler probably needs
                    ' to be implemented in prior section(s)
                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(gnMatlQty), gnQtyUnit, ". IF OKAY, CONTINUE.") 'prRmQty.Text
                    Stop
                    Stop 'BKPT-2021-1110-1647
                    ' CHANGE NEEDED[2021.11.10]
                    '     This Dictionary Property assignment
                    '     MUST be moved to the END of the function!
                    rt = dcWithProp(aiPropsUser, PnRmQty, gnMatlQty, rt) 'dcAddProp(prRmQty, rt)
                    Debug.Print("") 'Landing line for debugging. Do not disable.
                Else
                    Debug.Print("NON-STANDARD MATERIAL FAMILY (" & gnMatlFam & ")")
                    Debug.Print("PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW")
                    Debug.Print("PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED")
                    Debug.Print("PRESS [F5] WHEN READY TO CONTINUE")
                    Debug.Print("")
                    Debug.Print("gnPartFam = """ & gnPartFam & """ 'PART FAMILY")
                    Debug.Print("gnQtyUnit = """ & gnQtyUnit & """ 'UNIT OF MEASURE")
                    'gnPartFam = ""
                    'gnQtyUnit = "" 'may want function here
                    ' UPDATE[2018.05.30]: As noted above
                    '     However, might need more handling here.
                    Stop 'because we don't know WHAT to do with it
                End If
            Else
                If 0 Then Stop 'and regroup
                ' Things are looking a right royal mess
                ' at the moment I'm writing this comment.
            End If
        End If 'Sheetmetal vs Part

        'Stop 'BKPT-2021-1105-1011
        ' UPDATE[2021.11.10]
        '     Disabled this prRawMatl assignment pending removal.
        '     Counterpart moved below from sheet metal branch
        '     should serve in place of both branch instances.
        '     '
        '     Extraneous commentary removed.
        '     '
        'With prRawMatl
        '    If Len(Trim$(.Text)) > 0 Then
        '        If gnMatlNum <> .Text Then
        '            'Debug.Print("Raw Stock Selection"
        '            'Debug.Print("  Current : " & prRawMatl.Text
        '            'Debug.Print("  Proposed: " & gnMatlNum
        '            'Stop 'because we might not want to change existing stock ting
        '            'if
        '            ck = MsgBox( _
        '                Join({ _
        '                    "Raw Stock Change Suggested", _
        '                    "  for Item " & aiPartNum, _
        '                    "", _
        '                    "  Current : " & prRawMatl.Text, _
        '                    "  Proposed: " & gnMatlNum, _
        '                    "", "Change It?", "" _
        '                }, vbCrLf), _
        '                vbYesNo, aiPartNum & " Stock" _
        '            )
        '            '"Change Raw Material?"
        '            '"Suggested Sheet Metal"
        '            If ck = vbYes Then .Text = gnMatlNum
        '        End If
        '    Else
        '        .Text = gnMatlNum
        '    End If
        'End With
        ' rt = dcAddProp(prRawMatl, rt)

        'Stop 'BKPT-2021-1110-1130
        ' UPDATE[2021.11.10]
        '     Disabled this prRmUnit assignment pending removal.
        '     Duplicate moved below from sheet metal branch
        '     should serve in place of both branch instances.
        '     '
        '     Also moved End If AHEAD of this block to minimize
        '     comment clutter WITHIN branches.
        'With prRmUnit
        '    If Len(.Text) > 0 Then
        '        If Len(gnQtyUnit) > 0 Then
        '            If .Text <> gnQtyUnit Then
        '                Stop 'and check both so we DON'T
        '                'automatically "fix" the RMUNIT value
        '
        '                .Text = gnQtyUnit
        '
        '                If 0 Then Stop 'Ctrl-9 here to skip changing
        '            End If
        '        End If
        '    Else 'we're ting a new quantity unit
        '        .Text = gnQtyUnit
        '    End If
        'End With
        ' rt = dcAddProp(prRmUnit, rt)

        Stop 'BKPT-2021-1109-1610
        ' UPDATE[2021.11.10]
        '     Transported this prRawMatl assignment
        '     from sheet metal branch to consolidate
        '     both instances of duplicated process
        '     into one following both branches.
        '     '
        '     Extraneous commentary also removed.
        '     '
        If prRawMatl Is Nothing Then
            rt = dcWithProp(aiPropsUser, PnRawMaterial, gnMatlNum, rt)
            Debug.Print("") 'Breakpoint Landing
        Else
            With prRawMatl
                If Len(Trim$(.Text)) > 0 Then
                    If gnMatlNum <> .Text Then
                        'Debug.Print("Raw Stock Selection"
                        'Debug.Print("  Current : " & prRawMatl.Text
                        'Debug.Print("  Proposed: " & gnMatlNum
                        'Stop 'because we might not want to change existing stock ting
                        'if
                        ck = MsgBox(
                    Join({
                        "Raw Stock Change Suggested",
                        "  for Item " & aiPartNum,
                        "",
                        "  Current : " & prRawMatl.Text,
                        "  Proposed: " & gnMatlNum,
                        "", "Change It?", ""
                    }, vbCrLf),
                    vbYesNo, aiPartNum & " Stock"
                )
                        '"Change Raw Material?"
                        '"Suggested Sheet Metal"
                        If ck = vbYes Then .Text = gnMatlNum
                    End If
                Else
                    .Text = gnMatlNum
                End If
            End With
            rt = dcAddProp(prRawMatl, rt)
        End If

        'Stop 'BKPT-2021-1110-1133
        ' UPDATE[2021.11.10]
        '     Transported this prRmUnit assignment
        '     from sheet metal branch to consolidate
        '     both instances of duplicated process
        '     into one following both sheet metal
        '     and structural branches
        If prRmUnit Is Nothing Then
            rt = dcWithProp(aiPropsUser, PnRmUnit, gnQtyUnit, rt)
            Debug.Print("") 'Breakpoint Landing
        Else
            With prRmUnit
                If Len(.Text) > 0 Then
                    If Len(gnQtyUnit) > 0 Then
                        If .Text <> gnQtyUnit Then
                            Stop 'and check both so we DON'T
                            'automatically "fix" the RMUNIT value

                            .Text = gnQtyUnit

                            If 0 Then Stop 'Ctrl-9 here to skip changing
                        End If
                    End If
                Else 'we're ting a new quantity unit
                    .Text = gnQtyUnit
                End If
            End With
            rt = dcAddProp(prRmUnit, rt)
        End If
        ' rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt) 'gnQtyUnit WAS "FT2"
        ' Plan to remove commented line above,
        ' superceded by the one above that
        Debug.Print("") 'Breakpoint Landing


        'Stop 'BKPT-2021-1110-1133
        ' UPDATE[2021.11.09]
        '     This is a VERY crude implementation
        '     of the closing BOM Structure assignment.
        '     Plan on revision and cleanup in future.
        If gnBomType = aiBomType Then
            With oDoc.ComponentDefinition
                If .BOMStructure <> gnBomType Then
                    On Error Resume Next
                    .BOMStructure = gnBomType
                    If Err.Number = 0 Then
                        '        aiBomType = .BOMStructure
                    Else
                        Stop
                        Debug.Print("") 'Breakpoint Landing
                        '        aiBomType = BOMStructureEnum.kPurchasedBOMStructure
                        '
                        '        ' WARNING: NOT a good way to go about this
                        '        '     but will go with it for now
                    End If
                    On Error GoTo 0
                    Stop
                    Debug.Print("") 'Breakpoint Landing
                End If
                Debug.Print("") 'Breakpoint Landing
            End With
        Else
            Stop
            Debug.Print("") 'Breakpoint Landing
        End If
        'this With block was pulled down
        'from the BOMStructure section above.
        'It MIGHT want to be moved ahead
        'of the two preceding With blocks.
        'With .ComponentDefinition 'appears to be extraneous now.
        ' Disabling it does not lead to compilation errors, indicating
        ' nothing (active) within the block depends on it anymore.
        ' It does look like this disabled If block
        ' would still need it, though.
        ' This If block is meant to  BOMStructure
        ' according to information gathered from the
        ' Model, its Vault location, Genius, and if
        ' all else fails, the Users themselves.
        '
        ' Since it S a Model attribute, it belongs
        ' toward the bottom, with the blocks that 
        ' Model Properties. Plan to move it there.
        '
        'If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
        '    On Error Resume Next
        '    .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
        '    If Err.Number = 0 Then
        '        aiBomType = .BOMStructure
        '    Else
        '        aiBomType = BOMStructureEnum.kPurchasedBOMStructure
        '
        '        ' WARNING: NOT a good way to go about this
        '        '     but will go with it for now
        '    End If
        '    On Error GoTo 0
        'Else
        '    aiBomType = .BOMStructure 'to make sure this is captured
        'End If
        '    'End With
        'ElseIf aiBomType = BOMStructureEnum.kPurchasedBOMStructure Then
        '    ' As mentioned above, gnPartFam
        '    ' SHOULD be  at this point
        '    If Len(gnPartFam) = 0 Then
        '        If 1 Then Stop 'because we might
        '        'need to check out the situation
        '        gnPartFam = "D-PTS" 'by default
        '    End If
        'Else
        '    Stop 'because we might need
        '    'to do something else
        '    'based on an unexpected
        '    'BOM Structure
        'End If

        'Stop 'BKPT-2021-1105-1020
        ' CHANGE NEEDED[2021.11.05]:
        '     Family assignment should be
        '     ported up into collective
        '     Property assignment, although
        '     its position here assures
        '     one instance of the sequence
        '     catches ALL divergent cases
        '     leading up to this point.
        '     '
        '     Ultimately, those cases probably
        '     need to be consolidated HERE
        '     if, or WHEN possible.
        '     '
        ' Get the design tracking property ,
        ' and update the Cost Center Property
        If oDoc.ComponentDefinition.IsContentMember Then
            ' Don't muck around with the Family!
        Else
            If Len(gnPartFam) > 0 Then
                dcVlGn.Item("Family") = gnPartFam
                If aiPartFam = gnPartFam Then
                Else
                    On Error Resume Next
                    prFamily.Text = gnPartFam
                    If Err.Number Then
                        Debug.Print("CHGFAIL[FAMILY]{'" _
                        & prFamily.Text & "' -> '" & gnPartFam & "'}: " _
                        & oDoc.DisplayName & " (" & oDoc.FullDocumentName & ")")
                        If MsgBox(
                        "Couldn't Change Family" & vbCrLf _
                        & "for Item " & oDoc.DisplayName & vbCrLf _
                        & vbCrLf & "(" & oDoc.FullDocumentName & ")" _
                        & vbCrLf & vbCrLf & "Stop to Review?",
                        vbYesNo Or vbDefaultButton2,
                        oDoc.DisplayName
                    ) = vbYes Then
                            Stop
                        End If
                    Else
                    End If
                    On Error GoTo 0
                End If
                rt = dcAddProp(prFamily, rt)
                Debug.Print("") 'Breakpoint Landing
                ' rt = dcWithProp(aiPropsDesign, pnFamily, gnPartFam, rt)
            End If
        End If
        'End With
        ' UPDATE[2021.11.09]
        '     Moved Part Mass Property assignment
        '     out of the main With block, modified
        '     to take its value from the new Values
        '     Dictionary.
        'rt = dcWithProp(
        '    aiPropsUser, pnMass,
        '    dcVlPr.Item(pnMass), rt
        ') 'System.Math.Round(cvMassKg2LbM * .Mass, 4)

        '    Call iSyncPartFactory(oDoc) 'Backport Properties to iPart Factory
        '    dcGeniusPropsPartRev20180530_broken2 = rt
        'End If
    End Function

    Public Function dcGeniusPropsPartRev20180530_broken(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim dcChg As Scripting.Dictionary
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

        Dim txFilePath As String

        If dc Is Nothing Then
            dcGeniusPropsPartRev20180530_broken =
        dcGeniusPropsPartRev20180530_broken(
            oDoc, New Scripting.Dictionary
        )
        Else
            rt = dc
            dcChg = New Scripting.Dictionary

            With oDoc
                txFilePath = .FullFileName

                ' Get Property s
                With .Propertys
                    aiPropsUser = .Item(GnCustom)
                    aiPropsDesign = .Item(GnDesign)
                End With

                ' Get Custom Properties
                prRawMatl = aiGetProp(aiPropsUser, PnRawMaterial, 1)
                prRmUnit = aiGetProp(aiPropsUser, PnRmUnit, 1)
                prRmQty = aiGetProp(aiPropsUser, PnRmQty, 1)

                ' Part Number and Family properties
                prPartNum = aiGetProp(aiPropsDesign, PnPartNum) 'ADDED 2021.03.11
                pnModel = prPartNum.Text
                prFamily = aiGetProp(aiPropsDesign, PnFamily)

                ' UPDATE[2018.02.06]: Using new UserForm, see below
                With .ComponentDefinition
                    ' Request #1: Get the Mass in Pounds
                    With .MassProperties
                        rt = dcWithProp(
                        aiPropsUser, PnMass,
                        System.Math.Round(CvMassKg2LbM * .Mass, 4), rt
                    )
                    End With

                    bomStruct = .BOMStructure 'BOMStructureEnum.kDefaultBOMStructure '''''
                    dcChg = d2g1f1(prFamily, dcChg)
                End With
                ' At this point, nmFamily SHOULD be 

                'Request #4: Change Cost Center iProperty.
                If bomStruct = BOMStructureEnum.kNormalBOMStructure Then
                    '----------------------------------------------------'
                    If .SubType = GuidSheetMetal Then 'for SheetMetal ---'
                        '----------------------------------------------------'
                        ' NOTE[2018-05-31]: At this point, we MAY wish
                        'Request #3: Get sheet metal extent area
                        rt = dcFlatPatProps(.ComponentDefinition, rt)
                        ' NOTE[2018-05-30]: Raw Material Quantity value

                        'NOTE: THIS call might best be combined somehow
                        If prRawMatl Is Nothing Then
                            If rt.Exists("OFFTHK") Then
                                ' UPDATE[2018.05.30]: Restoring original key check
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
                            ''  ACTION ADVISED[2018.09.14]: pnStock can probably be 
                            If Len(prRawMatl.Text) > 0 Then
                                pnStock = prRawMatl.Text
                                ' This With block copied and modified [2021.03.11]
                                With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                    If .BOF Or .EOF Then
                                        'Stop 'because Material value likely invalid
                                        pnStock = ptNumShtMetal(oDoc.ComponentDefinition)
                                        Debug.Print("") 'Breakpoint Landing
                                        ''  ACTION TAKEN[2021.03.11]: temporary measure to try to ensure
                                    Else
                                        ''  This section retained from source,
                                        ''  With .Fields
                                        ''      mtFamily = .Item("Family").Value
                                        ''  End With
                                    End If
                                End With
                                '
                                ' This section likely should be removed when primary issue resolved.
                                '
                            Else
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If

                            If Len(pnStock) = 0 Then
                                ' UPDATE[2018.05.30]: Pulling ALL code/text from this section
                                With newFmTest1()
                                    If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop

                                    bd = nuAiBoxData().UsingInches.SortingDims(
                                    oDoc.ComponentDefinition.RangeBox
                                )
                                    ck = .AskAbout(oDoc,
                                    "No Stock Found! Please Review" _
                                    & vbCrLf & vbCrLf & bd.Dump(0)
                                )

                                    If ck = vbYes Then
                                        ' UPDATE[2018.05.30]: Pulling some extraneous commented code
                                        With .ItemData
                                            If .Exists(PnFamily) Then
                                                nmFamily = .Item(PnFamily)
                                                Debug.Print(PnFamily & "=" & nmFamily)
                                            End If

                                            If .Exists(PnRawMaterial) Then
                                                pnStock = .Item(PnRawMaterial)
                                                Debug.Print(PnRawMaterial & "=" & pnStock)
                                            End If
                                        End With
                                        If 0 Then Stop 'Use this for a debugging shim
                                    End If
                                End With
                            ElseIf Left$(pnStock, 2) = "LG" Then 'it's probably lagging
                                Debug.Print(pnModel & ": PROBABLE LAGGING")
                                Debug.Print("  TRY TO IDENTIFY, AND FILL IN BELOW.")
                                Debug.Print("  PRESS ENTER ON pnStock LINE WHEN")
                                Debug.Print("  COMPLETED, THEN F5 TO CONTINUE.")
                                Debug.Print("  pnStock = """ & pnStock & """")
                                Stop
                            End If

                            If Len(pnStock) > 0 Then 'and ONLY then do we look for a Raw Material Family!

                                With CnGnsDoyle().Execute(
                                "select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                    If .BOF Or .EOF Then
                                        Stop 'because Material value likely invalid
                                        ''  ACTION ADVISED[2018.09.14]: Will need to address this situation
                                    Else
                                        With .Fields
                                            mtFamily = .Item("Family").Value
                                        End With

                                        ' UPDATE[2021.06.18]: New pre-check for Material Item
                                        If mtFamily Like "?-MT*" Then
                                            Debug.Print(pnModel & "[" & prRmQty.Text & qtUnit & "*" & pnStock & ": " & aiPropsDesign.item(PnDesc).Text & "]") ' prRawMatl.Text
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
                                            ' UPDATE[2018.05.30]: Moving part family assignment
                                        ElseIf mtFamily = "D-BAR" Then
                                            ' UPDATE[2021.06.18]: Added check for Part Family already 
                                            If Len(nmFamily) = 0 Then
                                                nmFamily = "R-RMT"
                                            Else
                                                Debug.Print("") 'Breakpoint Landing
                                                'Stop
                                            End If

                                            qtUnit = prRmUnit.Text '"IN"
                                            ''may want function here
                                            ' UPDATE[2018.05.30]: As noted above
                                            Debug.Print(pnModel & " [" & prRawMatl.Text & "]: " & aiPropsDesign.item(PnDesc).Text)
                                            ' UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,")
                                            Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                                            Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                            With oDoc.ComponentDefinition.RangeBox
                                                '    Debug.Print("")
                                                '(.MaxPoint.X - .MinPoint.X) / 2.54, _
                                                '(.MaxPoint.Y - .MinPoint.Y) / 2.54, _
                                                '(.MaxPoint.Z - .MinPoint.Z) / 2.54
                                            End With
                                            'Debug.Print("CURRENT RAW MATERIAL QUANTITY (";
                                            'Debug.Print(CStr(prRmQty.Text), ") IS SHOWN BELOW."
                                            Debug.Print("")
                                            Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text))
                                            'Debug.Print("qtUnit = """, qtUnit, """"
                                            Debug.Print("qtUnit = ""IN""")
                                            'Debug.Print(""
                                            Stop 'because we might want a D-BAR handler
                                            ' Actually, we might NOT need to stop here
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                            Stop
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
                                        Stop 'and check both so we DON'T automatically "fix" the RMUNIT value

                                        .Text = qtUnit

                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        Debug.Print("") 'Another landing line

                        '--------------------------------------------'
                    Else 'for standard Part (NOT Sheet Metal) ---'
                        '--------------------------------------------'
                        ' [2018.07.31 by AT] Duped following block from above
                        With newFmTest1()
                            If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop

                            ' [2018.07.31 by AT] Added the following to try to
                            bd = nuAiBoxData().UsingInches.SortingDims(
                                    oDoc.ComponentDefinition.RangeBox
                                )

                            ck = .AskAbout(oDoc,
                                    "Please Select Stock for Machined Part" _
                                    & vbCrLf & vbCrLf & bd.Dump(0)
                                )

                            If ck = vbYes Then
                                ' UPDATE[2018.05.30]: Pulling some extraneous commented code
                                With .ItemData
                                    If .Exists(PnFamily) Then
                                        nmFamily = .Item(PnFamily)
                                        Debug.Print(PnFamily & "=" & nmFamily)
                                    End If

                                    If .Exists(PnRawMaterial) Then
                                        pnStock = .Item(PnRawMaterial)
                                        Debug.Print(PnRawMaterial & "=" & pnStock)
                                    End If
                                End With
                                If 0 Then Stop 'Use this for a debugging shim
                                ''  We're going to need something here
                            End If
                        End With
                        '
                        '
                        '
                        '
                        ' The following If block is copied wholesale from sheet metal section above.
                        If Len(pnStock) > 0 Then 'and ONLY then do we look for a Raw Material Family!

                            ' This enclosing With block should NOT be necessary
                            With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]: Will need to address this situation
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    '
                                    ' Content formerly here moved BELOW and OUT of this section
                                End If
                            End With
                            ' These closing statements moved up from below following If block
                            '

                            If mtFamily = "DSHEET" Then
                                Stop
                                'because we should NOT be doing Sheet Metal in this section.
                                nmFamily = "D-RMT"
                                qtUnit = "FT2"
                                ' UPDATE[2018.05.30]: Moving part family assignment
                            ElseIf mtFamily = "D-BAR" Then
                                nmFamily = "R-RMT"
                                qtUnit = prRmUnit.Text '"IN"
                                ''may want function here
                                ' UPDATE[2018.05.30]: As noted above Will keep Stop for now
                                Debug.Print(pnModel, " [", prRawMatl.Text, "]: ", aiPropsDesign.item(PnDesc).Text)
                                ' UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,")
                                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                                Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                '   Debug.Print(oDoc.ComponentDefinition.RangeBox.MaxPoint.X - oDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Y - oDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Z - oDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
                                Debug.Print("")
                                Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.")
                                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.")
                                Debug.Print("")
                                Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text))
                                Debug.Print("qtUnit = ""IN""")
                                Debug.Print("")
                                Stop 'because we might want a D-BAR handler
                                ' Actually, we might NOT need to stop here
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                Stop
                                rt = dcAddProp(prRmQty, rt)
                                Debug.Print("") 'Landing line for debugging. Do not disable.
                            Else
                                nmFamily = ""
                                qtUnit = "" 'may want function here
                                ' UPDATE[2018.05.30]: As noted above
                                Stop 'because we don't know WHAT to do with it
                            End If
                        Else
                            If 0 Then Stop 'and regroup
                            ' Things are looking a right royal mess
                            ' at the moment I'm writing this comment.
                        End If

                        With prRawMatl
                            If Len(Trim$(.Text)) > 0 Then
                                If pnStock <> .Text Then
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
                                        Stop 'and check both so we DON'T automatically "fix" the RMUNIT value

                                        .Text = qtUnit

                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                    End If 'Sheetmetal vs Part
                ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                    ' As mentioned above, nmFamily SHOULD be  at this point
                    If Len(nmFamily) = 0 Then
                        If 1 Then Stop 'because we might need to check out the situation
                        nmFamily = "D-PTS" 'by default
                    End If
                Else
                    Stop 'because we might need to do something else
                End If

                ' Get the design tracking property ,
                If oDoc.ComponentDefinition.IsContentMember Then
                    ' Don't muck around with the Family!
                Else
                    If Len(nmFamily) > 0 Then
                        On Error Resume Next
                        prFamily.Text = nmFamily
                        If Err.Number Then
                            Debug.Print("CHGFAIL[FAMILY]{'" _
                            & prFamily.Text & "' -> '" & nmFamily & "'}: " _
                            & oDoc.DisplayName & " (" & oDoc.FullDocumentName & ")")
                            If MsgBox(
                            "Couldn't Change Family" & vbCrLf _
                            & "for Item " & oDoc.DisplayName & vbCrLf _
                            & vbCrLf & "(" & oDoc.FullDocumentName & ")" _
                            & vbCrLf & vbCrLf & "Stop to Review?",
                            vbYesNo Or vbDefaultButton2,
                            oDoc.DisplayName
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

            Call iSyncPartFactory(oDoc) 'Backport Properties to iPart Factory
            dcGeniusPropsPartRev20180530_broken = rt
        End If
    End Function

    Public Function dcGeniusPropsPartDvl20210929(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
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
            dcGeniusPropsPartDvl20210929 =
        dcGeniusPropsPartDvl20210929(
            oDoc, New Scripting.Dictionary
        )
        Else
            rt = dc

            With oDoc
                ' Get Property s
                With .Propertys
                    aiPropsUser = .Item(GnCustom)
                    aiPropsDesign = .Item(GnDesign)
                End With

                ' Get Custom Properties
                prRawMatl = aiGetProp(aiPropsUser, PnRawMaterial, 1)
                prRmUnit = aiGetProp(aiPropsUser, PnRmUnit, 1)
                prRmQty = aiGetProp(aiPropsUser, PnRmQty, 1)

                ' Part Number and Family properties
                prPartNum = aiGetProp(aiPropsDesign, PnPartNum) 'ADDED 2021.03.11
                pnModel = prPartNum.Text
                prFamily = aiGetProp(aiPropsDesign, PnFamily)

                ' Request #1: Get the Mass in Pounds
                With .ComponentDefinition.MassProperties
                    rt = dcWithProp(aiPropsUser, PnMass, System.Math.Round(CvMassKg2LbM * .Mass, 4), rt)
                End With

                ' NOTE[2021.10.01]: This block is for Purchased Part Determination! (see below)
                ' UPDATE[2018.02.06]: Using new UserForm, see below
                With .ComponentDefinition
                    ' Get BOM Structure type, correcting if appropriate,
                    ck = vbNo
                    ' UPDATE[2018.05.31]: Combined both InStr checks
                    ' look at N++ tab "new 7" for content here
                End With
                ' NOTE[2021.10.01]: END OF BLOCK for Purchased Part Determination!

                'Request #4: Change Cost Center iProperty.
                If bomStruct = BOMStructureEnum.kNormalBOMStructure Then
                    ' look at N++ tab "new 7" for content here
                ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                    ' look at N++ tab "new 7" for content here
                Else
                    ' look at N++ tab "new 7" for content here
                End If

                ' Get the design tracking property ,
                If oDoc.ComponentDefinition.IsContentMember Then
                    ' look at N++ tab "new 7" for content here
                Else
                    ' look at N++ tab "new 7" for content here
                End If
            End With
        End If
    End Function

    Public Function d2g2f1(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary 'Inventor.BOMStructureEnum
        '
        ' d2g2f1 -- (to be determined)
        '
        ' code here extracted for development
        ' from Function dcGeniusPropsPartRev20180530
        '     in module modGPUpdateAT (start line 559)
        '     lines 1056 (497 down from start)
        '        to 1201 (146 lines copied)
        ' along with necessary declarations:
        Dim pnModel As String
        Dim nmFamily As String
        Dim mtFamily As String
        Dim pnStock As String
        Dim qtUnit As String
        Dim bd As AiBoxData
        Dim ck As VbMsgBoxResult
        '
        ' followed by new declarations:
        Dim rt As Scripting.Dictionary

        rt = New Scripting.Dictionary

        With oDoc.ComponentDefinition
            If .Document Is oDoc Then
                bd = nuAiBoxData(
            ).UsingInches.SortingDims(
                .RangeBox
            )
                With newFmTest1() '== Original Line 1056 ==
                    'If Not (oDoc.ComponentDefinition.Document Is oDoc) Then Stop
                    'moved this check outside this form block (see above)

                    ' [2018.07.31 by AT]
                    ' Added the following to try to
                    ' preselect non-sheet metal stock
                    '.dbFamily.Text = "D-BAR"
                    '.LbxFamily.Text = "D-BAR"
                    ' Doesn't quite do it.

                    ck = .AskAbout(oDoc,
                    "Please Select Stock for Machined Part" _
                    & vbCrLf & vbCrLf & bd.Dump(0)
                )

                    If ck = vbYes Then
                        ' UPDATE[2018.05.30]:
                        '     Pulling some extraneous commented code
                        '     from here and beginning of block
                        With .ItemData()
                            If .Exists(PnFamily) Then
                                nmFamily = .Item(PnFamily)
                                rt.Add(PnFamily, nmFamily)
                                Debug.Print(PnFamily & "=" & nmFamily)
                            End If

                            If .Exists(PnRawMaterial) Then
                                pnStock = .Item(PnRawMaterial)
                                rt.Add(PnRawMaterial, pnStock)
                                Debug.Print(PnRawMaterial & "=" & pnStock)
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
            Else
                Stop
            End If
        End With

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
                End If
            End With

            If mtFamily = "DSHEET" Then
                Stop 'because we should NOT be doing Sheet Metal in this section.
                'might require further investigation and/or development, if encountered.
                'We should be okay. This is sheet metal stock
                nmFamily = "D-RMT"
                qtUnit = "FT2"
                ' UPDATE[2018.05.30]:
                '     Moving part family assignment
                '     to this section for better mapping
                '     and updating to new Family names
                '     as well as pulling up qtUnit assignment
            ElseIf mtFamily = "D-BAR" Then
                nmFamily = "R-RMT"
                Stop 'and note disabled qtUnit -- needs work here
                'qtUnit = prRmUnit.Text '"IN"
                ''may want function here
                ' UPDATE[2018.05.30]: As noted above
                '     Will keep Stop for now
                '     pending further review,
                '     hopefully soon
                Stop 'and note disabled prRawMatl too
                'Debug.Print(pnModel, " [", prRawMatl.Text, "]: ", aiPropsDesign(pnDesc).Text
                ' UPDATE[2021.03.11]: Replaced
                ' aiPropsDesign.Item(pnPartNum)
                ' as noted above
                Stop 'and note disabled prRmQty
                'Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,"
                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                '  Debug.Print(oDoc.ComponentDefinition.RangeBox.MaxPoint.X - oDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Y - oDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Z - oDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
                Debug.Print("")
                Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.")
                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.")
                Debug.Print("")
                Stop 'and note disabled prRmQty again
                'Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text)
                Debug.Print("qtUnit = ""IN""")
                Debug.Print("")
                Stop 'because we might want a D-BAR handler
                ' Actually, we might NOT need to stop here
                ' if bar stock is already selected,
                ' because quantities would presumably
                ' have been established already.
                ' Any D-BAR handler probably needs
                ' to be implemented in prior section(s)
                Stop 'and note one moredisabled prRmQty
                'Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE."
                Stop
                Stop 'and note prRmQty once more disabled
                'this one really DOES need removed from this function
                ' rt = dcAddProp(prRmQty, rt)
                Debug.Print("") 'Landing line for debugging. Do not disable.
            Else
                Stop 'because we don't know WHAT to do with it
                nmFamily = ""
                qtUnit = "" 'may want function here
                ' UPDATE[2018.05.30]: As noted above
                '     However, might need more handling here.
            End If

        Else
            If 0 Then Stop 'and regroup
            ' Things are looking a right royal mess
            ' at the moment I'm writing this comment.
        End If

        d2g2f1 = rt
        ''Debug.Print(ConvertToJson(d2g2f1(aiDocPart(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences.ItemByName("04-18-102-1006:1").Definition.Document)), "  ")
    End Function

    Public Function d2g1f1(
    prFamily As Inventor.Property,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary 'Inventor.BOMStructureEnum
        Dim rt As Scripting.Dictionary
        'Dim oDoc As Inventor.PartDocument
        Dim ck As VbMsgBoxResult
        Dim txFilePath As String
        Dim isNow As Inventor.BOMStructureEnum
        Dim shdBe As Inventor.BOMStructureEnum
        Dim gnFam As String
        Dim aiFam As String
        Dim ptNum As String
        Dim Fm As New FmTest2

        If dc Is Nothing Then
            rt = d2g1f1(prFamily, New Scripting.Dictionary)
        Else
            rt = dc

            With prFamily
                ' Get Family from Model
                aiFam = .Text

                With .Parent 'Property 
                    ' Get Part Number from Model
                    ptNum = .Item(PnPartNum).Text

                    ' Then try to get Family from Genius
                    With CnGnsDoyle()
                        With .Execute(Join({
                        "select ISNULL(i.Family, '') Family",
                        "from vgMfiItems i right join",
                        "(values ('" & ptNum & "')) ls(Item)",
                        "on i.Item = ls.Item",
                        ";"
                    }, vbCrLf))
                            If .BOF Or .EOF Then
                                Stop 'because something went wrong
                                gnFam = ""
                            Else
                                gnFam = .GetRows()(0, 0)
                            End If
                            .Close()
                        End With
                    End With

                    With .Parent ' OF Property s
                        ' Get File Path to check for Purchased Part
                        txFilePath = aiDocument(.Parent).FullFileName

                        'Request #2: Change Cost Center iProperty.
                        If ck4ContentMember(.Parent) Then
                            If Len(gnFam) = 0 Then
                                gnFam = "D-HDWR"
                            ElseIf gnFam = "D-HDWR" Then
                            ElseIf gnFam = "D-PTS" Then
                            ElseIf gnFam = "R-PTS" Then
                            Else
                                Stop
                            End If
                        End If

                        isNow = bomStructOf(.Parent)

                        ' Fm = newFmTest2()

                        ' Check Model Family against Genius Family,
                        ' if defined, and if different, ask whether
                        ' it should be changed.
                        If Len(gnFam) > 0 Then
                            If gnFam <> aiFam Then
                                ck = Fm.AskAbout(.Parent,
                                Join({
                                    "Model Family " & aiFam & " does not",
                                    "match Genius Part Family " & gnFam
                                }, vbCrLf),
                                Join({
                                    "Should Model be updated",
                                    "to match Genius?"
                                }, vbCrLf)
                            )
                                If ck = vbCancel Then
                                    Stop
                                End If

                                If ck = vbYes Then
                                    rt.Add(prFamily.Name, gnFam)
                                Else
                                    gnFam = aiFam
                                    'going to need final family below
                                    'and Genius Family will be changed
                                    'anyway, if Model Family isn't
                                End If
                            End If
                        End If

                        ' Get BOM Structure type,
                        ' correcting if appropriate,
                        ' UPDATE[2018.05.31]: Combined both InStr checks
                        If InStr(1, txFilePath,
                        "\Doyle_Vault\Designs\purchased\"
                    ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                        "|" & prFamily.Text & "|"
                    ) > 0 Then
                            shdBe = BOMStructureEnum.kPurchasedBOMStructure
                        Else
                            shdBe = isNow
                        End If

                        If shdBe <> isNow Then
                            ck = Fm.AskAbout(.Parent,
                            Join({
                                "Model Family " & gnFam & " or File Path",
                                "(" & txFilePath & ")",
                                "indicates a Purchased Part, but BOM",
                                "Structure is NOT  to match"
                            }, vbCrLf),
                            Join({
                                "Should BOM Structure",
                                "be  to Purchased?"
                            }, vbCrLf)
                        )
                            If ck = vbCancel Then
                                Stop
                            End If

                            If ck = vbYes Then
                                'On Error Resume Next
                                '.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                                'If Err.Number = 0 Then
                                '    bomStruct = .BOMStructure
                                'Else
                                '    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                                '    ' WARNING: NOT a good way to go about this
                                '    '     but will go with it for now
                                'End If
                                'On Error GoTo 0
                                rt.Add("BOMstructure", shdBe)
                            Else
                                shdBe = isNow
                                'want to know what the USER says
                                'it should be, and BOM structure
                                'MIGHT affect Genius Part status
                                'in subsequent import operation.

                                'bomStruct = .BOMStructure 'to make sure this is captured
                            End If
                        End If
                    End With
                End With
                Debug.Print("") 'Breakpoint Landing
            End With
        End If

        d2g1f1 = rt
        'Debug.Print(txDumpLs(d2g1f1(aiDocActive().Propertys(gnDesign).Item(pnFamily)).Keys)
    End Function

    Public Function dcCtOfEach(ls As Object) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ck As Object
        Dim mx As Long
        Dim dx As Long

        rt = New Scripting.Dictionary
        If TypeOf ls Is Array Then
            mx = UBound(ls)
            dx = LBound(ls)
            With rt
                Do Until dx > mx
                    ck = ls(dx)
                    If .Exists(ck) Then
                        .Item(ck) =
                    .Item(ck) + 1
                    Else
                        .Add(ck, 1)
                    End If

                    dx = 1 + dx
                Loop
            End With

            dcCtOfEach = rt
        Else
            dcCtOfEach =
            dcCtOfEach({ls})
        End If
    End Function

    Public Function dcGnsMatlOps(ThisApplication As Inventor.Application,
    DimCt As Scripting.Dictionary,
    Optional MtSpec As String = ""
) As Scripting.Dictionary 'defaulted to SS, but maybe not such a great idea
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim BillRow As Scripting.Dictionary
        Dim rs As ADODB.Record
        Dim ky As Object

        rt = New Scripting.Dictionary
        With CnGnsDoyle()
            On Error Resume Next

            Err.Clear()
            rs = .Execute(
        SqlOf_GnsMatlOptions(ThisApplication,
            MtSpec, DimCt.Keys
        ))

            If Err.Number = 0 Then
                With DcFromAdoRS(rs, "")
                    For Each ky In .Keys
                        BillRow = dcOb(.Item(ky))
                        If BillRow Is Nothing Then
                            Stop
                        Else
                            rt.Add(BillRow.Item("Item"), BillRow)
                        End If
                    Next : End With

                rs.Close()
            Else
                Stop
                Err.Clear()
            End If

            On Error GoTo 0
            .Close
        End With
        dcGnsMatlOps = rt
    End Function

    Public Function ck4ContentMember(
    AiDoc As Inventor.Document
) As Boolean
        ck4ContentMember =
    ptIsContentMember(
    aiDocPart(AiDoc))
    End Function

    Public Function ptIsContentMember(
    AiDoc As Inventor.PartDocument
) As Boolean
        If AiDoc Is Nothing Then
            ptIsContentMember = 0
        Else
            ptIsContentMember = AiDoc.ComponentDefinition.IsContentMember
        End If
    End Function

    Public Function bomStructOfPart(
    AiDoc As Inventor.PartDocument
) As Inventor.BOMStructureEnum
        If AiDoc Is Nothing Then
            bomStructOfPart = 0
        Else
            bomStructOfPart = AiDoc.ComponentDefinition.BOMStructure
        End If
    End Function

    Public Function bomStructOfAssy(
    AiDoc As Inventor.AssemblyDocument
) As Inventor.BOMStructureEnum
        If AiDoc Is Nothing Then
            bomStructOfAssy = 0
        Else
            bomStructOfAssy = AiDoc.ComponentDefinition.BOMStructure
        End If
    End Function

    Public Function bomStructOf(
    AiDoc As Inventor.Document
) As Inventor.BOMStructureEnum
        If AiDoc Is Nothing Then
            bomStructOf = 0
        ElseIf TypeOf AiDoc Is Inventor.PartDocument Then
            bomStructOf = bomStructOfPart(AiDoc)
        ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
            bomStructOf = bomStructOfAssy(AiDoc)
        Else
            bomStructOf = 0
        End If
    End Function
End Module