Module Module1
    Dim ThisApplication As Inventor.Application
    Public Function dcCutTimePerimeter(ad As Inventor.Document, ThisApplication As Inventor.Application,
                                       Optional dc As Scripting.Dictionary = Nothing,
                                       Optional incTop As Long = 0) As Scripting.Dictionary
        Dim rt As New Scripting.Dictionary
        Dim oDoc As Inventor.Document
        Dim ky As Object

        For Each ky In dcAiDocComponents(ad, If(dc Is Nothing, New Scripting.Dictionary, dc), incTop).Keys
            oDoc = aiDocument(dcAiDocComponents(ad, dc, incTop).Item(ky))
            ' Add perimeter calculation or other logic here as needed
            ' Example: rt.Add oDoc.Propertys.Item(GnDesign).Item(PnPartNum).Text, fpPerimeterInch(ThisApplication, oDoc)
        Next

        Return rt
    End Function
    'Debug.Print(dumpLsKeyVal(dcCutTimePerimeter(ThisApplication.ActiveDocument))

    Public Function mdl1g0f0(ThisApplication As Inventor.Application) As Long
        Dim dc As Scripting.Dictionary = dcAssyDocComponents(ThisApplication.Documents.ItemByName(
            "C:\Doyle_Vault\Designs\Misc\andrewT\02\02-weldmentStd-01.iam"))
        Dim ky As Object
        Dim ad As Inventor.Document
        Dim InvProperty As Inventor.Property

        For Each ky In dc.Keys
            ad = aiDocument(dc.Item(ky))
            With dcGeniusProps(ad)
                If .Exists(PnRawMaterial) Then
                    InvProperty = ad.Propertys(GnCustom).Item(PnRawMaterial)
                    If InvProperty.Text Like "FM-*" Then
                        With New FmTest1
                            If .AskAbout(ad) = vbYes Then
                                InvProperty.Text = .ItemData.Item(PnRawMaterial)
                            End If
                        End With
                    End If
                End If
            End With
        Next
        Return 0
    End Function
    'Debug.Print(cnGnsDoyle.Execute("select I.ItemID, I.Thickness, I.Item, I.Description1 from vgMfiItems as I where I.Family='DSHEET'").GetString

    Public Function mdl1g1f0() As Long
        With New FmTest1
            .AskAbout(ThisApplication.ActiveDocument)
        End With
    End Function

    Public Function mdl1g1f2(Label As MSForms.Label) As Single ', txt As String
        Dim x0 As Long, x1 As Single
        Dim y0 As Long, y1 As Single
        Dim ct As MSForms.Control

        ct = Label
        With ct
            'x0 = .Left
            'y0 = .Top
            x1 = .Width
            y1 = .Height
            With Label
                '.Caption = txt
                .AutoSize = True
                .AutoSize = False
            End With
            .Width = x1
            mdl1g1f2 = .Height - y1
        End With
    End Function

    Public Function mdl1g1f3(
    ct As MSForms.Control,
    byX As Single, byY As Single
) As Single
        With ct
            .Left = .Left + byX
            .Top = .Top + byY
        End With
        mdl1g1f3 = Math.Sqrt((byX * byX) + (byY * byY))
    End Function

    ' For lack of a better place to put it, creating this node
    ' The following is a basic example of accessing Parameters,
    ' such as dimensions, from an Inventor Part Document.
    Public Function mdl1g1f1() As Long
        With aiDocPart(ThisApplication.ActiveDocument)
            With .ComponentDefinition.Parameters.Item("Thickness")
                Debug.Print("Thickness parameter exposed as property: " & .ExposedAsProperty)
                Debug.Print("Thickness parameter text: " & .Text)
            End With
        End With
        Return 0
    End Function
    ' This example was written as a quick one-off to see how
    ' an Inventor Parameter like the Thickness ting for
    ' Sheet Metal Parts might have its "Export" status
    ' modified programmatically.

    Public Function aiPropVal(InvProperty As Inventor.Property, Optional ifNot As Object = "") As Object
        If InvProperty Is Nothing Then
            Return ifNot
        Else
            Return aiPropValAux(InvProperty.Text, ifNot)
        End If
    End Function
    ' This example was written as a quick one-off to see how
    ' an Inventor Parameter like the Thickness ting for
    ' Sheet Metal Parts might have its "Export" status
    ' modified programmatically.

    Public Function aiPropValAux(vl As Object, Optional ifNot As Object = "") As Object
        If vl IsNot Nothing AndAlso vl.GetType.IsClass Then
            If TypeOf vl Is stdole.StdPicture Then
                Debug.Print("") ' Breakpoint Landing
                Return "<stdole.StdPicture>"
            Else
                Stop 'and see what we need to do
                Return "<Object:" & TypeName(vl) & ">"
            End If
        End If
        Return ifNot
    End Function

    Public Function aiPropGnsItmFamily(
    AiDoc As Inventor.Document
) As Inventor.Property
        If AiDoc Is Nothing Then
            aiPropGnsItmFamily = Nothing
        Else
            aiPropGnsItmFamily = AiDoc.Propertys(
            GnDesign).Item(PnFamily
        )
        End If
    End Function

    Public Function aiPropShtMetalThickness(
    adPart As Inventor.PartDocument
) As Inventor.Property
        If adPart Is Nothing Then
            aiPropShtMetalThickness = Nothing
        Else
            With adPart
                If .SubType = GuidSheetMetal Then
                    If smThicknessExposed(.ComponentDefinition) Then
                        aiPropShtMetalThickness = .Propertys(GnCustom).Item(PnThickness)
                    Else
                        aiPropShtMetalThickness = Nothing
                    End If
                Else
                    aiPropShtMetalThickness = Nothing
                End If
            End With
        End If
    End Function

    Public Function smThicknessExposed(
    smDef As Inventor.SheetMetalComponentDefinition
) As Long
        If smDef.Parameters.IsExpressionValid(PnThickness, "in") Then
            smThicknessExposed = parExposed(
        smDef.Parameters(PnThickness), 1
    )
            'With smDef.Parameters(pnThickness)
            'If Not .ExposedAsProperty Then
            '.ExposedAsProperty = True
            'End If
            'smThicknessExposed = IIf(.ExposedAsProperty, -1, 0)
            'End With
        Else
            Stop
        End If
    End Function

    Public Function parExposed(
    par As Inventor.Parameter,
    Optional tryTo As Long = 0
) As Long
        ''  Check Inventor Parameter for exposure as Property.
        ''  Return 0 if not, unless caller requests exposure
        ''  (tryTo <> 0). Nonzero return indicates exposed
        ''  Parameter, with sign indicating initial status.
        ''  -1 indicates Parameter already exposed
        ''  1 indicates status change to expose it.
        ''  No provision is made for failure to expose,
        ''  nor to reverse exposure status.
        With par
            If .ExposedAsProperty Then
                parExposed = -1
            ElseIf tryTo Then
                .ExposedAsProperty = True
                parExposed = 1 And parExposed(par)
            Else
                parExposed = 0
            End If
        End With
    End Function

    Public Function dcGnsPropsListed(
    ad As Inventor.Document, ls As Object,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional ifNone As Long = 1
) As Scripting.Dictionary
        '
        ' dcGnsPropsListed --
        '     Return a Dictionary of any
        '     Properties in the supplied list
        '     from the "custom" Property.
        '
        '     Missing Property names are addressed
        '     in one of three (present) ways,
        '     based on optional argument ifNone:
        '         0 - do not add to Dictionary
        '             missing name is missing
        '         1 - attempt to create. failure
        '             returns Nothing, which is
        '             still not added
        '         2 - add Nothing to Dictionary
        '             under missing name
        '         3 - attempt to create, adding
        '             Nothing for any failures
        '             (combines options 1 and 2)
        '
        Dim PropertySet As Inventor.Property
        Dim InvProperty As Inventor.Property
        Dim ky As Object
        Dim mkNf As Long 'try to make any not found
        Dim rtNf As Long 'return Nothing for not found
        Dim wk As Object

        ' rt = New Scripting.Dictionary
        PropertySet = ad.Propertys.Item(GnCustom)

        If dc Is Nothing Then
            dcGnsPropsListed = dcGnsPropsListed(
            ad, ls, New Scripting.Dictionary, ifNone
        )
        Else
            If IsArray(ls) Then
                mkNf = ifNone And 1 'IIf(ifNone = 1, 1, 0)
                rtNf = ifNone And 2 'IIf(ifNone = 2, 1, 0)
                '
                ' originally used IIf construct
                ' to force mapping of exact values
                ' of ifNone to corresponding behaviors.
                '
                ' changed to bitcode matching once clear
                ' that each bit would map exclusively
                ' to a particular behavior, and could
                ' be combined with the other, if desired.
                '

                For Each ky In ls '{pnMass _
                    '    , pnArea, pnWidth, pnLength _
                    '    , pnRawMaterial, pnRmQty, pnRmUnit _
                    ') ' _
                    '    , "SPEC01", "SPEC02", "SPEC03" _
                    '    , "SPEC04", "SPEC05", "SPEC06" _
                    '    , "SPEC07", "SPEC08", "SPEC16" _
                    ''
                    Dim oDoc As Inventor.Document
                    InvProperty = aiGetProp(PropertySet, CStr(ky), mkNf)
                    wk = oDoc

                    If InvProperty Is Nothing Then 'check
                        'if supposed to return Nothings

                        If rtNf = 0 Then wk = {}
                        'and empty the array if not
                    End If

                    If UBound(wk) < LBound(wk) Then
                        'we have an empty array,
                        'and nothing to add,
                        'not even Nothing!
                    Else
                        If dc.Exists(ky) Then
                            dc.Remove(ky)
                            ' WARNING[2021.11.19]
                            '     This was added to permit
                            '     replacement of elements
                            '     already present under a
                            '     supplied key. It might
                            '     NOT be the best way to
                            '     address this situation.
                            '     Be prepared to correct
                            '     this with a more robust
                            '     solution in future.
                            ' Meanwhile, have a
                            Debug.Print("") 'Breakpoint Landing
                            '
                        End If

                        dc.Add(ky, InvProperty) 'rt
                    End If
                Next
                dcGnsPropsListed = dc 'rt
            ElseIf VarType(ls) = vbString Then
                dcGnsPropsListed _
            = dcGnsPropsListed(
                ad, {ls},
                dc, ifNone
            )
            Else
                Stop
            End If
        End If
    End Function

    Public Function dcGnsPropsPart(ad As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional ifNone As Long = 1
) As Scripting.Dictionary
        '
        ' dcGnsPropsPart
        '
        ' REV[2021.11.18]:
        '     Added pnThickness to list
        '     of Properties to return.
        '
        dcGnsPropsPart = dcGnsPropsListed(ad, {PnMass, PnArea, PnWidth, PnLength, PnThickness, PnRawMaterial, PnRmQty, PnRmUnit}, dc, ifNone)
    End Function

    Public Function dcGnsPropsAssy(ad As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional ifNone As Long = 1
) As Scripting.Dictionary
        'Dim rt As Scripting.Dictionary
        'Dim PropertySet As Inventor.Property
        'Dim Property As Inventor.Property
        'Dim ky As Object

        dcGnsPropsAssy = dcGnsPropsListed(ad, New String() {PnMass, "SPEC01", "SPEC02", "SPEC03", "SPEC04", "SPEC05", "SPEC06", "SPEC07", "SPEC08", "SPEC16"}, dc, ifNone)
        ' rt = New Scripting.Dictionary
        ' PropertySet = ad.Propertys.Item(gnCustom)

        'If dc Is Nothing Then
        ' dcGnsPropsAssy = dcGnsPropsAssy( _
        '        ad(), New Scripting.Dictionary _
        ')
        'Else
        ''        'For Each ky In {pnMass _
        '            , "SPEC01", "SPEC02", "SPEC03" _
        '            , "SPEC04", "SPEC05", "SPEC06" _
        '            , "SPEC07", "SPEC08", "SPEC16" _
        '        ) ' _
        '            , pnArea, pnWidth, pnLength _
        '            , pnRawMaterial, pnRmQty, pnRmUnit _
        '        ''
        ' Property = aiGetProp(ps, CStr(ky), 1)
        'If Property Is Nothing Then
        'nothing we can do (as yet?)
        'Else
        'dc.Add(ky, Property 'rt
        'End If
        'Next
        ' dcGnsPropsAssy = dc 'rt
        'End If'
    End Function

    Public Function dcProps4genius(ad As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional Create As Long = 1
) As Scripting.Dictionary
        'Dim rt As Scripting.Dictionary
        Dim PropertySet As Inventor.Property
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        If dc Is Nothing Then
            dcProps4genius = dcProps4genius(
            ad, New Scripting.Dictionary, Create
        )
        Else
            With ad
                If .DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    dcProps4genius = dcGnsPropsAssy(ad, dc, Create)
                ElseIf .DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    dcProps4genius = dcGnsPropsPart(ad, dc, Create)
                Else
                    dcProps4genius = dc
                End If
            End With
        End If

        'Stop 'With Exit Function above, should never get here
        '
        '' rt = New Scripting.Dictionary
        ' PropertySet = ad.Propertys.Item(gnCustom)
        '
        'If dc Is Nothing Then
        '     dcProps4genius = dcProps4genius( _
        '        ad, New Scripting.Dictionary _
        '    )
        'Else
        '    For Each ky In {pnMass _
        '        , pnArea, pnWidth, pnLength _
        '        , pnRawMaterial, pnRmQty, pnRmUnit _
        '        , "SPEC01", "SPEC02", "SPEC03" _
        '        , "SPEC04", "SPEC05", "SPEC06" _
        '        , "SPEC07", "SPEC08", "SPEC16" _
        '    ) ' _
        '    '
        '         Property = aiGetProp(ps, CStr(ky), 1)
        '        If Property Is Nothing Then
        '            'nothing we can do (as yet?)
        '        Else
        '            dc.Add(ky, Property 'rt
        '        End If
        '    Next
        '     dcProps4genius = dc 'rt
        'End If
    End Function

    Public Function mdl1g2f1(ad As Inventor.Document) As Inventor.WorkPlanes
        mdl1g2f1 = aiDocPart(ad).ComponentDefinition.WorkPlanes
    End Function

    Public Function mdl1g3f0(ad As Inventor.Document) As Double
        Dim rt As Double

        Select Case ad.DocumentType
            Case DocumentTypeEnum.kPartDocumentObject
                rt = aiDocPart(ad).ComponentDefinition.MassProperties.Mass
            Case DocumentTypeEnum.kAssemblyDocumentObject
                rt = aiDocAssy(ad).ComponentDefinition.MassProperties.Mass
            Case Else
                rt = 0#
        End Select

        With ad.UnitsOfMeasure
            mdl1g3f0 = .ConvertUnits(rt,
            UnitsTypeEnum.kKilogramMassUnits,
            .kLbMassMassUnits
        ) '.MassUnits)
        End With
    End Function

    Public Function mdl1g4f0() As Long
        Dim mx As Long
        Dim dx As Long

        With ThisApplication.CommandManager.ControlDefinitions
            mx = .Count
            For dx = 1 To mx
                With .Item(dx)
                    If .InternalName Like "*ault*" Then
                        Debug.Print(CStr(dx) & ": " & .InternalName & "/" & .DisplayName)
                    End If
                End With
            Next
        End With
    End Function

    Public Function mdl1g5f0(ad As Inventor.Document) As Scripting.Dictionary
        ' The purpose of this function is to return a Dictionary
        ' of Genius Family Inventor Properties
        ' for each component Document of an assembly
        ' or a single part Document.
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dcAiDocComponents(ad, New Scripting.Dictionary, 1) 'sc
            For Each ky In .Keys
                With aiDocument(.Item(ky))
                    rt.Add(.FullFileName, aiPropGnsItmFamily(.Propertys.Parent))
                End With
            Next
        End With
        mdl1g5f0 = rt
    End Function

    Public Function mdl1g5f1(ad As Inventor.Document) As Scripting.Dictionary
        ' This function calls mdl1g5f0 to retrieve a Dictionary
        ' of Genius Family Inventor Properties, and then
        ' transforms it into a Dictionary of Dictionaries
        ' grouped by Family Property Value
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim ky As Object
        Dim Fm As String
        Dim InvProperty As Inventor.Property

        rt = New Scripting.Dictionary
        With mdl1g5f0(ad)
            For Each ky In .Keys
                InvProperty = aiProperty(.Item(ky))
                Fm = InvProperty.Text
                With rt
                    If Not .Exists(Fm) Then
                        .Add(Fm, New Scripting.Dictionary)
                    End If
                    dcOb(.Item(Fm)).Add(ky, InvProperty)
                End With
            Next
        End With
        mdl1g5f1 = rt
    End Function
    'Debug.Print(txDumpLs(mdl1g5f1(ThisApplication.ActiveDocument).Keys)
    'Debug.Print(txDumpLs(dcOb(mdl1g5f1(ThisApplication.ActiveDocument).Item("")).Keys)

    Public Function mdl1g5f2(ad As Inventor.Document) As Scripting.Dictionary
        ' The purpose of this function is to return a Dictionary
        ' of Genius Family Inventor Properties
        ' for each component Document of an assembly
        ' or a single part Document.
        Dim rt As Scripting.Dictionary
        Dim Fm As FmTest1
        Dim ky As Object

        Fm = New FmTest1
        rt = New Scripting.Dictionary
        With mdl1g5f1(ad)
            If .Exists("") Then
                With dcOb(.Item(""))
                    For Each ky In .Keys
                        With aiProperty(.Item(ky))
                            If Fm.AskAbout(.Parent) = vbOK Then
                                Stop
                            Else
                                Stop
                            End If
                            Stop
                            'rt.Add(.FullFileName, aiPropGnsItmFamily(.Propertys.Parent)
                        End With
                    Next
                End With
            End If
        End With
        mdl1g5f2 = rt
    End Function

    Public Function mdl1g5f3(ad As Inventor.AssemblyDocument) As Scripting.Dictionary
        ' Scan immediate members of Assembly document
        ' and group in Dictionary by declared Part Number
        ' and sub-grouped by Full Document Name.
        '
        ' (I wonder if an ADO Record wouldn't be a better choice?)
        '
        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.Document
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        'Dim Fm As FmTest1
        'Dim ky As Object
        Dim pn As String

        ' Fm = New FmTest1

        rt = New Scripting.Dictionary
        For Each oc In ad.ComponentDefinition.Occurrences
            sd = oc.Definition.Document 'aiDocument()
            pn = sd.Propertys.Item(GnDesign).Item(PnPartNum).Text
            With rt
                If .Exists(pn) Then
                    gp = dcAiDocsByFullDocName(sd, .Item(pn))
                Else
                    .Add(pn, dcAiDocsByFullDocName(sd,
                    New Scripting.Dictionary
                ))
                End If
            End With
        Next
        mdl1g5f3 = rt
    End Function
    'Debug.Print(txDumpLs(mdl1g5f3(ThisApplication.ActiveDocument).Keys)

    Public Function mdl1g5f4(dc As Scripting.Dictionary) As Scripting.Dictionary
        ' Transform keys from supplied Dictionary
        ' (expected from mdl1g5f3)
        ' into header/member indented form.
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim dl As String

        dl = vbCrLf & vbTab
        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                rt.Add(ky & vbTab & Join(
                dcOb(.Item(ky)).Keys,
                vbCrLf & ky & vbTab
            ), .Item(ky))
            Next
        End With
        mdl1g5f4 = rt
    End Function
    'Debug.Print(txDumpLs(mdl1g5f4(mdl1g5f3(ThisApplication.ActiveDocument)).Keys)

    Public Function dcAiDocsByFullDocName(
    ad As Inventor.Document,
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ' Add supplied Inventor Document
        ' to supplied Dictionary
        ' under its Full Document Name
        ' (supports mdl1g5f3)
        Dim ky As String

        ky = ad.FullDocumentName
        If dc Is Nothing Then
            dcAiDocsByFullDocName = dcAiDocsByFullDocName(ad, New Scripting.Dictionary)
        Else
            With dc
                If .Exists(ky) Then
                    .Item(ky) = 1 + .Item(ky)
                Else
                    .Add(ky, 1)
                End If
            End With
            dcAiDocsByFullDocName = dc
        End If
    End Function

    Public Function dcAssyDocsByPtNum(ad As Inventor.AssemblyDocument) As Scripting.Dictionary
        ' Derived from mdl1g5f3
        '
        ' Scan immediate members of Assembly Document
        ' and collect source Documents in Dictionary,
        ' grouped by declared Part Number.
        '
        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.Document
        Dim rt As Scripting.Dictionary
        Dim pn As String

        rt = New Scripting.Dictionary
        For Each oc In ad.ComponentDefinition.Occurrences
            sd = oc.Definition.Document
            pn = sd.Propertys.Item(GnDesign).Item(PnPartNum).Text
            With rt
                If .Exists(pn) Then
                    If sd Is .Item(pn) Then 'we're okay
                        'so carry on
                    Else 'we've got a problem, so
                        Stop 'and check it out
                    End If
                Else
                    .Add(pn, sd)
                End If
            End With
        Next
        dcAssyDocsByPtNum = rt
    End Function
    'Debug.Print(txDumpLs(dcAssyDocsByPtNum(ThisApplication.ActiveDocument).Keys)

    Public Function dcAiDocsByCompList(dc As Scripting.Dictionary) As Scripting.Dictionary
        ' Derived from mdl1g5f4
        ' Transform keys from supplied Dictionary
        ' (expected from dcAssyDocsByPtNum)
        ' into tab-delimited list form.
        Dim rt As Scripting.Dictionary
        Dim sd As Inventor.Document
        Dim ky As Object
        Dim dl As String

        dl = vbCrLf & vbTab
        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                sd = aiDocument(.Item(ky))
                With sd
                    If .DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        rt.Add(ky & vbTab & Join(
                        Split(TxDumpLs(mdl1g5f4(mdl1g5f3(sd)).Keys), vbCrLf),
                        vbCrLf & ky & vbTab
                         ), sd)
                    ElseIf .DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        'Stop
                        With dcAiPropsIn(sd.Propertys.Item(GnCustom))
                            If .Exists(PnRawMaterial) Then
                                dl = Trim$(aiProperty(.Item(PnRawMaterial)).Text)
                                If Len(dl) = 0 Then
                                    dl = "NO_RAW_STOCK" & vbTab & "<No Raw Stock Declared>"
                                Else
                                    Stop
                                    With CnGnsDoyle.Execute(Join({(
                                    "select Description1",
                                    "from vgMfiItems",
                                    "where Item = '" & dl & "';"
                                ), vbCrLf}))
                                        If .BOF Or .EOF Then
                                            dl = dl & vbTab & "<Stock Number Not Found>"
                                        Else
                                            dl = dl & vbTab & .Fields(0).Value
                                        End If
                                    End With
                                End If
                            Else
                                dl = "NO_RAW_STOCK" & vbTab & "<No Raw Stock Declared>"
                            End If
                        End With
                        rt.Add(ky & vbTab & dl, sd)
                        'rt.Add(ky & vbTab & "(RAW STOCK NOT YET IMPLMENTED)", sd
                    Else
                        rt.Add(ky & vbTab & "(UNSUPPORTED DOCUMENT TYPE)", sd)
                    End If
                End With
            Next
        End With
        dcAiDocsByCompList = rt
    End Function
    'Debug.Print(txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)
    'send2clipBd txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)

    Public Function rsWinUpdHist() As ADODB.Record
        ' Windows Update History
        Dim it As WUApiLib.IUpdateHistoryEntry
        Dim rt As ADODB.Record
        Dim ls As Object

        rt = rsNewWinUpdHist()
        ls = {"ResultCode", "Operation", "Title", "Description", "Date"}
        With New WUApiLib.UpdateSession '.CreateUpdateSearcher
            With .CreateUpdateSearcher
                For Each it In .QueryHistory(0, .GetTotalHistoryCount)
                    With it
                        'Debug.Print(.ResultCode, .Operation, .Title, .Description, .Date
                        rt.AddNew(ls, {
                        .ResultCode, .Operation,
                        .Title, .Description, .Date
                    })
                    End With
                Next
                rt.Filter = ""
            End With
        End With
        rsWinUpdHist = rt
    End Function

    Public Function rsNewWinUpdHist() As ADODB.Record
        Dim rt As ADODB.Record
        rt = New ADODB.Record
        With rt
            With .Fields
                '.Append "", adBigInt
                '.Append "", adVarChar, 1024
                .Append("ResultCode", DataTypeEnum.adBigInt)
                .Append("Operation", DataTypeEnum.adBigInt)
                .Append("Title", DataTypeEnum.adVarChar, 256)
                .Append("Description", DataTypeEnum.adVarChar, 1024)
                .Append("Date", DataTypeEnum.adDBDate)
            End With
            .Open()
        End With
        rsNewWinUpdHist = rt
    End Function

    Public Function rsShtMtlCutPars(ThisApplication As Inventor.Application,
    ad As Inventor.Document,
    Optional incTop As Long = 0
) As ADODB.Record
        ' Windows Update History
        Dim rt As ADODB.Record
        Dim oDoc As Inventor.Document
        Dim ls As Object
        Dim ky As Object

        rt = rsNewShtMtlCutPars()
        ls = {"Item", "Description", "Thickness", "Perimeter"}

        With dcAiDocComponents(ad, , incTop)
            For Each ky In .Keys
                oDoc = aiDocument(.Item(ky))
                With oDoc.Propertys.Item(GnDesign)
                    rt.AddNew(ls, {
                    .Item(PnPartNum).Text,
                    .Item(PnDesc).Text,
                    aiPropVal(aiPropShtMetalThickness(aiDocPart(oDoc)), -1),
                    fpPerimeterInch(oDoc, ThisApplication)
                })
                End With
            Next
            rt.Filter = ""
        End With

        rsShtMtlCutPars = rt
    End Function
    'send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , "|")
    'send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , vbTab)

    Public Function rsNewShtMtlCutPars() As ADODB.Record
        Dim rt As ADODB.Record

        rt = New ADODB.Record
        With rt
            With .Fields
                '.Append "", adBigInt
                '.Append "", adVarChar, 1024
                '.Append "Date", adDBDate
                '
                .Append("Item", DataTypeEnum.adVarChar, 32, FieldAttributeEnum.adFldKeyColumn)
                .Append("Description", DataTypeEnum.adVarChar, 128)
                .Append("Thickness", DataTypeEnum.adDouble)
                .Append("Perimeter", DataTypeEnum.adDouble)
            End With
            .Open
        End With

        rsNewShtMtlCutPars = rt
    End Function
End Module