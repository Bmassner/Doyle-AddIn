Module dvlDict0
    Dim inventorApp As Inventor.Application
    Public Function aiDocPartNum(
    AiDoc As Inventor.Document,
    Optional ifNone As String = ""
) As String
        If AiDoc Is Nothing Then
            aiDocPartNum = ifNone
        Else
            aiDocPartNum = AiDoc.Propertys _
        .Item(gnDesign) _
        .Item(pnPartNum) _
        .Text
        End If
    End Function

    Public Function dc0g1f0(
    AiDoc As Inventor.Document,
    Optional PropertyName As String = "",
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim pn As String
        Dim ky As String

        If dc Is Nothing Then
            dc0g1f0 = dc0g1f0(
            AiDoc, PropertyName,
            New Scripting.Dictionary
        )
        Else
            rt = dc

            With AiDoc
                pn = .Propertys.Item(gnDesign).Item(pnPartNum).Text
                ky = PropertyName & vbTab & pn
                With rt
                    If .Exists(ky) Then
                        If .Item(ky) Is AiDoc Then
                        Else
                            'Stop
                        End If
                    Else
                        .Add(ky, AiDoc)
                    End If
                End With

                If .DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    rt = dc0g1f1(AiDoc, rt)
                ElseIf .DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    Stop
                Else
                End If
            End With

            dc0g1f0 = rt
        End If
    End Function

    Public Function dc0g1f1(
    AiDoc As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim aiOcc As Inventor.ComponentOccurrence
        Dim pn As String

        If dc Is Nothing Then
            dc0g1f1 = dc0g1f1(AiDoc,
        New Scripting.Dictionary)
        Else
            rt = dc

            With AiDoc
                pn = .Propertys.Item(gnDesign).Item(pnPartNum).Text
                With .ComponentDefinition
                    For Each aiOcc In .Occurrences
                        rt = dc0g1f0(aiOcc.Definition.Document, pn, rt)
                    Next
                End With
            End With

            dc0g1f1 = rt
        End If
    End Function

    Public Function dc0g2f0(
    Optional AiDoc As Inventor.Document = Nothing
) As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If AiDoc Is Nothing Then
            dc0g2f0 = dc0g2f0(
            InventorApp.ActiveDocument
        )
        Else
            wk = dc0g2f0.dcAiDocsByPtNum(
            dcAiDocComponents(AiDoc)
        ) 'dcAiDocPartNumbers
            rt = New Scripting.Dictionary
            With wk
                For Each ky In .Keys
                    rt = dc0g2f2(aiDocument(obOf(.Item(ky))), rt)
                Next
            End With
            dc0g2f0 = rt
        End If
    End Function

    Public Function dc0g2f1(
    AiDoc As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        'Dim rt As Scripting.Dictionary
        Dim aiOcc As Inventor.ComponentOccurrence
        Dim PropertyName As String
        Dim ptName As String
        Dim ky As String
        Dim ct As Long

        If dc Is Nothing Then
            dc0g2f1 = dc0g2f1(AiDoc,
        New Scripting.Dictionary)
        Else
            PropertyName = aiDocPartNum(AiDoc)
            With AiDoc.ComponentDefinition
                For Each aiOcc In .Occurrences
                    ptName = aiDocPartNum(
                    aiOcc.Definition.Document
                )
                    ky = PropertyName & vbTab & ptName

                    With dc
                        If .Exists(ky) Then
                            ct = .Item(ky)
                            .Item(ky) = 1 + ct
                        Else
                            .Add(ky, 1)
                        End If
                    End With
                Next
            End With
            dc0g2f1 = dc
        End If
    End Function

    Public Function dc0g2f2(AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        If AiDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            dc0g2f2 = dc0g2f1(AiDoc, dc)
        Else
            dc0g2f2 = dc
        End If
    End Function

    Public Function dc0g3f0(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' (just so we don't forget what this is for)
        ' This function accepts a Dictionary
        ' of Inventor Documents, and generates
        ' a new Dictionary of Dictionaries,
        ' keyed on Genius Family names, and
        ' containing all Documents in its Family,
        ' themselves keyed on Part Number.
        '
        ' Function dc0g3f1 makes use of this below.
        '
        Dim rt As Scripting.Dictionary
        Dim Fm As Scripting.Dictionary
        Dim ky As Object
        Dim ad As Inventor.Document
        Dim nm As String
        Dim pn As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                ad = aiDocument(obOf(.Item(ky)))
                If ad Is Nothing Then
                    'Stop
                Else
                    With ad.Propertys.Item(gnDesign)
                        nm = .Item(pnFamily).Text
                        pn = .Item(pnPartNum).Text
                    End With

                    With rt
                        If .Exists(nm) Then
                            Fm = .Item(nm)
                        Else
                            Fm = New Scripting.Dictionary
                            .Add(nm, Fm)
                        End If
                    End With

                    With Fm
                        If .Exists(pn) Then
                            Stop
                        Else
                            .Add(pn, ad)
                        End If
                    End With
                End If
            Next
        End With

        dc0g3f0 = rt
    End Function

    Public Function dc0g3f1() As Scripting.Dictionary
        '
        ' This function calls dc0g3f0 above
        ' against a Dictionary of Inventor Documents
        ' generated from the components of the active
        ' Inventor Document. It then dumps a list of
        ' the Genius Family names encountered, and if
        ' any were blank, the list of part numbers
        ' in the blank Family group is also revealed.
        '
        With dc0g3f0(dcAssyDocComponents(aiDocAssy(aiDocActive())))
            Debug.Print(TxDumpLs(.Keys))
            Stop
            If .Exists("") Then
                Debug.Print("NO FAMILY")
                With dcOb(.Item(""))
                    Debug.Print(TxDumpLs(.Keys))
                End With
                Stop
            Else
                Stop
            End If
        End With
    End Function

    Public Function dc0g4f0(AiDoc As Inventor.AssemblyDocument) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ki As Object

        rt = New Scripting.Dictionary
        With dcAssyComponentsImmediate(AiDoc)
            For Each ky In .Keys
                With dcAssyComponentsImmediate(aiDocAssy(.Item(ky)))
                    For Each ki In .Keys
                        If Not rt.Exists(ki) Then
                            rt.Add(ki, .Item(ki))
                        End If
                    Next
                End With
            Next
        End With
        dc0g4f0 = rt
    End Function
    'Debug.Print(txDumpLs(dcAssyComponentsImmediate(InventorApp.ActiveDocument).Keys)
    'Debug.Print(txDumpLs(dc0g4f0(InventorApp.ActiveDocument).Keys)
    'Debug.Print(txDumpLs(dcAiDocPartNumbers(dc0g4f0(InventorApp.ActiveDocument)).Keys)

    Public Function dc0g4f1(
    adIn As Inventor.AssemblyDocument,
    adOut As Inventor.AssemblyDocument
) As Scripting.Dictionary
        Dim ky As Object
        Dim tg As Inventor.ComponentOccurrences
        Dim oc As Inventor.ComponentOccurrence
        Dim PropertySet As Inventor.Matrix

        PropertySet = InventorApp.TransientGeometry.CreateMatrix()

        tg = adOut.ComponentDefinition.Occurrences
        With dc0g4f0(adIn)
            For Each ky In .Keys
                oc = tg.Add(ky, PropertySet)
            Next
        End With
    End Function

    Public Function dcBoxDims(bx As Inventor.Box) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim mx As Inventor.Point
        Dim mn As Inventor.Point

        rt = New Scripting.Dictionary

        With bx
            mx = .MaxPoint
            mn = .MinPoint
        End With

        With rt
            .Add("X", mx.X - mn.X) '/ 2.54
            .Add("Y", mx.Y - mn.Y) '/ 2.54
            .Add("Z", mx.Z - mn.Z) '/ 2.54
        End With

        dcBoxDims = rt
    End Function

    Public Function dcBoxDimsCm2in(bx As Inventor.Box) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary

        With dcBoxDims(bx)
            For Each ky In .Keys
                rt.Add(ky, CDbl(.Item(ky)) / cvLenIn2cm)
            Next
        End With

        dcBoxDimsCm2in = rt
    End Function

    Public Function dcAiPropsIn(
    PropertySet As Inventor.Property
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property

        rt = New Scripting.Dictionary
        'For Each InvProperty In PropertySet
        '    If rt.Exists(InvProperty.Name) Then
        '        Stop
        '    Else
        '        rt.Add(InvProperty.Name, InvProperty)
        '    End If
        'Next
        dcAiPropsIn = rt
    End Function
    'Debug.Print(Join(Filter(dcAiPropsIn(InventorApp.ActiveDocument.Propertys(gnDesign)).Keys, "web", , vbTextCompare), vbCrLf)

    Public Function dcAiDocParVals(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        dcAiDocParVals _
    = dcAiParValues(
    dcAiDocParams(AiDoc))
    End Function

    Public Function dcAiParValues(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim InvProperty As Inventor.Parameter

        rt = New Scripting.Dictionary

        If dc Is Nothing Then
        Else
            'With dc : For Each ky In .Keys
            '        InvProperty = obAiParam(obOf(.Item(ky)))
            '        If InvProperty Is Nothing Then
            '        Else
            '            rt.Add(ky, {PartDoc.Text, InvProperty.Units})
            '        End If
            '    Next : End With
        End If

        dcAiParValues = rt
    End Function

    Public Function dcAiDocParams(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        dcAiDocParams _
    = dcCompDefParams(
    compDefOf(AiDoc))
    End Function
    'Debug.Print(Join(Filter(dcAiDocParams(InventorApp.ActiveDocument.Propertys(gnDesign)).Keys, "web", , vbTextCompare), vbCrLf)

    Public Function dcCompDefParams(
    CpDef As Inventor.ComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        If CpDef Is Nothing Then
            dcCompDefParams = New Scripting.Dictionary
        ElseIf TypeOf CpDef Is Inventor.AssemblyComponentDefinition Then
            dcCompDefParams = dcCompDefParamsAssy(CpDef)
        ElseIf TypeOf CpDef Is Inventor.PartComponentDefinition Then
            dcCompDefParams = dcCompDefParamsPart(CpDef)
        Else
            dcCompDefParams = dcCompDefParams(Nothing)
        End If
    End Function

    Public Function dcCompDefParamsPart(
    CpDef As Inventor.PartComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim InvProperty As Inventor.Parameters

        If CpDef Is Nothing Then
            InvProperty = Nothing
        Else
            InvProperty = CpDef.Parameters
        End If

        dcCompDefParamsPart _
    = dcOfAiParameters(InvProperty, dc)
    End Function

    Public Function dcCompDefParamsAssy(
    CpDef As Inventor.AssemblyComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim InvProperty As Inventor.Parameters

        If CpDef Is Nothing Then
            InvProperty = Nothing
        Else
            InvProperty = CpDef.Parameters
        End If

        dcCompDefParamsAssy _
    = dcOfAiParameters(InvProperty, dc)
    End Function

    Public Function dcOfAiParameters(
    AiPars As Inventor.Parameters,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Parameter

        If dc Is Nothing Then
            rt = dcOfAiParameters(AiPars,
        New Scripting.Dictionary)
        Else
            rt = dc

            If AiPars Is Nothing Then
            Else
                For Each InvProperty In AiPars
                    rt.Add(InvProperty.Name, InvProperty)
                Next
            End If
        End If

        dcOfAiParameters = rt
    End Function

    Public Function dcOfPropsInAiDoc(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim PropertySet As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary

        If AiDoc Is Nothing Then
        Else
            With AiDoc : For Each PropertySet In .Propertys
                    wk = dcAiPropsIn(PropertySet)

                    With DcKeysMissing(wk, rt)
                        For Each ky In .Keys
                            rt.Add(ky, .Item(ky))
                            wk.Remove(ky)
                        Next : End With

                    With wk 'dcKeysInCommon(wk, rt)
                        If .Count > 0 Then
                            Debug.Print("=== DUPLICATE PROPERTY NAMES ===")
                            Debug.Print("  Item " & aiProperty(rt.Item(pnPartNum)).Text & " (" & AiDoc.FullDocumentName & ")")
                            Debug.Print(DumpLsKeyVal(dcPropVals(wk), ": "))
                            Debug.Print("--- previously found")
                            Debug.Print(DumpLsKeyVal(dcPropVals(DcKeysInCommon(wk, rt, 2)), ": "))
                            Debug.Print("") 'Breakpoint Landing
                        End If
                    End With
                Next : End With
        End If

        dcOfPropsInAiDoc = rt
    End Function

    Public Function dcAiPropValsFromDc(
    dc As Scripting.Dictionary,
    Optional Flags As Long = 0
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        With DcNewIfNone(dc) : For Each ky In .Keys
                InvProperty = aiProperty(obOf(.Item(ky)))
                If InvProperty Is Nothing Then
                    If Flags And 1 Then
                        'Keep non-Property Items
                        rt.Add(ky, .Item(ky))
                    End If
                Else
                    rt.Add(ky, aiPropVal(InvProperty, vbEmpty))
                End If
            Next : End With

        Debug.Print("") 'Breakpoint Landing
        dcAiPropValsFromDc = rt
    End Function

    Public Function dcForAiDocIType(
    dc As Scripting.Dictionary,
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ky As String

        If TypeOf AiDoc Is Inventor.PartDocument Then
            ky = "Part"
            With aiDocPart(AiDoc).ComponentDefinition
                If .IsContentMember Then ky = "c" & ky
                If .IsiPartFactory Then ky = "f" & ky
                If .IsiPartMember Then ky = "i" & ky
                If .IsModelStateFactory Then 'ky = "msf" & ky
                    ' ???
                End If
                If .IsModelStateMember Then ky = "s" & ky
            End With
        ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
            ky = "Assy"
            With aiDocAssy(AiDoc).ComponentDefinition
                If .IsiAssemblyFactory Then ky = "f" & ky
                If .IsiAssemblyMember Then ky = "i" & ky
                If .IsModelStateFactory Then ky = "msf" & ky
                If .IsModelStateMember Then ky = "s" & ky
            End With
        Else
            ky = ""
        End If

        With dc
            If Not .Exists(ky) Then
                .Add(ky, New Scripting.Dictionary)
            End If
            dcForAiDocIType = .Item(ky)
        End With
    End Function

    Public Function dcAiDocsByIType(
    dc As Scripting.Dictionary,
    Optional Flags As Long = 0
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        'Dim Property As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        With DcNewIfNone(dc) : For Each ky In .Keys
                ' Property = aiProperty(obOf(.Item(ky)))
                'If Property Is Nothing Then
                'If Flags And 1 Then
                'Keep non-Property Items
                rt.Add(ky, .Item(ky))
                'End If
                'Else
                'rt.Add(ky, aiPropVal(InvProperty, Empty)
                'End If
            Next : End With

        Debug.Print("") 'Breakpoint Landing
        dcAiDocsByIType = rt
    End Function

    Public Function nvmTest01() As Object
        Dim ad As Inventor.ApplicationAddIn
        Dim il As Object 'Inventor.ApplicationAddIn '
        Dim mp As Inventor.NameValueMap
        Dim md As Inventor.Document

        ad = InventorApp.ApplicationAddIns.ItemById(guidILogicAdIn)
        If Not ad.Activated Then ad.Activate()
        il = ad.Automation
        md = InventorApp.Documents.ItemByName("C:\Doyle_Vault\Designs\Misc\andrewT\dvl\iLogVltSrch_2022-0622_01.ipt")
        mp = dc2aiNameValMap(
        NuDcPopulator().Setting(
        "PartNumber", "60-"
    ).Dictionary)  'IN 60- 04-

        il.RunRuleWithArguments(md, "vlt02", mp) 'tst01
        'il.RunRule md, "tst01" ', mp

        Debug.Print(mp.Text("OUT"))

        Debug.Print(mp.Count)
    End Function

    '
    '
    '
    Public Function dvlDict0() As String
        dvlDict0 = "dvlDict0"
    End Function
End Module