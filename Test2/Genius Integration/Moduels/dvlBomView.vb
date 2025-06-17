Module dvlBomView
    Dim ThisApplication As Inventor.Application
    Public Function bomViewStruct(
    PartDesc As Inventor.AssemblyDocument
) As Inventor.BOMView
        '
        ' bomViewStruct -- Get Structured BOM View
        '     for supplied Assembly, if available
        '
        Dim BillVeiw As Inventor.BOMView
        Dim BillRow As Inventor.BOMRow

        If PartDesc Is Nothing Then
            BillVeiw = Nothing
        Else
            With PartDesc 'aiDocAssy(aiDocActive())
                On Error Resume Next
                Err.Clear()
                BillVeiw = .ComponentDefinition.BOM.BOMViews.Item("Structured")

                If Err.Number = 0 Then 'we're okay
                Else 'we got nothin'
                    BillVeiw = Nothing
                End If

                On Error GoTo 0
                'Stop
            End With
        End If
        bomViewStruct = BillVeiw
    End Function

    Public Function dBVg1f1(itmPath As String) As String()
        Dim rt(1) As String
        Dim bk As Long

        bk = InStrRev(itmPath, ".")

        If bk > 0 Then
            rt(0) = Left$(itmPath, bk - 1)
            rt(1) = Mid$(itmPath, bk + 1)
        Else
            rt(0) = ""
            rt(1) = itmPath
        End If

        dBVg1f1 = rt
    End Function

    Public Function bomLnumBkDn(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim ls() As String

        With dc
            If .Exists("path") Then
                ls = dBVg1f1(dc.Item("path"))
                .Item("base") = ls(0)
                .Item("seq") = ls(1)
            End If
        End With
        bomLnumBkDn = dc
    End Function

    Public Function bomLineInfo(
    brw As Inventor.BOMRow
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        rt = New Scripting.Dictionary
        With brw
            rt.Add("bomStruct", .BOMStructure)
            rt.Add("path", .ItemNumber)
            'rt.Add("seq", .ItemNumber

            rt.Add("qty", .ItemQuantity)
            rt.Add("qtTotal", .TotalQuantity)
            rt.Add("qtUnit", "EA")

            rt.Add("mrg", .Merged)
            rt.Add("pro", .Promoted)
            rt.Add("rol", .RolledUp)

            '.ChildRows
        End With
        bomLineInfo = bomLnumBkDn(rt)
    End Function

    Public Function dBVg1f2(
    AiDoc As Inventor.Document,
    wk As Scripting.Dictionary,
    dc As Scripting.Dictionary
    ) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim bs As String
        Dim pn As String
        Dim rm As String
        Dim k0 As String
        Dim k1 As Object

        With wk
            bs = CStr(.Item("path"))
            pn = CStr(.Item("ptNum"))
        End With

        With dcOfPropsInAiDoc(AiDoc)
            If .Exists("RM") Then
                rm = CStr(aiProperty(.Item("RM")).Text)
                k0 = pn & "|" & rm

                rt = New Scripting.Dictionary
                rt.Add("bomStruct", BOMStructureEnum.kPurchasedBOMStructure)

                rt.Add("path", bs & ".1")

                If .Exists("RMQTY") Then
                    rt.Add("qty", CStr(aiProperty(.Item("RMQTY")).Text))
                Else
                    rt.Add("qty", -1)
                End If
                rt.Add("qtTotal", rt.Item("qty"))

                If .Exists("RMUNIT") Then
                    rt.Add("qtUnit", CStr(aiProperty(.Item("RMUNIT")).Text))
                Else
                    rt.Add("qtUnit", "EA")
                End If

                rt.Add("mrg", False)
                rt.Add("pro", False)
                rt.Add("rol", False)

                rt.Add("base", CStr(bs))
                rt.Add("seq", "1")

                rt.Add("ptNum", CStr(rm))
                'rt.Add("aiDoc", ""

                If dc.Exists(CStr(k0)) Then
                    ck = dcOb(dc.Item(CStr(k0)))
                    'send2clipBd ConvertToJson( _
                    DcWBQbyCmpResult(
                        DcCmpTextOf2dc(ck, rt)
                    )
                    With DcWBQbyCmpResult(
                        DcCmpTextOf2dc(
                            ck, rt
                        )
                    )
                        If .Exists("!=") Then
                            With dcOb(.Item("!="))
                                For Each k1 In .Keys
                                    ck.Item(k1) = ck.Item(k1) _
                                        & vbTab & rt.Item(k1)
                                    'Debug.Print(Join({ _
                                    'k0(), k1, ck.Item(k1) ), "|")
                                    'Stop
                                Next
                            End With
                        End If
                    End With
                    Debug.Print("") 'Breakpoint Landing
                    'Stop
                Else
                    dc.Add(CStr(pn & "|" & rm), rt)
                End If
            Else
            End If
        End With

        'Stop
        dBVg1f2 = dc
    End Function

    Public Function bomItemInfo(
    BillRow As Inventor.BOMRow,
    Optional dc As Scripting.Dictionary = Nothing
    ) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        'Dim wk As Scripting.Dictionary
        Dim df As Inventor.ComponentDefinition
        Dim PartDoc As Inventor.Document
        Dim pn As String
        Dim ck As String
        Dim fn As String
        Dim ct As Long

        If dc Is Nothing Then
            rt = bomLineInfo(BillRow)
        Else
            rt = dc
        End If

        With BillRow
            With .ComponentDefinitions
                ct = .Count
                If ct > 0 Then
                    With .Item(1)
                        PartDoc = aiDocument(.Document)
                        pn = aiDocPartNum(PartDoc)
                        fn = PartDoc.FullDocumentName
                    End With
                Else
                    Stop
                    pn = ""
                    fn = ""
                End If

                rt.Add("ptNum", CStr(pn))
            End With

            If ct > 1 Then
                Stop
                fn = ""
                For Each df In .ComponentDefinitions
                    PartDoc = aiDocument(.Document)
                    ck = aiDocPartNum(PartDoc)
                    If ck = pn Then
                        fn = fn & vbCrLf & PartDoc.FullDocumentName
                    Else
                        Stop
                    End If
                Next
            End If

            rt.Add("aiDoc", CStr(fn))
        End With

        bomItemInfo = rt
    End Function

    Public Function dBVg7f4(
    aiAssy As Inventor.AssemblyDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' not sure what doing with this one
        ' further developemt on hold
        '
        'Dim PartDesc As Inventor.AssemblyDocument
        Dim rt As Scripting.Dictionary

        Dim BillVeiw As Inventor.BOMView
        Dim BillRow As Inventor.BOMRow

        If dc Is Nothing Then
            rt = dBVg7f4(aiAssy, New Scripting.Dictionary)
        ElseIf aiAssy Is Nothing Then
            rt = dc
        Else
            With aiAssy
                BillVeiw = .ComponentDefinition.BOM.BOMViews.Item("Structured")
                With BillVeiw
                    Stop
                End With
            End With
        End If
        dBVg7f4 = rt
    End Function

    Public Function bomInfoBkDn(
    BillRow As Inventor.BOMRowsEnumerator,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional fn As String = ""
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary

        Dim pn As String
        Dim ck As String

        If dc Is Nothing Then
            rt = bomInfoBkDn(BillRow, New Scripting.Dictionary, fn)
        Else
            rt = dc
            If BillRow Is Nothing Then 'likely a Part
                'might need/want to look for raw material here
            Else 'it's an Assembly
                For Each BillRow In BillRow
                    'Stop
                    System.Windows.Forms.Application.DoEvents()
                    wk = bomItemInfo(BillRow)
                    wk.Add("ptOf", fn)
                    pn = CStr(wk.Item("ptNum"))
                    ck = fn & "|" & pn
                    With rt
                        If .Exists(ck) Then
                            'Stop
                            ''Debug.Print(ConvertToJson(dcCmpTextOf2dc(wk,dcOb(.Item(ck))),vbTab)
                            ''Debug.Print(ConvertToJson(
                            With DcWBQbyCmpResult(DcCmpTextOf2dc(wk, dcOb(.Item(ck)))) ',vbTab)
                                With dcOb(.Item("!="))
                                    .Remove("path")
                                    .Remove("base")
                                    If .Count > 0 Then
                                        Debug.Print("MISMATCH: ", TxDumpLs(.Keys, ", "))
                                        Stop
                                    End If
                                End With
                            End With
                        Else
                            .Add(ck, wk)
                        End If
                    End With

                    With BillRow
                        If .ChildRows Is Nothing Then
                            dc = dBVg1f2(
                            ThisApplication.Documents.ItemByName(
                                wk.Item("aiDoc")
                            ), wk, dc)
                        Else
                            System.Windows.Forms.Application.DoEvents()
                            dc = bomInfoBkDn(
                            .ChildRows, dc, pn
                        )
                        End If
                    End With
                    System.Windows.Forms.Application.DoEvents()
                Next
            End If
        End If
        bomInfoBkDn = rt
        'Debug.Print(txDumpLs(bomInfoBkDn(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows).Keys)
        'send2clipBd ConvertToJson(bomInfoBkDn(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows), vbTab)
    End Function

    Public Function dcOfBomsFromAiStructured(
    BillRow As Inventor.BOMRowsEnumerator,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional fn As String = ""
) As Scripting.Dictionary
        '
        ' dcOfBomsFromAiStructured --
        '     generate Dictionary of BOMs:
        '     one for each distinct Assembly in
        '     supplied Inventor BOM (structured)
        '
        '     returned as Dictionary of Assembly
        '     sub Dictionaries, each keyed to its
        '     Part Number and containing a  of
        '     Item sub Dictionaries, again keyed
        '     to Item P/N. Each Item sub Dictionary
        '     represents a BOM line item
        '
        Dim rt As Scripting.Dictionary

        dcOfBomsFromAiStructured =
        dBV0g0f4(dBV0g0f3(dBV0g0f1(
        BillRow, dc, fn
    )))
        Debug.Print("") 'Breakpoint Landing
        ''Debug.Print(ConvertToJson(dBV0g0f2(dcOb(dcOb(rt.Item("19-240-79925")).Item("19-240-90004"))), vbTab)
        'Debug.Print(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows).Count
    End Function

    Public Function dBV0g0f1(
    BillRow As Inventor.BOMRowsEnumerator,
    Optional dc As Scripting.Dictionary = Nothing,
    Optional fn As String = ""
) As Scripting.Dictionary
        '
        ' dBV0g0f1 -- retrieve BOM data
        '     from a BOMRowsEnumerator
        '     and its child row enumerators
        '
        Dim rt As Scripting.Dictionary
        Dim PartDesc As Scripting.Dictionary
        Dim it As Scripting.Dictionary
        Dim dt As Scripting.Dictionary
        Dim fd As Scripting.Dictionary
        '
        Dim kyIt As Object
        '

        Dim pn As String
        Dim th As String

        If dc Is Nothing Then
            rt = dBV0g0f1(BillRow,
        New Scripting.Dictionary _
        , fn)
        Else
            rt = dc
            If BillRow Is Nothing Then
                'Stop
                Debug.Print("") 'Breakpoint Landing
            Else

                With rt
                    If Not .Exists(fn) Then
                        .Add(fn, New Scripting.Dictionary)
                    End If
                    PartDesc = .Item(fn)
                End With

                For Each BillRow In BillRow
                    System.Windows.Forms.Application.DoEvents()

                    dt = bomItemInfo(BillRow)
                    With dt
                        pn = CStr(.Item("ptNum"))
                        th = CStr(.Item("path"))
                    End With

                    With PartDesc
                        If Not .Exists(pn) Then
                            .Add(pn, New Scripting.Dictionary)
                        End If
                        it = .Item(pn)
                    End With

                    With it
                        If .Exists(th) Then
                            Stop
                        Else
                            .Add(th, dt)
                        End If
                    End With

                    rt = dBV0g0f1(BillRow.ChildRows, rt, pn)
                    System.Windows.Forms.Application.DoEvents()
                Next
            End If
        End If

        'If Len(fn) = 0 Then Stop
        dBV0g0f1 = rt
    End Function

    Public Function dBV0g0f2(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim dcField As Scripting.Dictionary
        Dim k0path As Object
        Dim k1Field As Object
        Dim itValue As Object

        rt = New Scripting.Dictionary
        With dc : For Each k0path In .Keys
                With dcOb(.Item(k0path))
                    For Each k1Field In .Keys
                        itValue = .Item(k1Field)

                        With rt
                            If Not .Exists(k1Field) Then
                                .Add(k1Field, New Scripting.Dictionary)
                            End If
                            dcField = .Item(k1Field)
                        End With

                        With dcField
                            If Not .Exists(itValue) Then
                                .Add(itValue, New Scripting.Dictionary)
                            End If

                            With dcOb(.Item(itValue))
                                If .Exists(k0path) Then
                                Else
                                    .Add(k0path, 1)
                                End If
                            End With
                        End With
                    Next : End With
            Next : End With

        dBV0g0f2 = rt
    End Function

    Public Function dBV0g0f3(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dBV0g0f3 -- summarize BOM line item fields
        '
        Dim rt As Scripting.Dictionary
        Dim dcProd As Scripting.Dictionary
        Dim k0Prod As Object
        Dim k1Item As Object

        rt = New Scripting.Dictionary

        With dc : For Each k0Prod In .Keys
                With rt
                    .Add(k0Prod, New Scripting.Dictionary)
                    dcProd = .Item(k0Prod)
                End With

                With dcOb(.Item(k0Prod))
                    For Each k1Item In .Keys
                        dcProd.Add(k1Item,
            dBV0g0f2(.Item(k1Item)))
                    Next : End With
            Next : End With

        dBV0g0f3 = rt
    End Function

    Public Function dBV0g0f4(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dBV0g0f4 -- reduce results of dBV0g0f3
        '     to single values per field
        '     for each Item under each Product
        '
        Dim rt As Scripting.Dictionary
        Dim dcProd As Scripting.Dictionary
        Dim dcItem As Scripting.Dictionary
        Dim k0Prod As Object
        Dim k1Item As Object
        Dim k2Feld As Object

        rt = New Scripting.Dictionary

        With dc : For Each k0Prod In .Keys 'Products
                With rt
                    .Add(k0Prod, New Scripting.Dictionary)
                    dcProd = .Item(k0Prod)
                End With

                With dcOb(.Item(k0Prod)) 'Items
                    For Each k1Item In .Keys
                        With dcProd
                            .Add(k1Item, New Scripting.Dictionary)
                            dcItem = .Item(k1Item)
                        End With

                        With dcOb(.Item(k1Item)) 'Fields
                            On Error Resume Next
                            .Remove("path")
                            .Remove("base")
                            On Error GoTo 0

                            For Each k2Feld In .Keys
                                With dcOb(.Item(k2Feld)) 'Value(s)
                                    If .Count > 1 Then
                                        Stop
                                    Else
                                        dcItem.Add(k2Feld,
                        .Keys(0)) '.Item()
                                    End If : End With
                            Next
                        End With
                    Next : End With
            Next : End With

        dBV0g0f4 = rt
    End Function

    Public Function dBV0g0f5(
    dc As Scripting.Dictionary,
    Optional dlm As String = "|"
) As Scripting.Dictionary
        '
        ' dBV0g0f5 -- reduce results of dBV0g0f3
        '     to single values per field
        '     for each Item under each Product
        '
        Dim rt As Scripting.Dictionary
        Dim dcItem As Scripting.Dictionary
        Dim k0Prod As Object
        Dim k1Item As Object
        Dim k2Feld As Object
        Dim BillRow As String
        Dim co As String

        rt = New Scripting.Dictionary

        With dc : For Each k0Prod In .Keys 'Products
                If Len(k0Prod) > 0 Then
                    With dcOb(.Item(k0Prod)) 'Items
                        For Each k1Item In .Keys
                            dcItem = .Item(k1Item)

                            With dcItem
                                BillRow = k0Prod & dlm & k1Item
                                For Each k2Feld In {
                        "seq", "qty", "qtUnit"}

                                    If .Exists(k2Feld) Then
                                        co = CStr(.Item(k2Feld))
                                    Else
                                        co = ""
                                    End If

                                    BillRow = BillRow & dlm & co
                                Next
                                Debug.Print("") 'Breakpoint Landing
                            End With

                            rt.Add(BillRow, dcItem)
                        Next : End With
                Else 'top-level assembly
                    'skip (for now)
                End If
            Next : End With

        dBV0g0f5 = rt
        'send2clipBdWin10 txDumpLs(dBV0g0f5(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows)).Keys)
    End Function

    Public Function csvOfBomsFromDc(
    dc As Scripting.Dictionary,
    Optional dlm As String = "|"
) As String
        'Product|Item|ItemOrder|QuantityInConversionUnit|ConversionUnit
        'NOTE[2021.08.20]: want to change 'Item'
        '   to 'ItemCode' for compatibility with
        '   current Genius BOM import format.
        '   Will hold off for now.
        csvOfBomsFromDc = Join({
        "Product", "Item", "ItemOrder",
        "QuantityInConversionUnit",
        "ConversionUnit"}, dlm) & vbCrLf & TxDumpLs(dBV0g0f5(dc, dlm).Keys)
        'send2clipBdWin10 csvOfBomsFromDc(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows))
    End Function

    Public Function csvOfBomsFromAiStructured(
    AiDoc As Inventor.Document,
    Optional dlm As String = "|"
) As String
        csvOfBomsFromAiStructured =
        csvOfBomsFromDc(
        dcOfBomsFromAiStructured(
        bomViewStruct(aiDocAssy(
        aiDocActive())).BOMRows
    ), dlm)
        'send2clipBdWin10 csvOfBomsFromAiStructured(aiDocActive())
    End Function

    '
    '
    '
    Private Function dvlBomView() As String
        dvlBomView = "dvlBomView"
    End Function

    '        ''Debug.Print(ConvertToJson(dcCmpTextOf2dc(fd,dcOb(.Item(dt))),vbTab)
    '        ''Debug.Print(ConvertToJson(
    '        With dcWBQbyCmpResult(dcCmpTextOf2dc(fd, dcOb(.Item(dt)))) ',vbTab)
    '            With dcOb(.Item("!="))
    '                .Remove("path"
    '                .Remove("base"
    '                If .Count > 0 Then
    '                    Debug.Print("MISMATCH: ", txDumpLs(.Keys, ", ")
    '                    Stop
    '                End If
    '            End With
    '        End With
    '    Else
    '        .Add(dt, fd

    'With BillRow
    '    If .ChildRows Is Nothing Then
    '         dc = dBVg1f2( _
    '            ThisApplication.Documents.ItemByName( _
    '                fd.Item("aiDoc") _
    '            ), fd, dc)
    '    Else
    '       Application.DoEvents()
    '         dc = bomInfoBkDn( _
    '            .ChildRows, dc, pn _
    '        )
    '    End If
    'End With


End Module