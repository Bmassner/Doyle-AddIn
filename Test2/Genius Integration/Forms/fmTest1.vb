Class FmTest1
    Private cn As ADODB.Connection
    Private rsFam As ADODB.Record
    Private rsPrt As ADODB.Record
    Private rsItm As ADODB.Record

    Private dc As Scripting.Dictionary
    Private ad As Inventor.Document

    Private PropertySetDsn As Inventor.Property
    Private PropertySetUsr As Inventor.Property

    Private dcDsn As Scripting.Dictionary
    Private dcUsr As Scripting.Dictionary

    Private FamilyProp As Inventor.Property
    Private StockProp As Inventor.Property
    Private ThicknessProp As Inventor.Property

    Public Function AskAbout(
    AiDoc As Inventor.Document,
    Optional txMsg As String = ""
) As VbMsgBoxResult
        Dim pc As stdole.IPictureDisp
        Dim ck As VbMsgBoxResult
        Dim pn As String    'part number
        Dim StockNumb As String    'material (stock) number
        Dim StockFam As String    'material (stock) family
        Dim PartDesc As String    'part description
        Dim df As Single

        ad = AiDoc
        With ad
            On Error Resume Next
            Err.Clear()
            pc = .Thumbnail
            If Err.Number = 0 Then 'we're good
            Else 'no image
            End If
            On Error GoTo 0

            PropertySetDsn = .Propertys(gnDesign)
            PropertySetUsr = .Propertys(gnCustom)

            dcDsn = dcAiPropsIn(PropertySetDsn)
            dcUsr = dcAiPropsIn(PropertySetUsr)

            FamilyProp = PropertySetDsn.Item(pnFamily)

            ''  Get Sheet Metal Thickness Property
            ThicknessProp = aiPropShtMetalThickness(ad)
            ''  NOTE: Function returns Nothing
            ''      if Part is NOT Sheet Metal!

            With dcUsr
                If .Exists(pnRawMaterial) Then
                    StockProp = PropertySetUsr.Item(pnRawMaterial)
                Else
                    On Error Resume Next
                    Err.Clear()
                    StockProp = PropertySetUsr.Add("", pnRawMaterial)
                    If Err.Number Then
                        Stop
                    Else
                        .Add(pnRawMaterial, StockProp)
                    End If
                    On Error GoTo 0
                End If
            End With

            ' REV[2022.04.28.1615]
            ' added initializtion of Dictionary dc
            ' with initial raw material ting.
            ' StockNumb now assigned from the Dictionary.
            ' NOTE: probably want to  initial
            ' values in a separate "recovery"
            ' Dictionary to be restored if
            ' the User chooses to cancel.
            ' Also, see function/method dcUpd.
            ' looks like it gets called when
            ' something changes. Easy to miss!
            dc.Item(pnRawMaterial) = StockProp.Text
            StockNumb = dc.Item(pnRawMaterial)
            pn = PropertySetDsn.Item(pnPartNum).Text
            PartDesc = PropertySetDsn.Item(pnDesc).Text
        End With

        With Me
            .Text = "Please Review Part Number: " & pn

            If pc Is Nothing Then
            Else
                .imThmNail.Image = pc
            End If

            With .lbMsg
                .Text = pn & ": " & PartDesc _
                & vbCrLf & txMsg & IIf(Len(txMsg) > 0, vbCrLf, "") _
                & ft1g0f0(pnCatWebLink, PropertySetDsn.Item(pnCatWebLink)) & vbCrLf _
                & ft1g0f0(pnMaterial, PropertySetDsn.Item(pnMaterial)) & vbCrLf _
                & ft1g0f0(pnThickness, ThicknessProp) & vbCrLf _
                & "" _
                & vbCrLf & pnThickness & ": " & PropertySetUsr.Item(pnThickness).Text
                '
            End With
            df = mdl1g1f2(.lbMsg)
            'If df > 0 Then
            '    mdl1g1f3.lbMtFamily(, 0, df)
            '    mdl1g1f3.LbxFamily(, 0, df)
            'End If

            .dbFamily.Text = FamilyProp.Text

            If Len(StockNumb) > 0 Then
                With cn.Execute(
                "select Family from vgMfiItems where Item = '" _
                & Replace(StockNumb, "'", "''") & "'"
            )
                    ' REV[2022.08.19.1359]
                    ' temporarily replacing direct use of StockNumb
                    ' with call to Replace single quotes
                    ' in string with doubled single quotes
                    ' (NOT double quotes!) to "escape" the
                    ' character in a string value.
                    ' '
                    ' will ultimately want to produce some
                    ' sort of 'handler' to preprocess values
                    ' for use in SQL commands to avoid errors
                    ' that arise from this sort of thing.
                    If .BOF Or .EOF Then
                        StockFam = ""
                    Else
                        StockFam = .Fields(0).Value
                    End If
                    ' NOTE[2022.04.28.1625]
                    ' THOUGHT Material Family was also
                    ' added to Dictionary dc, but believe
                    ' that's actually the PART Family.
                End With

                If Len(StockFam) = 0 Then 'selected Material
                    'EITHER doesn't have a Family,
                    'OR is not (yet) in Genius.
                    'SO, let's just ...
                    StockFam = "DSHEET" 'as a default!
                Else 'no need to do anything
                    'comments below are from when
                    'this was the NO family block.
                    'retain for now, until we're sure
                    'this are working all right.

                    'this SHOULDN'T happen, so hopefully
                    'things won't come to this...
                    'Stop
                    'worry about the handler later
                End If

                On Error Resume Next
                Err.Clear()
                .LbxFamily.Text = StockFam
                If Err.Number Then
                    Debug.Print("FAILED TO  MATERIAL FAMILY " & StockFam)
                    '    ck = MsgBox(Join({
                    '    "Part Number " & pn, "uses Material " & StockNumb _
                    '    , "which is a" & IIf(
                    '        InStr(1, "AEIOU", UCase$(Left$(StockFam, 1))),
                    '        "n ", " "
                    '    ) & StockFam & " Item." _
                    '    , "" _
                    '    , "This interface does not presently" _
                    '    , "support Materials from this Family." _
                    '    , "" _
                    '    , "You might not be able to find the correct" _
                    '    , "Material for this Part, and might wish" _
                    '    , "to avoid changing it here." _
                    '    , "" _
                    '    , "Do you wish to proceed anyway?"
                    '}, vbCrLf),
                    '    vbYesNoCancel + vbExclamation + vbDefaultButton2,
                    '    "Material Family not Supported"
                    ')
                    If ck = vbCancel Then
                        Stop
                    End If
                Else
                    Err.Clear()
                    .LbxItem.Text = StockNumb

                    ' REV[2022.05.06.1329]
                    ' added intermediate error handler
                    ' to capture failure in Material
                    ' Family selector to adopt new Value.
                    ' it re-implements process of Event
                    ' handler Sub LbxFamily_Change
                    ' against variable 'StockFam' directly
                    ' in an effort to force population
                    ' of Material list.
                    If Err.Number Then
                        Debug.Print("") 'Breakpoint Landing
                        Err.Clear()
                        rsItm.Filter = "Family = '" & StockFam & "'"
                        .LbxItem.Text = m0g3f1(rsItm)
                        .LbxItem.Text = StockNumb
                    End If
                    ' something MIGHT have happened
                    ' to prevent normal Value update
                    ' when LbxFamily is  above.
                    ' further investigation may be
                    ' warranted.

                    If Err.Number Then
                        Debug.Print("FAILED TO  MATERIAL " & StockNumb)
                        ck = MsgBox(Join({"!!WARNING!!", "" _
                        , "Active Material " & StockNumb _
                        , "for Part Number " & pn _
                        , "could NOT be selected," _
                        , "and might be unavailable." _
                        , "" _
                        , "You might wish to avoid" _
                        , "making Material changes" _
                        , "to this Part here." _
                        , "" _
                        , "Do you wish to proceed anyway?"
                    }, vbCrLf),
                        vbYesNoCancel + vbExclamation,
                        "Active Material Not Found!"
                    )
                        If ck = vbCancel Then
                            Stop
                        End If
                    Else
                        ck = vbYes
                        LbxItem_Change()
                        'LbxFamily_Change
                        rsItm.Filter = "Family = '" & StockFam & "'"
                        .LbxItem.Text = m0g3f1(rsItm)
                    End If
                End If
                On Error GoTo 0
            Else
                ck = vbYes
            End If

            If ck = vbYes Then
                .Show()
            End If
        End With
        AskAbout = ck 'vbYes ' = 1
    End Function

    Private Function ft1g0f0(
    pn As String, InvProperty As Inventor.Property
) As String
        If InvProperty Is Nothing Then
            ft1g0f0 = ""
        Else
            ft1g0f0 = vbCrLf & pn & ": " & InvProperty.Text
        End If
    End Function

    Private Sub dbFamily_Change()
        Debug.Print(dcUpd(pnFamily, dbFamily.Text))
    End Sub
    'Me.LbxItem.ColumnWidths = "84 PartDoc;6 PartDoc;180 PartDoc"
    'Me.LbxItem.ColumnWidths = "84 PartDoc;48 PartDoc;216 PartDoc"

    Private Sub lbMsg_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        Stop
    End Sub

    Private Sub LbxFamily_Change()
        With Me
            rsItm.Filter = "Family = '" & .LbxFamily.Text & "'"
            .LbxItem.Text = m0g3f1(rsItm)
        End With
    End Sub

    Public Function ItemData() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                rt.Add(ky, .Item(ky))
            Next
        End With
        ItemData = rt
    End Function

    Public Function Synch() As Scripting.Dictionary
        With dc
            If .Exists(pnFamily) Then FamilyProp.Text = dc.Item(pnFamily)
            If .Exists(pnRawMaterial) Then StockProp.Text = dc.Item(pnRawMaterial)
        End With

        Synch = Me.ItemData
    End Function

    Private Function dcUpd(ky As String, vl As Object) As String
        Dim rt As String

        If IsNothing(vl) Then
            dcUpd = dcUpd(ky, "")
        Else
            With dc
                If .Exists(ky) Then
                    rt = CStr(.Item(ky))
                    .Item(ky) = vl
                    dcUpd = "CHANGE[" & ky & "] FROM '" & rt _
                & "' TO '" & CStr(.Item(ky)) & "'"
                Else
                    .Add(ky, vl)
                    dcUpd = "[" & ky & "] TO '" _
                & CStr(.Item(ky)) & "'"
                End If
            End With
        End If
    End Function

    Private Sub LbxItem_Change()
        Debug.Print(dcUpd(pnRawMaterial, LbxItem.Text))
    End Sub

    Private Sub UserForm_Initialize()
        dc = New Scripting.Dictionary
        cn = CnGnsDoyle()

        ' Uncomment and fix if needed: rsFam = .Execute(...)
        ' With cn
        '     Set rsFam = .Execute(Join(Array( _
        '         "select Family, Description1", _
        '         "from vgMfiFamilies", _
        '         "order by Family" _
        '     ), " "))
        ' End With

        With cn
            rsPrt = .Execute(Join({(
                "select Family, FamilyGroup, Description1",
                "from vgMfiFamilies",
                "order by Family"
            ), vbCrLf}))
            rsItm = .Execute(Join({(
                "Select I.Item, I.Family, I.Description1, I.Specification1",
                "From vgMfiItems as I",
                "Inner Join vgMfiFamilies as F",
                "On I.Family = F.Family",
                "Where F.FamilyGroup = 'RAW'",
                "order by Family, Item"
            ), " "}))
        End With

        With Me
            rsPrt.Filter = "FamilyGroup = 'RAW'"
            .LbxFamily.Text = m0g3f1(rsPrt) 'rsFam

            rsPrt.Filter = "FamilyGroup = 'PARTS'"
            .dbFamily.Text = m0g3f1(rsPrt)
        End With
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        Cancel = 1
        Me.Hide()
    End Sub

    Private Sub UserForm_Terminate()
        cn.Close()
        cn = Nothing
    End Sub

End Class