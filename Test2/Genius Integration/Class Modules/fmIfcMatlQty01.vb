Imports System.Windows
Imports System.Windows.Forms

Class FmIfcMatlQty01

    Private WithEvents Fm As FmMatlQty
    Private WithEvents LbxMatlQty As MSForms.TextBox
    Private WithEvents txbMatlQty As MSForms.TextBox
    Private WithEvents cbxUnitQty As MSForms.ComboBox
    Private WithEvents imgThmNail As Forms.PictureBox

    'Private WithEvents cmdOK        As MSForms.CommandButton
    'Private WithEvents cmdCancel    As MSForms.CommandButton
    'Private WithEvents cmdOK        As MSForms.CommandButton

    Private LblPartNumber As MSForms.Label
    Private lblPartInfo As MSForms.Label
    Private lblMatlNumber As MSForms.Label
    Private lblMatlInfo As MSForms.Label
    'lblMatlQty
    'lblNoImg

    'imThmNail

    Private dcResult As Scripting.Dictionary

    'Private dcGiven As Scripting.Dictionary
    'Private dcWorkg As Scripting.Dictionary
    'Private FmStatus As VbMsgBoxResult

    Private Const FmVersion As String = "Material Quantity Form Interface 0.0.0.0 [2022.03.04]"
    '
    '
    '

    Private Sub Class_Initialize()
        'Dim ctl As MSForms.Control

        dcResult = New Scripting.Dictionary
        With dcResult
            .Add(pnRmQty, 0)
            .Add(pnRmUnit, "")
        End With

        Fm = New FmMatlQty
        With Fm
            'For Each ctl In .Controls
            '    Debug.Print(ctl.Name
            'Next

            cbxUnitQty = .cbxUnitQty
            cbxUnitQty.Text = Join(Split("IN FT FT2 IN2 EA"), ", ")

            LbxMatlQty = .LbxMatlQty
            txbMatlQty = .txbMatlQty

            imgThmNail = .imThmNail
            LblPartNumber = .LblPartNumber
            lblPartInfo = .lblPartInfo
            lblMatlNumber = .lblMatlNumber
            lblMatlInfo = .lblMatlInfo
        End With
    End Sub

    Public Function Result() As Scripting.Dictionary
        Result = DcCopy(dcResult)
    End Function

    Private Function Changes(
    wkg As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With wkg : For Each ky In dcResult.Keys
                If .Exists(ky) Then
                    If .Item(ky) = dcResult.Item(ky) Then
                        'no difference, skip it
                    Else
                        rt.Add(ky, .Item(ky))
                    End If
                Else 'not , not sure what to do here
                    'rt.Add(ky, Empty
                    'don't really like this option
                    'so leaving it disabled for now
                End If
            Next : End With

        Changes = DcCopy(dcResult)
    End Function

    Private Function Commit(
    src As Scripting.Dictionary
) As Scripting.Dictionary
        Dim ky As Object

        With dcResult
            For Each ky In .Keys
                If src.Exists(ky) Then
                    .Item(ky) = src.Item(ky)
                End If
            Next : End With

        Commit = DcCopy(dcResult)
    End Function

    Public Function SeeUser(
    Optional About As Object = Nothing
) As Scripting.Dictionary 'fmIfcMatlQty01
        Dim ky As String
        Dim ck As String

        If About Is Nothing Then
            SeeUser = SeeUser(NuDcPopulator(
            ).Setting(pnRmQty & "()",
                NuDcPopulator(
                    ).Setting(4, 1
                    ).Setting(2, 1
                    ).Setting(24, 1
                ).Dictionary()
            ).Setting(pnRmQty, 24
            ).Setting(pnRmUnit, "IN"
            ).Setting(pnPartNum, "NO-ITM-GIVEN"
            ).Setting(pnRawMaterial, "NO-MTL-GIVEN"
        ).Dictionary())
        ElseIf TypeOf About Is Scripting.Dictionary Then
            SeeUser = SeeUserWithDict(About)
        ElseIf TypeOf About Is Inventor.PartDocument Then
            SeeUser = SeeUserWithPart(About)
        ElseIf TypeOf About Is Inventor.Property Then
            SeeUser = SeeUserWithQtyProp(About)
        Else
            SeeUser = SeeUser()
        End If
    End Function

    ' make this one Public later
    ' once Part version is working
    Private Function SeeUserWithModel(
    About As Inventor.Property
) As FmIfcMatlQty01
    End Function

    Public Function SeeUserWithPart(
    About As Inventor.PartDocument
) As Scripting.Dictionary 'fmIfcMatlQty01
        If About Is Nothing Then
            SeeUserWithPart = SeeUser(About)
        Else
            Dim dcPr As Scripting.Dictionary
            Dim obPr As Inventor.Property
            Dim kyPr As Object
            Dim op As Long

            dcPr = New Scripting.Dictionary
            With About.Propertys
                With .Item(gnDesign)
                    dcPr.Add(pnPartNum, .Item(pnPartNum).Text)
                    dcPr.Add(pnDesc, .Item(pnDesc).Text)
                End With

                On Error Resume Next
                With .Item(gnCustom)
                    For Each kyPr In {
                    pnRawMaterial, pnRmQty, pnRmUnit}

                        Err.Clear()
                        obPr = .Item(CStr(kyPr))
                        If Err.Number = 0 Then
                            dcPr.Add(kyPr, obPr.Text)
                        Else
                            Debug.Print(Err.Description)
                            Stop
                            Err.Clear()
                        End If
                    Next
                End With
                On Error GoTo 0
            End With

            '
            ' prepare Dictionary of Dimensions
            ' with Count of Each
            '
            Dim dcDm As Scripting.Dictionary
            Dim vlDm As Object
            Dim ctDm As Long

            dcDm = New Scripting.Dictionary

            With nuAiBoxData().UsingInches(1)
                For op = 0 To 1
                    With .UsingModel(About, op)
                        For Each vlDm In {
                    System.Math.Round(.SpanX, 4),
                    System.Math.Round(.SpanY, 4),
                    System.Math.Round(.SpanZ, 4),
                0}
                            With dcDm
                                If vlDm > 0 Then
                                    If .Exists(vlDm) Then
                                        ctDm = .Item(vlDm) + 1
                                        .Item(vlDm) = ctDm
                                    Else
                                        .Add(vlDm, ctDm)
                                    End If
                                End If : End With : Next
                    End With : Next
            End With

            With dcPr
                .Add(pnRmQty & "()", dcDm)
                .Add("img", About.Thumbnail)
            End With

            SeeUserWithPart _
        = SeeUserWithDict(dcPr)
        End If
    End Function

    Private Function SeeUserWithQtyProp(
    About As Inventor.Property
) As Scripting.Dictionary 'fmIfcMatlQty01
        '
        ' this one will have to be heavily modified
        ' likely dumping a bunch of code now implemented
        ' in SeeUserWithPart, which can simply be
        ' called with the Document containing
        ' the supplied Property
        '

        If About Is Nothing Then
            Stop
        Else
            ' these variables are for use
            ' in separating quantity from
            ' unit of measure in Value of
            ' supplied Property
            Dim vlIn As String
            Dim arIn As Object
            Dim qtIn As Double
            Dim unIn As String
            ' split incoming Property Value into
            ' Quantity and Unit of Measurement
            vlIn = CStr(About.Text) & " "
            ' note: concatenated space at end
            ' of Value text should ensure two
            ' members of arIn, as follows
            arIn = Split(vlIn, " ", 2)

            qtIn = System.Math.Round(Val(arIn(0)), 4)
            If UBound(arIn) > 0 Then
                unIn = Trim$(arIn(1))
            End If
            ' this section and its associated variables
            ' will likely be exported to a separate function

            ' force blank Unit of
            ' Measure to default inches
            If Len(unIn) = 0 Then unIn = "IN"

            ' the following section SHOULD be
            ' implemented now in SeeUserWithPart
            ' it should be possible to simply
            ' call that function, completely
            ' ignoring the supplied Property
            '
            ' prepare Dictionary of Dimensions
            ' with Count of Each
            '
            Dim dcDm As Scripting.Dictionary
            Dim vlDm As Object
            Dim ctDm As Long

            dcDm = New Scripting.Dictionary
            If qtIn > 0 Then dcDm.Add(qtIn, 1)

            '
            ' get all necessary information
            ' from Inventor Model
            '
            Dim md As Inventor.Document
            Dim mdPt As Inventor.Property
            Dim mdMt As Inventor.Property

            md = aiDocument(About.Parent.Parent.Parent)
            With md.Propertys
                mdPt = .Item(gnDesign).Item(pnPartNum)
                On Error Resume Next
                Err.Clear()
                mdMt = .Item(gnCustom).Item(pnRawMaterial)
                If Err.Number = 0 Then
                Else
                    Stop
                End If
                On Error GoTo 0
            End With

            With nuAiBoxData(
        ).UsingInches(1
        ).UsingModel(About)
                For Each vlDm In {
                System.Math.Round(.SpanX, 4),
                System.Math.Round(.SpanY, 4),
                System.Math.Round(.SpanZ, 4),
            0}
                    With dcDm
                        If vlDm > 0 Then
                            If .Exists(vlDm) Then
                                ctDm = .Item(vlDm) + 1
                                .Item(vlDm) = ctDm
                            Else
                                .Add(vlDm, ctDm)
                            End If
                        End If : End With : Next
            End With

            With NuDcPopulator(
            ).Setting(pnRmQty & "()", dcDm
            ).Setting(pnRmQty, qtIn
            ).Setting(pnRmUnit, unIn
            ).Setting(pnPartNum, mdPt.Text
            ).Setting(pnRawMaterial, mdMt.Text
        )
                SeeUserWithQtyProp _
            = SeeUserWithDict(
            .Dictionary()
        )
            End With
        End If
    End Function

    Public Function SeeUserWithDict(
    About As Scripting.Dictionary
) As Scripting.Dictionary 'fmIfcMatlQty01
        Dim ky As String
        Dim ck As String

        If About Is Nothing Then
            SeeUserWithDict = SeeUserWithDict(NuDcPopulator(
            ).Setting(pnRmQty & "()",
                NuDcPopulator(
                    ).Setting(4, 1
                    ).Setting(2, 1
                    ).Setting(24, 1
                ).Dictionary()
            ).Setting(pnRmQty, 24
            ).Setting(pnRmUnit, "IN"
            ).Setting(pnPartNum, "NO-ITM-GIVEN"
            ).Setting(pnRawMaterial, "NO-MTL-GIVEN"
        ).Dictionary())
        Else
            With About
                '.Add("img", About.Thumbnail
                If .Exists("img") Then
                    imgThmNail.Image _
                = .Item("img")
                End If

                If .Exists(pnDesc) Then
                    txbMatlQty.Text = Val(
                CStr(.Item(pnDesc)))
                End If

                ky = pnRmQty & "()"
                If .Exists(ky) Then
                    LbxMatlQty.Text =
                dcOb(.Item(ky)).Keys
                End If

                If .Exists(pnRmQty) Then
                    txbMatlQty.Text = Val(
                CStr(.Item(pnRmQty)))
                End If

                If .Exists(pnRmUnit) Then
                    On Error Resume Next

                    Err.Clear()
                    cbxUnitQty.Text = .Item(pnRmUnit)
                    If Err.Number Then
                        Debug.Print("") 'Breakpoint Landing
                        cbxUnitQty.Text = "IN"
                    End If
                    On Error GoTo 0
                End If

                '
                ' Following are "boilerplate" elements
                ' for Part/Item and Raw Material numbers,
                ' along with their descriptions.
                '
                ' A thumbnail image of the Part is also
                ' expected to be supplied at some point,
                ' but will be held off for now, pending
                ' successful testing of the form's main
                ' functions.
                '
                ' Part/Item Number
                If .Exists(pnPartNum) Then
                    LblPartNumber.Text _
                = CStr(.Item(pnPartNum))
                End If

                ' Material Number
                If .Exists(pnRawMaterial) Then
                    lblMatlNumber.Text _
                = CStr(.Item(pnRawMaterial))
                End If

                ' Item Description
                If .Exists(pnDesc) Then
                    lblPartInfo.Text _
                = CStr(.Item(pnDesc))
                End If

                ' Material Description
                ' (not expected at this time)
                ky = pnRawMaterial & ":"
                If .Exists(ky) Then
                    lblMatlInfo.Text _
                = CStr(.Item(ky))
                End If

                'imThmNail
            End With

            With Commit(About)
            End With

            Fm.Show()
            'Stop

            With NuDcPopulator(
        ).Setting(pnRmQty,
            System.Math.Round(Val(
            txbMatlQty.Text
            ), 4)
        ).Setting(pnRmUnit,
            cbxUnitQty.Text
        ) 'Mapping...
                ' txbMatlQty -> pnRmQty
                ' cbxUnitQty -> pnRmUnit

                SeeUserWithDict = Commit(.Dictionary)
            End With
        End If
    End Function

    Public Function Version()
        Version = FmVersion
        'fmStatus=vbRetry
    End Function

    Private Sub Class_Terminate()
        ' dcWorkg = Nothing
        ' dcGiven = Nothing

        imgThmNail.Image = Nothing
        imgThmNail = Nothing

        cbxUnitQty = Nothing

        LbxMatlQty = Nothing
        txbMatlQty = Nothing

        LblPartNumber = Nothing
        lblPartInfo = Nothing
        lblMatlNumber = Nothing
        lblMatlInfo = Nothing

        Fm = Nothing
    End Sub

    Private Sub Fm_Sent(Signal As VbMsgBoxResult)
        Dim ck As VbMsgBoxResult

        If Signal = vbCancel Then
            ck = MsgBox(Join({
            "Material Quantity",
            "and Units will",
            "remain unchanged."
        }, vbCrLf), vbYesNo,
            "Cancel Update?"
        )
            'Stop

            If ck = vbYes Then
                With dcResult
                    txbMatlQty.Text = CStr(.Item(pnRmQty))
                    cbxUnitQty.Text = .Item(pnRmUnit)
                End With
                Fm.Hide()
            ElseIf ck = vbCancel Then
                Stop 'drop and debug
                'NOTE: Without Cancel Button
                'available on MsgBox, this
                'option won't be accessible.

                'a proposed "debug" mode that
                'would add the Cancel Button to
                'the MsgBox(has not yet been
                'implemented, but might in future.
            End If
            Debug.Print("") 'Breakpoint Landing
        ElseIf Signal = vbOK Then
            ck = MsgBox(Join({
            "Update Material",
            "Quantity to " _
            & CStr(System.Math.Round(Val(
                txbMatlQty.Text
            ), 4)) _
            & cbxUnitQty.Text & "?"
        }, vbCrLf), vbYesNo,
            "Update Quantity?"
        )
            'Stop

            If ck = vbYes Then
                Fm.Hide()
                Debug.Print("") 'Breakpoint Landing
            ElseIf ck = vbCancel Then
                Stop 'drop and debug
                'NOTE: Without Cancel Button
                'available on MsgBox, this
                'option won't be accessible.

                'a proposed "debug" mode that
                'would add the Cancel Button to
                'the MsgBox(has not yet been
                'implemented, but might in future.
            End If
            Debug.Print("") 'Breakpoint Landing
        Else
            Stop
        End If
    End Sub

    Private Sub LbxMatlQty_DblClick(
    ByVal Cancel As Microsoft.Vbe.Interop.Forms.ReturnBoolean
)
        txbMatlQty.Text = LbxMatlQty.Text
    End Sub

    Private Sub LbxMatlQty_MouseMove(
    ByVal Button As Integer, ByVal Shift As Integer,
    ByVal X As Single, ByVal Y As Single
)
        If Button = 1 Then
            LbxMatlQty.DoDragDrop(LbxMatlQty.Text, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub TxbMatlQty_Change()
        Dim ck As Double
        Dim tx As String
        Dim gp As Object
        Dim mx As Long
        Dim dx As Long

        With txbMatlQty
            tx = .Text
            gp = Split(tx, ".")
            mx = UBound(gp)

            For dx = LBound(gp) To mx
                ck = Val(gp(dx))

                If ck > 0 Then
                    gp(dx) = CStr(ck)
                ElseIf dx > 0 Then
                    gp(dx) = ""
                Else
                    gp(dx) = "0"
                End If
            Next
            tx = Join(gp, ".")

            If tx <> .Text Then
                Forms.Application.DoEvents()
                .Text = tx
            End If
        End With
    End Sub
End Class