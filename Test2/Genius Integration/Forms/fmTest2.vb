Class FmTest2
    Private ad As Inventor.Document

    Private PropertySetDsn As Inventor.Property
    Private PropertySetUsr As Inventor.Property

    Private dcDsn As Scripting.Dictionary
    Private dcUsr As Scripting.Dictionary

    Private FamilyProp As Inventor.Property
    Private StockProp As Inventor.Property

    'Private dmFmHt As Long
    'Private dmFmWd As Long
    ''Private dmLbMsHt As Long
    Private dmLbMsWd As Long
    ''Private dmDfFmMsHt As Long
    ''Private dmDfFmMsWd As Long
    Private dmFmHt2cmdTop As Long

    Private rtAnswer As VbMsgBoxResult

    Public Function AskAbout(
    Optional AiDoc As Inventor.Document = Nothing,
    Optional txPre As String = "",
    Optional txPost As String = ""
) As VbMsgBoxResult
        '
        ' AskAbout -- prompt User for action
        '     to take on supplied Document
        ' UPDATE[2021.12.13]
        '     Document parameter now Optional.
        '     will attempt to use previously
        '     registered Document when none
        '     supplied. Warning/error message
        '     will be presented if no Document
        '     is registered OR supplied.
        '
        Dim pc As stdole.IPictureDisp
        Dim pn As String
        Dim StockNumb As String
        Dim PartDesc As String
        Dim dj As Single 'use to adjust
        '   form height and positions
        '   of command buttons

        rtAnswer = vbCancel
        If Not AiDoc Is Nothing Then
            ad = AiDoc
        End If

        If ad Is Nothing Then
            MsgBox("Review or Update requested" _
            & vbCrLf & "but no Document provided!" _
            & vbCrLf & "" _
            & vbCrLf & "" _
        , vbOKOnly, "No Document!")
            rtAnswer = vbNo
        ElseIf aiDocPartFromCCtr(ad) Is Nothing Then 'AiDoc
            ' ad = AiDoc
            With ad
                pc = .Thumbnail
                PropertySetDsn = .Propertys(gnDesign)
                PropertySetUsr = .Propertys(gnCustom)

                dcDsn = dcAiPropsIn(PropertySetDsn)
                dcUsr = dcAiPropsIn(PropertySetUsr)

                FamilyProp = PropertySetDsn.Item(pnFamily)
                With dcUsr
                    If .Exists(pnRawMaterial) Then
                        StockProp = PropertySetUsr.Item(pnRawMaterial)
                    Else
                        On Error Resume Next
                        Err.Clear()
                        StockProp = PropertySetUsr.Add("", pnRawMaterial)
                        If Err.Number Then
                            Debug.Print(Err.Number, Err.Description)
                            Stop
                        Else
                            .Add(pnRawMaterial, StockProp)
                        End If
                        On Error GoTo 0
                    End If
                End With

                If Not StockProp Is Nothing Then StockNumb = StockProp.Text
                pn = PropertySetDsn.Item(pnPartNum).Text
                PartDesc = PropertySetDsn.Item(pnDesc).Text
            End With

            With Me
                .Text = "Please Review Item: " & pn

                If pc Is Nothing Then
                Else
                    .imThmNail.Image = pc
                End If

                dj = FmHtAdjust(lblHtAdjust(.lbMsg,
                IIf(Len(txPre) > 0,
                    txPre & vbCrLf & vbCrLf, ""
                ) & Join({pn & ": " & PartDesc,
                    pnCatWebLink & ": " & PropertySetDsn.Item(pnCatWebLink).Text,
                    pnMaterial & ": " & PropertySetDsn.Item(pnMaterial).Text
                }, vbCrLf & vbCrLf) _
                & IIf(Len(txPost) > 0,
                    vbCrLf & vbCrLf & txPost, ""
                )
            ))
                '.dbFamily.Text = FamilyProp.Text

                .Show()
            End With
        Else
            MsgBox(ad.DisplayName _
            & vbCrLf & "is a Content Center part" _
            & vbCrLf & "and cannot be updated." _
            & vbCrLf & "" _
            & vbCrLf & "" _
        , vbOKOnly, "Can't Update!") 'AiDoc
            rtAnswer = vbYes
        End If

        AskAbout = rtAnswer 'vbYes ' = 1
    End Function

    Public Function UseDict(
    AiDoc As Inventor.Document
) As FmTest2
        '
        ' NEWMETHOD[2021.12.13]
        ' Using -- assign supplied Document
        '     for use in all subsequent calls
        '     to AskAbout without one.
        '
        rtAnswer = vbCancel

        If Not AiDoc Is Nothing Then
            ad = AiDoc
        End If

        Return Me
    End Function

    Public Function Document(
    Optional AiDoc As Inventor.Document = Nothing
) As Inventor.Document
        '
        ' NEWMETHOD[2021.12.13]
        ' Document -- return currently active Document
        '
        If AiDoc Is Nothing Then
            Document = ad
        Else
            Document = Me.UseDict(AiDoc).Document
        End If
    End Function

    Private Function FmHtAdjust(by As Long) As Single
        Dim cmdTop As Long

        With Me
            .Height = .Height + by

            .cmdLt.Top = .Height - dmFmHt2cmdTop
            .cmdCt.Top = .cmdLt.Top
            .cmdRt.Top = .cmdLt.Top

            FmHtAdjust = .Height
        End With
    End Function

    Private Function lblHtAdjust(
    Label As MSForms.Label, tx As String
) As Single
        Dim ct As MSForms.Control
        Dim au As Boolean
        Dim wd As Single
        Dim ht As Single

        ct = Label
        With ct
            wd = .Width
            ht = .Height

            With Label
                au = .AutoSize
                .Caption = tx
                .AutoSize = True
                ct.Width = dmLbMsWd
                .AutoSize = au
            End With

            lblHtAdjust = Int(.Height - ht)
        End With
    End Function

    Private Sub cmdCt_Click()
        rtAnswer = vbNo
        Me.Hide()
    End Sub

    Private Sub cmdLt_Click()
        rtAnswer = vbYes
        Me.Hide()
    End Sub

    Private Sub cmdRt_Click()
        rtAnswer = vbCancel
        Me.Hide()
    End Sub

    Private Sub UserForm_Initialize()
        '
        With Me
            'dmFmHt = .Height
            'dmFmWd = .Width
            With .lbMsg
                'dmLbMsHt = .Height
                dmLbMsWd = .Width
            End With
            dmFmHt2cmdTop = .Height - .cmdLt.Top
        End With
        'dmDfFmMsWd = dmFmWd - dmLbMsWd
        'dmDfFmMsHt = dmFmHt - dmLbMsHt
        rtAnswer = vbCancel
    End Sub

    Private Sub UserForm_Click()
        '
    End Sub

    Private Sub UserForm_Layout()
        'Stop
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        Cancel = 1
        Me.Hide()
    End Sub

    Private Sub UserForm_Terminate()
        'cn.Close
        ' cn = Nothing
    End Sub

End Class