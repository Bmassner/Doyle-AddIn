Class FmSelectorList
    Private Const tagSelected As String = "%%%"

    Private msCancelHead As String
    Private msCancelMain As String
    Private msNoSelHead As String
    Private msNoSelMain As String
    Private msOkHead As String
    Private msOkMain As String
    'Private msCancelMain As String
    'Private msCancelMain As String

    Public Function SetCancelMessage(message As String) As FmSelectorList
        Me.msCancelMain = message
        Return Me
    End Function

    Public Function MsgNoSelection(message As String) As FmSelectorList
        msNoSelMain = message
        MsgNoSelection = Me
    End Function

    Public Function MsgOK(message As String) As FmSelectorList
        msOkMain = message
        MsgOK = Me
    End Function

    Public Function HdrCancel(message As String) As FmSelectorList
        msCancelHead = message
        HdrCancel = Me
    End Function

    Public Function HdrNoSelection(message As String) As FmSelectorList
        msNoSelHead = message
        HdrNoSelection = Me
    End Function

    Public Function HdrOK(message As String) As FmSelectorList
        msOkHead = message
        HdrOK = Me
    End Function

    Public Function SelectIfIn(message As String) As FmSelectorList
        Dim dx As Long

        With Me.lsbSelection
            dx = .TabIndex
            On Error Resume Next
            Err.Clear()
            .Text = message
            If Err.Number Then
                .TabIndex = dx
                '.Text = ""
                Err.Clear()
            End If
            On Error GoTo 0
        End With

        SelectIfIn = Me
    End Function

    Public Function WithList(message As Object) As FmSelectorList
        If TypeOf message Is Object Then
            Me.lsbSelection.Text = message
        Else
            'Stop
            Debug.Print("") 'Breakpoint Landing
        End If
        WithList = Me
    End Function

    Private Sub BtnCancel_Click()
        ''
        If MsgBox(
        msCancelMain, vbYesNo, msCancelHead
    ) = vbYes Then
            Me.lsbSelection.TabIndex = -1
            Me.Hide()
        Else
            'Do nothing
        End If
    End Sub

    Private Sub BtnOk_Click()
        ''
        Dim ck As VbMsgBoxResult
        Dim mx As Long
        Dim dx As Long
        Dim ct As Long

        Dim ls As String

        With Me.lsbSelection
            If .CanSelect = .Multiline Then
                If .TabIndex < 0 Then
                    ck = MsgBox(msNoSelMain, vbYesNo, msNoSelHead)
                    If ck = vbYes Then Me.Hide()
                Else
                    ck = MsgBox(
                    Join(Split(msOkMain, tagSelected), .Text),
                    vbYesNoCancel, msOkHead
                )
                    If ck = vbYes Then
                        Me.Hide()
                    ElseIf ck = vbCancel Then
                        .TabIndex = -1
                        Me.Hide()
                    Else
                        'Do nothing
                    End If
                End If
            Else
                ls = LbxPickedStr(Me.lsbSelection, vbCrLf)

                'ct = 0
                'mx = .TextCount - 1
                'For dx = 0 To mx
                'If .Selected(dx) Then ct = 1 + ct
                'Next

                If Len(ls) > 0 Then
                    ck = MsgBox(Join(
                    Split(msOkMain, tagSelected),
                    vbCrLf & ls & vbCrLf
                    ), vbYesNoCancel, msOkHead
                )
                    If ck = vbYes Then
                        Me.Hide()
                    ElseIf ck = vbCancel Then
                        .TabIndex = -1
                        Me.Hide()
                    Else
                        'Do nothing
                    End If
                Else
                    ck = MsgBox(msNoSelMain, vbYesNo, msNoSelHead)
                    If ck = vbYes Then Me.Hide()
                End If
            End If
        End With
    End Sub

    Private Sub LsbSelection_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        BtnOk_Click()
    End Sub

    Private Sub UserForm_Initialize()
        '
        msCancelHead = "Cancel Operation?"
        msNoSelHead = "No Selection!"
        msOkHead = "Proceed?"
        '
        msCancelMain = "Selection will be canceled."
        msNoSelMain = Join({(
            "Do you wish to cancel?",
            "(Click NO to return to list)"
        ), vbCrLf})
        msOkMain = Join({(
            "Current selection is: " & tagSelected,
            "(Click CANCEL to quit with no selection)"
        ), vbCrLf})
        '
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        '
        Cancel = 1
        BtnCancel_Click()
    End Sub

    Public Function GetReply(
        Optional List As Object = Nothing,
        Optional DefaultValue As String = "%$#@"
    ) As String
        Dim rt As String

        rt = ""
        With Me.WithList(List).SelectIfIn(DefaultValue)
            '.lsbSelection.Text = lsWorkbooks()
            .Show()
            ' Ensure lsbSelection is a ListBox and check for MultiSelect property
            If TypeName(.lsbSelection) = "ListBox" Then
                If .lsbSelection.SelectedText = 0 Then ' 0 = fmMultiSelectSingle
                    rt = .lsbSelection.Text
                Else
                    rt = LbxPickedStr(.lsbSelection, vbCrLf)
                End If
            Else
                rt = ""
            End If
        End With
        GetReply = rt
    End Function

End Class