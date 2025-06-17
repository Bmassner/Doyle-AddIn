Class fmSelectorV2
    Private Const tagSelected As String = "%%%"
    Private dc As Scripting.Dictionary

    Private msCancelHead As String
    Private msCancelMain As String
    Private msNoSelHead As String
    Private msNoSelMain As String
    Private msOkHead As String
    Private msOkMain As String
    'Private msCancelMain As String
    'Private msCancelMain As String

    Public Function MsgCancel(Message As String) As fmSelectorV2
        msCancelMain = Message
        MsgCancel = Me
    End Function

    Public Function MsgNoSelection(Message As String) As fmSelectorV2
        msNoSelMain = Message
        MsgNoSelection = Me
    End Function

    Public Function MsgOK(Message As String) As fmSelectorV2
        msOkMain = Message
        MsgOK = Me
    End Function

    Public Function HdrCancel(Message As String) As fmSelectorV2
        msCancelHead = Message
        HdrCancel = Me
    End Function

    Public Function HdrNoSelection(Message As String) As fmSelectorV2
        msNoSelHead = Message
        HdrNoSelection = Me
    End Function

    Public Function HdrOK(Message As String) As fmSelectorV2
        msOkHead = Message
        HdrOK = Me
    End Function

    Public Function SelectIfIn(Message As String) As fmSelectorV2
        Dim dx As Long

        With Me.lsbSelection
            dx = .TabIndex
            On Error Resume Next
            Err.Clear()
            .Name = Message
            If Err.Number Then
                .TabIndex = dx
                '.Value = ""
                Err.Clear()
            End If
            On Error GoTo 0
        End With

        SelectIfIn = Me
    End Function

    Public Function WithList(Message As Object) As fmSelectorV2
        Dim ky As Object
        Dim it As String

        If TypeOf Message Is Object Then
            'Me.lsbSelection.List = Message
            If TypeOf Message Is Scripting.Dictionary Then
                dc = Message
            ElseIf TypeOf Message Is Inventor.NameValueMap Then
                dc = dcFromAiNameValMap(obOf(Message))
            Else
                'Stop
                Debug.Print("") 'Breakpoint Landing
                dc = Nothing
            End If

            If dc Is Nothing Then
                Me.lsbSelection.Text = "<no items>"
            ElseIf dc.Count > 0 Then
                Me.lsbSelection.Text = dc.Keys
            Else
                Me.lsbSelection.Text = "<no items>"
            End If
        ElseIf IsArray(Message) Then
            dc = New Scripting.Dictionary

            For Each ky In Message
                dc.Add(CStr(ky), ky)
            Next

            Me.lsbSelection.Text = dc.Keys
        Else
            'Stop
            Debug.Print("") 'Breakpoint Landing
        End If

        WithList = Me
    End Function

    Private Sub btnCancel_Click()
        ''
        If MsgBox(
        msCancelMain, vbYesNo, msCancelHead
    ) = vbYes Then
            Me.lsbSelection.TabIndex = -1
            Me.Hide
        Else
            'Do nothing
        End If
    End Sub

    Private Sub btnOk_Click()
        ''
        Dim ck As VbMsgBoxResult
        Dim mx As Long
        Dim dx As Long
        Dim ct As Long

        Dim ls As String

        With Me.lsbSelection
            If .SelectedIndex = 1 Then
                If .Text < 0 Then
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
                        .Text = -1
                        Me.Hide()
                    Else
                        'Do nothing
                    End If
                End If
            Else
                ls = LbxPickedStr(Me.lsbSelection, vbCrLf)

                'ct = 0
                'mx = .ListCount - 1
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
                        Me.Hide
                    ElseIf ck = vbCancel Then
                        .TabIndex = -1
                        Me.Hide
                    Else
                        'Do nothing
                    End If
                Else
                    ck = MsgBox(msNoSelMain, vbYesNo, msNoSelHead)
                    If ck = vbYes Then Me.Hide
                End If
            End If
        End With
    End Sub

    Private Sub lsbSelection_Change()
        '
        '
        Dim ck As String

        On Error Resume Next
        Err.Clear()
        ck = Me.lsbSelection.Text
        If Err.Number <> 0 Then ck = ""
        On Error GoTo 0

        With dc
            If .Exists(ck) Then
                On Error Resume Next

                Err.Clear()
                tbxView.Text = CStr(.Item(ck))
                If Err.Number <> 0 Then
                    tbxView.Text = "<not printable>"
                End If

                On Error GoTo 0
            Else
                tbxView.Text = "<no data>"
            End If
        End With
    End Sub

    Private Sub lsbSelection_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        btnOk_Click()
    End Sub

    Private Sub UserForm_Initialize()
        '
        msCancelHead = "Cancel Operation?"
        msNoSelHead = "No Selection!"
        msOkHead = "Proceed?"
        '
        msCancelMain = "Selection will be canceled."
        msNoSelMain = Join({
        "Do you wish to cancel?",
        "(Click NO to return to list)" _
    , vbCrLf})
        msOkMain = Join({
        "Current selection is: ", tagSelected,
        "(Click CANCEL to quit with no selection)" _
    , vbCrLf})
        '
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        '
        Cancel = 1
        btnCancel_Click()
    End Sub

    Public Function GetReply(
    Optional List As Object = Nothing,
    Optional DefaultValue As String = "%$#@"
) As String
        Dim rt As String

        rt = ""
        With Me.WithList(List).SelectIfIn(DefaultValue)
            '.lsbSelection.List = lsWorkbooks()
            .Show()
            If .lsbSelection.SelectedItems.Count = 1 Then
                rt = .lsbSelection.Text
            Else
                rt = LbxPickedStr(.lsbSelection, vbCrLf)
            End If
        End With
        GetReply = rt
    End Function

    Private Sub UserForm_Resize()
        '
    End Sub
End Class