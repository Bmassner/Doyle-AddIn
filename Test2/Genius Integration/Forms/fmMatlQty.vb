Class FmMatlQty
    Event Sent(Signal As VbMsgBoxResult)

    Private Sub cmdCancel_Click()
        RaiseEvent Sent(vbCancel)
    End Sub

    Private Sub cmdOk_Click()
        RaiseEvent Sent(vbOK)
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer,
    CloseMode As Integer
)
        Cancel = 1
        cmdCancel_Click()
    End Sub
End Class