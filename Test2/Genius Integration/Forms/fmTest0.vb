Class fmTest0
    Public Function ft0g0f0(im As stdole.StdPicture) As Long
        With Me
            .imTNail.Image = im
            .Show()
        End With
        ft0g0f0 = 0
    End Function

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        Cancel = 1
        Me.Hide
    End Sub
End Class