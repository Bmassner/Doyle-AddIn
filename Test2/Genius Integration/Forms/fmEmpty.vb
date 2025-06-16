Public Class fmEmpty
    Private Watchers As Scripting.Dictionary

    Event CloseRequested(CloseMode As Integer)

    Public Function Itself() As fmEmpty
        Itself = Me
    End Function

    Public Function Notify(ob As Object, Optional Ky As Object = vbEmpty) As Object
        Dim dx As Long

        With Watchers
            If IsNothing(Ky) Then
                dx = .Count
                Do While .Exists(dx)
                    dx = 1 + dx
                Loop
                Notify = Notify(ob, dx)
            Else
                If .Exists(Ky) Then
                    Notify = vbEmpty
                Else
                    .Add(Ky, ob)
                    Notify = Ky
                End If
            End If : End With
    End Function

    Public Function NoMsgs(nm As Object) As Object
        With Watchers
            If .Exists(nm) Then
                .Remove(nm)
                NoMsgs = nm
            Else
                NoMsgs = vbEmpty
            End If : End With
    End Function

    Private Sub UserForm_Initialize()
        Watchers = New Scripting.Dictionary
    End Sub

    Private Sub UserForm_QueryClose(
    Cancel As Integer, CloseMode As Integer
)
        Dim ck As VbMsgBoxResult

        Cancel = 1
        If Watchers.Count > 0 Then
            RaiseEvent CloseRequested(CloseMode)
        Else
            ck = MsgBox(Join({
            "Review any selections",
            "and select Yes if ready.",
            "Otherwise, select No.", vbNewLine, vbYesNo, "Close Form?"}))
            If ck = vbYes Then
                Me.Hide()
            Else
            End If
        End If
    End Sub

    Private Sub UserForm_Terminate()
        Watchers.RemoveAll()
        Watchers = Nothing
    End Sub
End Class