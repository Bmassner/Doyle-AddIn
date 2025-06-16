Public Class fmGetList
    'VERSION 5.00
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmGetList 
    '   Caption         =   "List Entry"
    '   ClientHeight    =   6975
    '   ClientLeft      =   120
    '   ClientTop       =   465
    '   ClientWidth     =   3525
    '   OleObjectBlob   =   "fmGetList.frx":0000
    '   StartUpPosition =   1  'CenterOwner
    'End
    'Attribute VB_Name = "fmGetList"
    'Attribute VB_GlobalNameSpace = False
    'Attribute VB_Creatable = False
    'Attribute VB_PredeclaredId = True
    'Attribute VB_Exposed = False

    'Option Explicit

    'Event CheckOut(Cancel As Long)

    Private bg As String
    Private rt As String

    Public Function AskUser(
        Optional usingText As String = ""
    ) As String
        bg = usingText          'save initial text
        txIn.Text = bg          'initialize text box
        Me.ShowDialog()         'and wait...
        AskUser = rt            'return final result
    End Function

    Private Sub CheckOut(NoChg As Long)
        Dim ck As System.Windows.Forms.DialogResult

        If NoChg = 0 Then
            ck = System.Windows.Forms.MessageBox.Show(
                "Use this List?",
                "Confirm",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question
            )
            If ck = DialogResult.Yes Then rt = txIn.Text
        Else
            ck = System.Windows.Forms.MessageBox.Show(
                "Cancel this Entry?",
                "Cancel",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question
            )
            If ck = DialogResult.Yes Then rt = bg
        End If

        If ck = DialogResult.Yes Then Me.Hide()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        CheckOut(1) 'no change
    End Sub

    Private Sub cmdOk_Click(sender As Object, e As EventArgs) Handles cmdOk.Click
        CheckOut(0) 'commit changes
    End Sub

    Private Sub fmGetList_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        CheckOut(1) 'no change
    End Sub
End Class