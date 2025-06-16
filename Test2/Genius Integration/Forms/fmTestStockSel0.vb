'VERSION 5.00
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmTestStockSel0 
'   Caption         =   "UserForm1"
'   ClientHeight    =   5490
'   ClientLeft      =   120
'   ClientTop       =   450
'   ClientWidth     =   4560
'   OleObjectBlob   =   "fmTestStockSel0.frx":0000
'   StartUpPosition =   1  'CenterOwner
'End
'Attribute VB_Name = "fmTestStockSel0"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False

'Option Explicit
Class FmTestStockSel0
    Private cn As ADODB.Connection
    Private rsFam As ADODB.Record
    Private rsItm As ADODB.Record

    Private Sub lbxFamily_Change()
        With Me
            rsItm.Filter = "Family = '" & .lbxFamily.Value & "'"
            .lbxItem.List = m0g3f1(rsItm)
        End With
    End Sub

    Private Sub UserForm_Initialize()
        cn = CnGnsDoyle()
        With cn
            rsFam = .Execute(Join(Array(
            "select Family, Description1",
            "from vgMfiFamilies",
            "where FamilyGroup = 'RAW'"
        ), " "))
            rsItm = .Execute(Join(Array(
            "Select I.Family, I.Item, I.Description1",
            "From vgMfiItems as I",
            "Inner Join vgMfiFamilies as F",
            "On I.Family = F.Family",
            "Where F.FamilyGroup = 'RAW'"
        ), " "))
        End With

        Me.lbxFamily.List = m0g3f1(rsFam)
    End Sub
End Class