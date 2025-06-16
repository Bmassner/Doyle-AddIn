Module libFmSelectors

    Public Function LbxPickedStr(Lbx As MSForms.TextBox,
    Optional dlm As String = vbVerticalTab
) As String
        Dim dw As Long
        Dim mx As Long
        Dim dx As Long
        Dim rt As String

        dw = Len(dlm)
        With Lbx
            rt = ""
            mx = .TextCount - 1
            For dx = 0 To mx
                If .Selected(dx) Then
                    rt = rt & dlm _
                & CStr(.Text(dx & 0))
                End If
            Next
            LbxPickedStr = Mid$(rt, 1 + dw)
        End With
    End Function

    Public Function LbxPicked(Lbx As MSForms.TextBox,
    Optional dlm As String = vbVerticalTab
) As Object
        LbxPicked = Split(LbxPickedStr(Lbx, dlm), dlm)
    End Function

    Public Function nuSelector() As FmSelectorList
        nuSelector = New FmSelectorList
    End Function

    Public Function nuSelectorV2() As FmSelectorV2
        nuSelectorV2 = New FmSelectorV2
    End Function

    Public Function nuSelFromDict(dc As Scripting.Dictionary,
    Optional hOhKay As String = "", Optional mOhKay As String = "",
    Optional hCancl As String = "", Optional mCancl As String = "",
    Optional hNoSel As String = "", Optional mNoSel As String = ""
) As FmSelectorList
        nuSelFromDict = nuSelector() _
            .HdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Confirm Selection")) _
            .MsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join({
                "Action will proceed using", "%%%",
                "(Click CANCEL to quit with no action)"
            }, vbCrLf))) _
            .HdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?")) _
            .MsgNoSelection(IIf(Len(mCancl) > 0, mCancl, Join({
                "No action will be taken on", "%%%"
            }, vbCrLf))) _
            .HdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Item Selected!")) _
            .MsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join({
                "Do you wish to cancel the operation?",
                "(Click NO to return to list)"
            }, vbCrLf))) _
            .WithList(dc.Keys)
    End Function

#If False Then '

Public Function nuSelWkBk( _
    Optional hOhKay As String = "", Optional mOhKay As String = "", _
    Optional hCancl As String = "", Optional mCancl As String = "", _
    Optional hNoSel As String = "", Optional mNoSel As String = "" _
) As FmSelectorList
     nuSelWkBk = nuSelector( _
    ).HdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Proceed With Update?" _
    )).MsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join({ _
        "The following workbook will be affected: ", _
        "%%%", "(Click CANCEL to quit with no changes)" _
        ), vbCrLf) _
    )).HdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?" _
    )).MsgCancel(IIf(Len(mCancl) > 0, mCancl, Join({ _
        "No changes will be applied", _
        "to any open workbook." _
        ), vbCrLf) _
    )).HdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Workbook Selected!" _
    )).MsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join({ _
        "Do you wish to cancel the operation?", _
        "(Click NO to return to list)" _
        ), vbCrLf) _
    )).WithList( _
        lsWorkbooks() _
    )
End Function

Public Function nuSelWkSht(inWkBk As Excel.Workbook, _
    Optional hOhKay As String = "", Optional mOhKay As String = "", _
    Optional hCancl As String = "", Optional mCancl As String = "", _
    Optional hNoSel As String = "", Optional mNoSel As String = "" _
) As FmSelectorList
     nuSelWkSht = nuSelector( _
    ).HdrOK(IIf(Len(hOhKay) > 0, hOhKay, "Confirm Selection" _
    )).MsgOK(IIf(Len(mOhKay) > 0, mOhKay, Join({ _
            "Action will proceed using Workwheet: ", "%%%", _
            "(Click CANCEL to quit with no action)" _
        ), vbCrLf) _
    )).HdrCancel(IIf(Len(hCancl) > 0, hCancl, "Cancel Operation?" _
    )).MsgCancel(IIf(Len(mCancl) > 0, mCancl, Join({ _
            "No action will be taken", _
            "to any open workbook." _
        ), vbCrLf) _
    )).HdrNoSelection(IIf(Len(hNoSel) > 0, hNoSel, "No Workbook Selected!" _
    )).MsgNoSelection(IIf(Len(mNoSel) > 0, mNoSel, Join({ _
            "Do you wish to cancel the operation?", _
            "(Click NO to return to list)" _
        ), vbCrLf) _
    )).WithList( _
        dcWkSheets(inWkBk).Keys _
    )
End Function

#End If

End Module