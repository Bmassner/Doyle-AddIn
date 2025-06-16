Imports System.Windows

Module FmDvl00

    'Public Function fd0g1f1(ls As Object) As Object
    '    Dim Fm As fmEmpty
    '    Dim ListBox As Forms.TextBox
    '    Dim mg As Single
    '    Dim wk As Object
    '    Dim dc As Scripting.Dictionary
    '    Dim dx As Long
    '    Dim rt As String

    '    If TypeOf ls Is Array Then
    '        wk = ls
    '    ElseIf TypeOf ls Is Object Then
    '        dc = dcOb(obOf(ls))
    '        If dc Is Nothing Then
    '            wk = vbEmpty
    '        Else
    '            wk = dcOb(ls).Keys
    '        End If
    '    Else
    '        wk = vbEmpty
    '    End If

    '    If wk Is Nothing Then wk = {
    '    "*vvvvvvvvvvv*",
    '    "*Unsupported*",
    '    "*List Source*",
    '    "*^^^^^^^^^^^*"
    '}

    '    Fm = nuFmEmpty()
    '    mg = 10

    '    ListBox = nuMsFmCtListBox(Fm, , "LbxA")
    '    With obMsFmControl(ListBox)
    '        .Top = mg
    '        .Left = mg
    '        .Height = Fm.InsideHeight - mg - mg
    '        .Width = Fm.InsideWidth - mg - mg
    '    End With

    '    With ListBox
    '        .MultiSelect = FmMultiSelectMulti ' FmMultiSelectExtended
    '        .TextStyle = FmListStyleOption

    '        .Text = wk
    '        Fm.Show(vbModal)

    '        'Stop
    '        rt = ""
    '        dx = 0
    '        Do Until dx = .TextCount
    '            If .Selected(dx) Then
    '                rt = rt & vbVerticalTab & .Text(dx)
    '            End If
    '            dx = 1 + dx
    '        Loop
    '    End With

    '    fd0g1f1 = Split(Mid$(rt, 2), vbVerticalTab)
    'End Function

    Public Function nuFmEmpty(Optional f As Object) As fmEmpty
        With New fmEmpty
            '
            nuFmEmpty = .Itself
        End With
    End Function

    Public Function obMsFmControl(it As Object) As MSForms.Control
        Dim ob As Object

        ob = obOf(it)
        If TypeOf ob Is MSForms.Control Then
            obMsFmControl = ob
        Else
            obMsFmControl = Nothing
        End If
    End Function

    Public Function nuMsFmCtListBox(fm As fmEmpty,
    Optional sp As Object = vbEmpty,
    Optional nm As Object = "",
    Optional vs As Boolean = True
) As MSForms.TextBox 'MSForms.UserForm
        '
        ' nuMsFmCtListBox -- add new ListBox
        '     Control to supplied FmEmpty Object
        '
        '     accepts, but does not yet use,
        '     a specification sp laying out
        '     the parameters defining the
        '     desired control
        '
        '     note that Fm MUST be FmEmpty
        '     a general MSForms UserForm is
        '     NOT accepted, because it does
        '     NOT support certain essential
        '     properties, for example, there
        '     are not properties to  size
        '     or position, which are essential
        '     to the goals of this system
        '
        Dim rt As MSForms.TextBox
        'Dim ct As MSForms.Control

        'fm.Left

        rt = fm.Controls.Add(
        "MSForms.TextBox.1", , vs
    )
        If Len(nm) > 0 Then
            obMsFmControl(rt).Name = nm
            ' ct = rt
            'ct.Name = nm
        End If

        With rt
        End With

        nuMsFmCtListBox = rt
    End Function
End Module