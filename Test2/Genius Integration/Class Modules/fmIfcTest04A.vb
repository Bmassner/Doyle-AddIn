Imports System.Windows

Class FmIfcTest04A

    Private Fm As FmTest04
    'Private fmSpec As fmTest04
    'Private WithEvents Fm As MSForms.UserForm 'fmTest04
    Private WithEvents LbxSpecOps As MSForms.TextBox
    Private WithEvents LbxSpecSet As MSForms.TextBox
    Private WithEvents LbxSpec As MSForms.TextBox

    Private dcSpecPairs As Scripting.Dictionary
    Private dcInitList As Scripting.Dictionary
    Private dcActvOps As Scripting.Dictionary
    Private dcActv As Scripting.Dictionary

    Private Const kyInitList As String = "initSpecs"    'key name identifying initial spec list
    Private Const vsnString As String = "Form Test04 Interface A v0.1.0.0 [2022.03.17]"
    ' prior values                    "Form Test04 Interface A v0.0.0.0 [2022.03.03]"
    '                                 ""
    '                                 ""
    '
    '
    '

    Private Sub Class_Initialize()
        Fm = New FmTest04
        LbxSpecOps = Fm.LbxSpecOps
        LbxSpec = Fm.LbxSpecSel

        dcSpecPairs = dcGnsMatlSpecPairings()
        ' probably want to make this controllable
        ' from client processes to support more
        ' general usage. this will do for now.

        dcInitList = dcSpecPairs
        ' like dcSpecPairs, probably want this
        ' to be controllable from client to
        ' facilitate flexible usage, but again,
        ' this will serve for the moment.

        dcActvOps = dcInitList
        dcActv = New Scripting.Dictionary
        ' since the user will not have been met
        ' at this point, no options would be
        ' selected at this moment, but ALL options
        ' in dcInitList should be presently active.
    End Sub

    Private Sub Class_Terminate()
        dcActv = Nothing
        dcActvOps = Nothing
        dcInitList = Nothing
        dcSpecPairs = Nothing

        LbxSpec = Nothing
        LbxSpecOps = Nothing
        Fm = Nothing
    End Sub

    Public Function Itself() As FmIfcTest04A
        ' returns this FmIfcTest04A class instance "Itself"
        ' should be HIGHLY useful inside a With context
        Itself = Me
    End Function

    Public Function UseDict(
    Optional About As Scripting.Dictionary = Nothing
) As FmIfcTest04A
        If About Is Nothing Then
            Return UseDict(NuDcPopulator(
            ).Setting(kyInitList, dcInitList
        ).Dictionary)

        ElseIf About.Exists(kyInitList) Then
            dcInitList = dcOb(About.Item(kyInitList))

            dcActvOps = dcInitList
            dcActv = New Scripting.Dictionary
            ' noting these steps are also taken
            ' in Class_Initialize, it's tempting
            ' to wonder why they should appear in
            ' both places, however, the purpose
            ' THERE is to ensure a valid up
            ' at the earliest possible moment.
            '
            ' it would likely be appropriate
            ' to consolidate the two into a single
            ' procedure to be called from either place.

            LbxSpecOps.Text = dcActvOps.Keys
            LbxSpec.Text = dcActv.Keys
        End If

        Return Me

    End Function

    Public Function SeeUser(
    Optional About As Scripting.Dictionary = Nothing
) As FmIfcTest04A
        ' REV[2022.03.17.1308]
        '     disabling If-Then-Else blocking,
        '     including content of If block.
        '     '
        '     Only active sections of Else
        '     block to remain active.
        '     '
        '     Since majority of process formerly
        '     performed here is now addressed by
        '     new Function Using, it should now
        '     be sufficient to call that Function
        '     for preparation, and then present
        '     the UserForm for user's response.
        '     '
        'If About Is Nothing Then
        '     SeeUser = SeeUser(nuDcPopulator( _
        '        ).Setting(kyInitList, dcInitList _
        '    ).Dictionary)
        'Else
        ' REV[2022.03.17.1258] -- IMPORTANT!
        '     the following section, copied
        '     to new Method Function Using
        '     (see above) has been disabled
        '     here pending removal
        ' dcInitList = dcOb(About.Item(kyInitList))
        '
        ' dcActvOps = dcInitList
        ' dcActv = New Scripting.Dictionary
        ' noting these steps are also taken
        ' in Class_Initialize, it's tempting
        ' to wonder why they should appear in
        ' both places, however, the purpose
        ' THERE is to ensure a valid up
        ' at the earliest possible moment.
        '
        ' it would likely be appropriate
        ' to consolidate the two into a single
        ' procedure to be called from either place.
        '
        'LbxSpecOps.Text = dcActvOps.Keys
        'LbxSpec.Text = dcActv.Keys
        '
        ' REV[2022.03.17.1258] ENDS HERE

        ' REV[2022.03.17.1301]
        '     implementation of Method Function
        '     Using, having taken over the steps
        '     disabled immediately above, is now
        '     called in their stead. Separation
        '     of that sequence into its own
        '     Function enables the preparation
        '     of this Class instance without
        '     immediately invoking the UserForm.
        With Me.UseDict(About)
            Fm.Visible = True

            '

            SeeUser = .Itself
        End With
        'End If
    End Function

    Public Function Version() As String
        Version = vsnString
    End Function

    Private Function ClsAddSpec(sp As String) As Long
        Dim rt As Long

        rt = 0
        If dcActvOps.Exists(sp) Then
            If dcActv.Exists(sp) Then
                rt = 1 'spec already 
                Stop '
            Else
                dcActv.Add(sp, sp)
            End If
            LbxSpec.Text = dcActv.Keys

            dcActvOps = dcSpecSubWith(sp, dcActvOps)
            LbxSpecOps.Text = dcActvOps.Keys
            Debug.Print("")  'Breakpoint Landing
        Else
            rt = 2
            Stop '
        End If

        ClsAddSpec = rt
    End Function

    Private Function ClsDropSpec(sp As String) As Long
        Dim rt As Long
        Dim ky As Object

        rt = 0

        dcActv.Remove(sp)

        ' first attempt to reinitialize
        ' dcActvOps from dcInitList
        dcActvOps = dcSpecSubWithAll(
        dcActv, dcInitList
    )

        ' that SHOULD have left sp back
        ' in dcActvOps. if not there,
        ' try the FULL  dcSpecPairs
        If Not dcActvOps.Exists(sp) Then
            dcActvOps = dcSpecSubWithAll(
        dcActv, dcSpecPairs
    )
        End If

        ' check once more to maks sure it's in
        ' if not, we've got a REAL problem.
        If Not dcActvOps.Exists(sp) Then
            rt = 1
            Stop
        End If

        LbxSpec.Text = dcActv.Keys
        LbxSpecOps.Text = dcActvOps.Keys
        ' this might be more flexibly implemented
        ' in a separate function or procedure

        ClsDropSpec = rt
    End Function

    Private Sub LbxSpecOps_DblClick(
    ByVal Cancel As Microsoft.Vbe.Interop.Forms.ReturnBoolean
)
        'Dim sp As String
        Dim ck As Long

        ck = ClsAddSpec(LbxSpecOps.Text)
        If ck Then 'we got an error
            Stop
        End If

        'sp = LbxSpecOps.Text
        'If dcActvOps.Exists(sp) Then
        '    If dcActv.Exists(sp) Then
        '        Stop '
        '    Else
        '        dcActv.Add(sp, sp
        '    End If
        '    LbxSpec.Text = dcActv.Keys
        '
        '     dcActvOps = dcSpecSubWith(sp, dcActvOps)
        '    LbxSpecOps.Text = dcActvOps.Keys
        '    Debug.Print("") 'Breakpoint Landing
        'Else
        '    Stop '
        'End If
    End Sub

    Private Sub LbxSpec_DblClick(
    ByVal Cancel As Microsoft.Vbe.Interop.Forms.ReturnBoolean
)
        Dim sp As String
        Dim ky As Object

        sp = LbxSpec.Text
        dcActv.Remove(sp)
        LbxSpec.Text = dcActv.Keys

        Debug.Print("") 'Breakpoint Landing
        ' NOTE: this section res dcActvOps to
        ' the original dcInitList (NOT dcSpecPairs)
        ' and then sequentially re-applies the
        ' active terms remaining in dcActv.
        dcActvOps = dcInitList
        For Each ky In dcActv.Keys
            dcActvOps = dcSpecSubWith(
            CStr(ky), dcActvOps
        )
        Next
        LbxSpecOps.Text = dcActvOps.Keys
        ' this might be more flexibly implemented
        ' in a separate function or procedure

        Debug.Print("") 'Breakpoint Landing
    End Sub

    Private Sub LbxSpecOps_Change()
        'Stop '
    End Sub

    Private Sub LbxSpecOps_Click()
        ' Stop '
    End Sub

    Private Sub LbxSpecOps_Error(ByVal Number As Integer, ByVal Description As Microsoft.Vbe.Interop.Forms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As Microsoft.Vbe.Interop.Forms.ReturnBoolean)
        Stop '
    End Sub

    Private Sub LbxSpec_BeforeDragOver(ByVal Cancel As Microsoft.Vbe.Interop.Forms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As Microsoft.Vbe.Interop.Forms.fmDragState, ByVal Effect As Microsoft.Vbe.Interop.Forms.ReturnEffect, ByVal Shift As Integer)
        Stop '
    End Sub

    Private Sub LbxSpec_BeforeDropOrPaste(ByVal Cancel As Microsoft.Vbe.Interop.Forms.ReturnBoolean, ByVal Action As Microsoft.Vbe.Interop.Forms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As Microsoft.Vbe.Interop.Forms.ReturnEffect, ByVal Shift As Integer)
        Stop '
    End Sub

    Private Sub LbxSpec_Change()
        'Stop '
    End Sub

    Private Sub LbxSpec_Click()
        ' Stop '
    End Sub

    Private Sub LbxSpec_Error(ByVal Number As Integer, ByVal Description As Microsoft.Vbe.Interop.Forms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As Microsoft.Vbe.Interop.Forms.ReturnBoolean)
        Stop '
    End Sub
End Class