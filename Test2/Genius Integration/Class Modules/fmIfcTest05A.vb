Imports System.Windows
Imports System.Windows.Forms

Class FmIfcTest05A
    Dim ThisApplication As Inventor.Application
    Public Event Sent(Signal As VbMsgBoxResult)
    Public Event GroupIs(Now As String)
    Public Event ItemIs(Now As String)

    Private obClient As Object
    ' this is meant to hold the "client" Object
    ' expected to make calls to this interface.
    ' not sure if we want to do anything with this yet

    Private WithEvents Fm As FmTest05
    'Private WithEvents Fm As MSForms.UserForm 'fmTest04

    Private WithEvents LbxItems As MSForms.ListBox
    Private WithEvents TbsItemGrps As Forms.TabControl
    Private WithEvents LblPartNum As MSForms.Label
    Private WithEvents LblDesc As MSForms.Label
    Private WithEvents ImgOfItem As Forms.PictureBox
    Private WithEvents CmdOpenItem As Forms.Button
    'Private WithEvents cmdEndCancel     As MSForms.CommandButton
    'Private WithEvents cmdEndSave       As MSForms.CommandButton

    'Private dcActv   As Scripting.Dictionary
    Private itemsFlat As Scripting.Dictionary
    Private allGroups As Scripting.Dictionary
    Private itmPicked As Scripting.Dictionary
    Private gdcActive As Scripting.Dictionary

    Private docActive As Inventor.Document

    Private Const txVersion As String = ""

    Private Const kyInitList As String = "initSpecs"    'key name identifying initial spec list
    Private Const vsnString As String = "Form Test04 Interface A v0.1.0.0 [2022.03.17]"
    ' prior values                    "Form Test04 Interface A v0.0.0.0 [2022.03.03]"
    '                                 ""
    '                                 ""
    '
    '
    '

    Public Function Itself() As FmIfcTest05A
        ' returns this FmIfcTest04A class instance "Itself"
        ' should be HIGHLY useful inside a With context
        Itself = Me
    End Function

    Public Function Usedict(
    Optional Dict As Scripting.Dictionary = Nothing
) As FmIfcTest05A 'fmTest05
        Dim ky As Object
        Dim dp As Long

        If Dict Is Nothing Then
            Return Me.Usedict(New Scripting.Dictionary)

        Else
            itemsFlat = Nothing
            allGroups = Nothing
            dp = dcDepthAiDocGrp(Dict)

            Select Case dp
                Case 1
                    Return withDcFlat(Dict)

                Case 2
                    Return withDcGrpd(Dict)

                    itemsFlat = New Scripting.Dictionary
                    With allGroups
                        For Each ky In .Keys
                            itemsFlat = DcKeysCombined(
dcOb(.Item(ky)), itemsFlat, 1
)
                        Next : End With
                Case Else
                    Return Me
            End Select
        End If

    End Function

    Public Function GroupNow() As String
        GroupNow = Fm.GroupNow()
    End Function

    Public Function InGroup(
    GrpId As String
) As FmIfcTest05A
        If Fm.InGroup(GrpId
    ).GroupNow() = GrpId Then
            ' change succeeded!
        Else
            ' couldn't change!
        End If

        InGroup = Me
    End Function

    Public Function ItemNow() As String
        ItemNow = Fm.ItemNow()
    End Function

    Public Function OnItem(
    ItemId As String
) As FmIfcTest05A
        If Fm.OnItem(ItemId
    ).ItemNow() = ItemId Then
            ' change succeeded!
        Else
            ' couldn't change!
            Stop
        End If

        OnItem = Me
    End Function

    Public Function Show(
    Modal As Object
) As FmIfcTest05A
        Fm.Show(Modal)
        Show = Me
    End Function

    Public Function Hide() As FmIfcTest05A
        Fm.Hide()
        Hide = Me
    End Function

    Public Function SaveAll() As Scripting.Dictionary
        'debug.Print(aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
        ' NOTE[2022.04.13.1224] (copied from ...)
        ' want to initiate 'save all' operation here
        ' or somewhere nearby. note immediate mode
        ' command in comment above
        Dim rtGd As Scripting.Dictionary
        'Dim rtBd As Scripting.Dictionary
        Dim wk As Inventor.Document
        Dim ky As Object
        'Dim mx As Long
        'Dim dx As Long

        rtGd = New Scripting.Dictionary
        ' rtBd = New Scripting.Dictionary

        With itemsFlat
            On Error Resume Next
            For Each ky In .Keys
                wk = .Item(ky)
                With wk
                    If .Dirty Then
                        Err.Clear()
                        .Save2()

                        If Err.Number = 0 Then
                            rtGd.Add(ky, wk)
                        Else
                        End If
                    Else
                        rtGd.Add(ky, wk)
                    End If : End With
            Next
            On Error GoTo 0
        End With

        SaveAll = rtGd
    End Function

    Private Function withDcFlat(
    Dict As Scripting.Dictionary
) As FmIfcTest05A 'fmTest05
        itemsFlat = DcCopy(Dict)
        withDcFlat = withDcGrpd(
    dcAiDocGrpsByForm(itemsFlat))
    End Function

    Private Function withDcGrpd(
    Dict As Scripting.Dictionary
) As FmIfcTest05A 'fmTest05
        Dim ky As Object
        Dim ls As TabControl.TabPageCollection

        itmPicked = New Scripting.Dictionary
        docActive = Nothing

        ls = TbsItemGrps.TabPages
        ls.Clear()

        allGroups = Dict

        With allGroups : For Each ky In Split(
        "MAYB DBAR SHTM ASSY PRCH HDWR"
    ) 'instead of .Keys, to
                '  ensure preferred order
                If .Exists(ky) Then
                    With dcOb(.Item(ky))
                        ''  check for group members
                        If .Count > 0 Then 'select the first
                            itmPicked.Add(ky, .Keys(0))
                            ls.Add(ky) 'this will want more development later
                        Else 'paint it blank
                            'itmPicked.Add(ky, ""
                            ''  actually, don't do anything
                            ''  like, don't even add the tab
                            ''  if nothing's going to be
                            ''  there anyway
                        End If : End With : End If
            Next : End With
        withDcGrpd = Me
    End Function

    Private Function gpActive() As String
        Dim tb As TabPage

        With TbsItemGrps
            tb = .TabPages.Item(.Name)
        End With
        gpActive = tb.Name
    End Function

    Private Sub CmdOpenItem_Click()
        Dim ck As VbMsgBoxResult
        Dim pn As String

        If docActive Is Nothing Then
            '
        Else
            With docActive
                pn = .Propertys.Item(gnDesign).Item(pnPartNum).Text

                ' REV[2022.05.06.1142]
                ' added check for Part Document to avoid error
                ' trying to edit Material for Assembly Documents.
                If TypeOf docActive Is Inventor.PartDocument Then
                    ck = MsgBox(Join({
                    "Would you rather just edit",
                    "material for " & pn & "?", "",
                    "(No to go ahead and open)"
                }, vbCrLf),
                    vbYesNoCancel + vbQuestion,
                    "Edit Material?"
                )
                Else
                    ck = vbNo
                End If

                If ck = vbCancel Then
                    Stop
                ElseIf ck = vbYes Then
                    ' NOTE[2022.05.06.1143]
                    ' this section throws an error
                    ' if the Document is an assembly.
                    ' REV[2022.05.06.1142] above adds
                    ' a check to prevent this branch
                    ' from being taken in that case.
                    'Debug.Print(ConvertToJson(
                    askUserForPartMatlUpdate(
                itemsFlat.Item(pn) & vbTab)
                Else
                    If .Open Then
                        ck = vbYes
                    Else
                        ck = MsgBox(Join({
                            "Document " & pn,
                            "is not presently open.",
                            "Go ahead and open it?"
                        }, vbCrLf),
                        vbYesNo, "Open " & pn & "?"
                    )
                    End If

                    If ck = vbYes Then
                        On Error Resume Next

                        Err.Clear()
                        .Activate()

                        If Err.Number = 0 Then
                        Else
                            If ThisApplication.Documents.Open(
                            .FullDocumentName, True
                        ) Is docActive Then
                                Debug.Print("")  'Breakpoint Landing
                            Else
                                Stop
                                Debug.Print("")  'Breakpoint Landing
                            End If

                            Err.Clear()
                            .Activate()
                            If Err.Number Then Stop

                            Debug.Print("")  'Breakpoint Landing
                        End If

                        On Error GoTo 0
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub Fm_GroupIs(Now As String)
        '
        RaiseEvent GroupIs(Now)
    End Sub

    Private Sub Fm_ItemIs(Now As String)
        '
        RaiseEvent ItemIs(Now)
    End Sub

    Private Sub Fm_Sent(Signal As VbMsgBoxResult)
        'Public Event Sent(Signal As VbMsgBoxResult)
        '
        Dim ck As VbMsgBoxResult

        If obClient Is Nothing Then
            ck = vbRetry

            Select Case Signal
                Case vbOK
                    'ck = MsgBox(Join({ _
                    '    "Save and Close", _
                    '    "Operation Selected" _
                    '}, vbCrLf), _
                    '    vbYesNoCancel, _
                    '    "Save Documents?" _
                    ')

                    'debug.Print(aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
                    ' NOTE[2022.04.13.1224]
                    ' want to initiate 'save all' operation here
                    ' or somewhere nearby. note immediate mode
                    ' command in comment above
                    With DcKeysMissing(
                itemsFlat,
                SaveAll()
            )
                        If .Count > 0 Then
                            ck = MsgBox(Join({
                                "Errors encountered trying to",
                                     "save the following Documents:",
                                    vbTab & TxDumpLs(.Keys,
                                vbCrLf & vbTab
                                                     ),
                                 "",
                             "Close anyway?"
                                        }, vbCrLf),
                        vbYesNoCancel,
                        "Errors on Save!"
                    )
                        Else
                            ck = vbYes
                        End If
                    End With
                Case vbAbort, vbCancel
                    ck = MsgBox(Join({
                "Cancel",
                "Operation",
                "Selected"
            }, vbCrLf),
                vbYesNoCancel,
                "Finished?"
            )
                    'Case Else
                Case Else
            End Select

            If ck = vbCancel Then
                Stop
            ElseIf ck = vbYes Then
                Hide()
            End If
        Else
            Stop
            RaiseEvent Sent(Signal)
        End If
    End Sub

    Private Sub LbxItems_Change()
        Dim pn As String
        Dim pc As stdole.StdPicture

        'Stop
        pn = LbxItems.Name
        With gdcActive
            If .Exists(pn) Then
                docActive = aiDocument(.Item(pn))
                With docActive
                    With .Propertys.Item(gnDesign)
                        LblPartNum.Text = .Item(pnPartNum).Text
                        LblDesc.Text = .Item(pnDesc).Text
                    End With

                    On Error Resume Next
                    pc = .Thumbnail

                    If Err.Number = 0 Then
                    Else
                        pc = Nothing
                    End If

                    ImgOfItem.Image = pc
                    On Error GoTo 0
                End With
                'docActive.Thumbnail
            Else
                docActive = Nothing
                LblPartNum.Text = "(select part)"
                LblDesc.Text = ""
            End If
        End With

        ' REV[2022.03.17.1348]
        '     addApplication.DoEvents() for rapid visual feedback
        '     (see TbsItemGrps_Change for details)
        Application.DoEvents()
    End Sub

    Private Sub TbsItemGrps_Change()
        Dim tb As Forms.TabControl
        Dim nm As String

        'With TbsItemGrps
        '     tb = .Tabs.Item(.Text)
        'End With
        nm = gpActive() 'tb.Name

        gdcActive = dcOb(
        allGroups.Item(nm)
    )
        With LbxItems
            .Items.Clear()
            .Items.AddRange(gdcActive.Keys)
            .SelectedItem = itmPicked.Item(nm)
        End With

        ' REV[2022.03.17.1348]
        '     addingApplication.DoEvents() steps to various
        '     Change Event handlers to try to
        '     ensure timely visual feedback
        '     to the User in-process
        Application.DoEvents()
    End Sub

    Private Sub Class_Initialize()
        ' Fm =
        With New FmTest05
            'End With
            'With Fm
            LbxItems.Text = .LbxItems.Text
            TbsItemGrps = .TbsItemGrps
            LblPartNum = .LblPartNum
            LblDesc = .LblDesc
            ImgOfItem = .ImgOfItem
            CmdOpenItem = .CmdOpenItem
            ' cmdEndCancel = .cmdEndCancel
            ' cmdEndSave = .cmdEndSave

            Fm = .Holding(Me)
        End With

        ' dcActv = New Scripting.Dictionary
        allGroups = New Scripting.Dictionary
    End Sub

    Private Sub Class_Terminate()
        allGroups = Nothing

        Fm = Fm.Dropping(Me)

        ' cmdEndSave = Nothing
        ' cmdEndCancel = Nothing
        CmdOpenItem = Nothing
        ImgOfItem = Nothing
        LblDesc = Nothing
        LblPartNum = Nothing
        TbsItemGrps = Nothing

        LbxItems = Nothing
        LbxItems = Nothing

        Fm = Nothing
    End Sub

End Class