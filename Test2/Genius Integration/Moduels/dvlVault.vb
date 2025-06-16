Module dvlVault
    ' '
    ' see module mod3 for other functions of possible use here
    ' named functions here originated there
    ' '

    Public Function ArrayFrom(ls As Object) As Object
        '
        ' ArrayFrom -- return basicObjectArray
        '     from one of several various types
        '     of suppliedObjectValues
        '
        Dim dc As Scripting.Dictionary

        dc = dcOb(obOf(ls))
        If dc Is Nothing Then
            If TypeOf ls Is Object Then
                ArrayFrom = {}
            ElseIf TypeOf ls Is Array Then
                ArrayFrom = ls
            Else
                ArrayFrom = {ls}
            End If
        Else
            ArrayFrom = dc.Keys
        End If
    End Function

    Public Function dcMapFSysVsVault(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim bp As String
        Dim fp As Object
        Dim vp As String

        rt = New Scripting.Dictionary

        bp = vaultBasePath()
        If Len(bp) = 0 Then
            Stop 'for debug/devel
            'not sure what to do
            'with no path found

            'default here is
            '"C:/Doyle_Vault/"
            'but don't want
            'to assume
        End If

        With dc 'dcAiDocComponents(aiDocActive())
            For Each fp In .Keys
                vp = Replace(Replace(
            fp, bp, "$/"
            ), "\", "/"
        )
                With rt
                    If .Exists(fp) Or .Exists(vp) Then
                        Stop
                    Else
                        rt.Add(fp, vp)
                        rt.Add(vp, fp)
                    End If : End With
            Next : End With

        dcMapFSysVsVault = rt
    End Function

    Public Function dcRemapped2vaultPaths(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim bp As String

        rt = New Scripting.Dictionary

        bp = vaultBasePath()
        If Len(bp) = 0 Then
            Stop 'for debug/devel
            'not sure what to do
            'with no path found

            'default here is
            '"C:/Doyle_Vault/"
            'but don't want
            'to assume
        End If

        With dc 'dcAiDocComponents(aiDocActive())
            For Each ky In .Keys
                rt.Add(Replace(Replace(
        ky, bp, "$/"
        ), "\", "/"
    ), .Item(ky))
            Next : End With

        dcRemapped2vaultPaths = rt
    End Function

    Public Function vaultBasePath() As String
        With dcOb(nuILogicIfc().Apply(
    "vltBasePath", New Scripting.Dictionary
    ))
            If .Exists("OUT") Then
                vaultBasePath = .Item("OUT")
            Else
                vaultBasePath = ""
            End If : End With
    End Function

    Public Function vaultPropKeys() As String
        With dcOb(nuILogicIfc().Apply(
    "vltPropKeys", New Scripting.Dictionary
    ))
            If .Exists("OUT") Then
                vaultPropKeys = .Item("OUT")
            Else
                vaultPropKeys = ""
            End If : End With
    End Function

    Public Function dcOfDcByVltPathAndName(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        '
        '
        ' this one should probably call
        ' dcOfDcByNameAndPath against dc
        '
        ' actually, EACH should call
        ' some common function
        ' to perform similar task
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object
        Dim bk As Long
        Dim bp As String
        Dim fn As String

        rt = New Scripting.Dictionary

        With dcRemapped2vaultPaths(dc)
            For Each ky In .Keys
                ar = { .Item(ky)}

                bk = InStrRev(ky, "/")
                If bk = 0 Then
                    Stop
                Else
                    fn = Mid$(ky, bk + 1)
                    bp = Left$(ky, bk - 1)
                    'Stop
                End If

                With rt
                    If Not .Exists(bp) Then
                        .Add(bp, New Scripting.Dictionary)
                    End If

                    ' gp =
                    With dcOb(.Item(bp))
                        .Add(fn, ar(0))
                    End With
                End With


            Next : End With

        dcOfDcByVltPathAndName = rt
    End Function

    Public Function dcOfDcByNameAndPath(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        '
        '
        ' closely related to
        ' dcOfDcByVltPathAndName
        ' (see above)
        '
        Dim rt As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim md As Inventor.Document
        Dim ky As Object
        Dim ar As Object
        Dim bk As Long
        Dim fn As String
        Dim bp As String

        rt = New Scripting.Dictionary

        With dc 'dcRemapped2vaultPaths(dc)
            For Each ky In .Keys
                'ar = {.Item(ky))
                md = aiDocument(obOf(.Item(ky)))
                If md Is Nothing Then
                    Debug.Print("") 'Breakpoint Landing
                Else
                    'Stop

                    bk = InStrRev(ky, "\")
                    If bk = 0 Then
                        Stop
                    Else
                        bp = Left$(ky, bk - 1)
                        fn = Mid$(ky, bk + 1)
                        'Stop
                    End If

                    With rt
                        If Not .Exists(fn) Then
                            .Add(fn, New Scripting.Dictionary)
                        End If

                        ' gp =
                        With dcOb(.Item(fn))
                            .Add(bp, md) 'ar(0)
                        End With
                    End With
                End If
            Next : End With

        dcOfDcByNameAndPath = rt
        'send2clipBdWin10 ConvertToJson(dcOfDcByNameAndPath(dcAiDocComponents(aiDocActive())), vbTab)
    End Function

    Public Function d0g1f4d(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' d0g1f4d - categorize supplied Dictionary
        '     of Part/Assembly components
        '     by Vault Property Values
        '     1 - takes same sort of
        '         Dictionary as d0g1f4c
        '     2 - applies d0g1f4c to it
        '     3 - rekeys the result
        '     4 - transposes its sub Dictionaries
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary

        With DcOfDcRekeyedSecToPri(d0g1f4c(dc))
            For Each ky In .Keys
                rt.Add(ky, DcTransGrouped(dcOb(.Item(ky))))
            Next : End With

        d0g1f4d = rt
    End Function

    Public Function d0g1f4c(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        'Dim ag As Scripting.Dictionary

        Dim ls As Object
        Dim ky As Object

        Dim sd As Scripting.Dictionary
        'Dim nm As Object
        Dim pg As Object
        Dim BillRow As Object
        Dim p2 As String
        'Dim fl As Scripting.File

        rt = New Scripting.Dictionary

        ls = dc.Keys
        With nuILogicIfc()
            For Each ky In ls
                'send2clipBdWin10 ConvertToJson(nuILogicIfc()
                Debug.Print("") 'Breakpoint Landing
                'pg =
                With .Apply("dvl0", NuDcPopulator(
                    ).Setting("PropName", "Name"
                    ).Setting("Value", ky
                ).Dictionary())
                    If .Exists("OUT") Then
                        pg = .Item("OUT")
                    Else
                    End If
                End With
                'PropName", "FolderPath
                '   FullPath
                '   FullName
                '"$/Designs/doyle/(72) G3 Conveyor/I Parts/72-XXX-90403 G3 HD 8IN WRAP DRIVE 6IN END ROLLERS CONVEYOR BELT CRESCENT TOP ASSEMBLY"
                ', vbTab)
                '
                ' REV[2023.03.03.1140]
                ' preceding pg assignment
                ' replaces the one following
                '
                pg = .Apply("vlt04", NuDcPopulator(
            ).Setting("fullname", ky
            ).Dictionary()
        ).Item("OUT") '.DataFor(CStr(nm))
                'or "Full Path"
                For Each BillRow In pg
                    If TypeOf BillRow Is Inventor.NameValueMap Then
                        sd = dcFromAiNameValMap(obOf(BillRow))
                    ElseIf TypeOf BillRow Is Scripting.Dictionary Then
                        sd = BillRow
                    Else
                        Stop
                        sd = Nothing
                    End If

                    If sd Is Nothing Then
                    Else
                        p2 = sd.Item("fullname") '.LocalForm()  'CStr(BillRow)
                        rt.Add(p2, sd) '{nm, fnExt(p2), fl)
                        'With sd
                        '    '.Add("ext", fnExt(p2)
                        '    '.Add("fileObj", fileIfPresent(p2)
                        'End With
                    End If
                Next
            Next : End With

        d0g1f4c = rt
    End Function

    Public Function d0g1f4b(ls As Object) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim nm As Object
        Dim pg As Object
        Dim BillRow As Object
        Dim p2 As String
        Dim fl As Scripting.File

        If TypeOf ls Is Object Then
        ElseIf TypeOf ls Is Array Then
        Else
            rt = d0g1f4b({ls})
            ' NOTE this isn't going to work.
            ' next  rt below is going
            ' to wipe this one right out!
        End If

        rt = New Scripting.Dictionary

        With nuILogicIfc() 'nuIfcVault()
            For Each nm In ArrayFrom(ls)
                pg = .Apply("vlt04",
            NuDcPopulator().Setting(
                "PartNumber", nm
            ).Dictionary()
        ).Item("OUT") '.DataFor(CStr(nm))
                For Each BillRow In pg 'Split(pg, vbCrLf)
                    'Stop
                    If TypeOf BillRow Is Inventor.NameValueMap Then
                        sd = dcFromAiNameValMap(obOf(BillRow))
                    ElseIf TypeOf BillRow Is Scripting.Dictionary Then
                        sd = BillRow
                    Else
                        Stop
                        sd = Nothing
                    End If

                    If sd Is Nothing Then
                    Else
                        p2 = sd.Item("fullname") '.LocalForm()  'CStr(BillRow)
                        rt.Add(p2, sd) '{nm, fnExt(p2), fl)
                        With sd
                            '.Add("ext", fnExt(p2)
                            .Add("fileObj", fileIfPresent(p2))
                        End With
                    End If
                Next
                With rt
                End With
            Next : End With

        Debug.Print("") 'Breakpoint Landing
        If False Then
            'for each k0 in rt.Keys: sd = rt.Item(k0):debug.Print(sd.Item("Folder Path"):next
        End If

        d0g1f4b = rt
        'send2clipBdWin10 ConvertToJson(d0g2f1b(d0g1f4b(Split(nu_FmGetList().AskUser(), vbCrLf))), vbTab)
        'send2clipBdWin10 ConvertToJson(d0g2f1b(d0g1f4b(Split(nu_FmGetList().AskUser(Join(dcOb(dcAiDocCompsByPtNum(aiDocActive()).Item(1)).Keys, vbCrLf)), vbCrLf))), vbTab)
    End Function

    Public Function d0g2f1d(
    dc As Scripting.Dictionary
) As ADODB.Record 'Scripting.Dictionary
        '
        ' d0g2f1d --
        '     derived from d0g2f1b
        '
        Dim rt As Scripting.Dictionary
        Dim rs As ADODB.Record
        Dim xt As Scripting.Dictionary
        Dim ls As Object
        Dim k0 As Object
        Dim k1 As Object
        Dim i0 As Scripting.Dictionary
        Dim fx As String
        Dim ds As String
        Dim pn As String

        rt = New Scripting.Dictionary

        ls = {
        "Part Number",
        "Description", "ext", "fullname"}
        '

        rs = New ADODB.Record
        With rs
            With .Fields : For Each k1 In ls
                    .Append(k1, ADODB.DataTypeEnum.adVarChar, 127)
                Next : End With
            .Open()
        End With

        With dc : For Each k0 In .Keys
                i0 = .Item(k0)

                rs.AddNew

                For Each k1 In ls
                    With i0
                        ds = ""
                        If .Exists(k1) Then
                            If .Item(k1) Is Nothing Then
                            Else
                                ds = .Item(k1)
                            End If
                        End If
                    End With

                    With rs.Fields(k1).Value = ds
                    End With
                Next
            Next : End With
        rs.Filter = ""

        d0g2f1d = rs
        'send2clipBdWin10 d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbCrLf))).GetString(adClipString, , vbTab, vbCrLf)
        'send2clipBdWin10 rsFiltered(d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbCrLf))), "Description <> ''").GetString(adClipString, , vbTab, vbCrLf)
        'send2clipBdWin10 "select i.ItemID Id, ls.pn, ls.ds Description1, i.Description1 oldDescription1 from (values" & vbCrLf & vbTab & "('" & rsFiltered(d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbCrLf))), "Description <> ''").GetString(adClipString, , "', '", "')," & vbCrLf & vbTab & "('") & "', '')" & vbCrLf & ") as ls(pn, ds) inner join vgMfiItems i on ls.pn = i.Item"
    End Function

    Public Function dVg1f1(argIn As Object) As Scripting.Dictionary
        Dim dc As New Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        dc.Add("IN", argIn)
        With nuILogicIfc()
            rt = .Apply("vlt05", dc)
        End With
        dVg1f1 = rt
    End Function

    Public Function dVg2f1(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' dVg2f1 - take Dictionary
        '     of Inventor Documents keyed
        '     to FullFileName as returned
        '     by dcAiDocComponents
        ' return Dictionary of Dictionaries
        '     of Inventor Documents keyed
        '     first to File Name only and
        '     then to ParentFolder Path
        '
        Dim rt As Scripting.Dictionary
        Dim ls As Scripting.Dictionary
        Dim fl As Scripting.File
        Dim ky As Object
        Dim nm As String
        Dim fp As String

        rt = New Scripting.Dictionary

        With dc : For Each ky In .Keys
                fl = fileIfPresent(CStr(ky))
                If Not fl Is Nothing Then
                    With fl
                        nm = .Name
                        fp = .ParentFolder.Path
                    End With

                    With rt
                        If Not .Exists(nm) Then
                            .Add(nm, New Scripting.Dictionary)
                        End If
                        ls = .Item(nm)
                    End With

                    With ls
                        If Not .Exists(fp) Then
                            .Add(fp, fl)
                        End If : End With
                End If
            Next : End With

        dVg2f1 = rt
    End Function

    Public Function dVg3f1(dc As Scripting.Dictionary) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ob As Object
        Dim ky As String

        rt = New Scripting.Dictionary
        With dVg1f1(dVg2f1(dc).Keys)
            If .Exists("OUT") Then
                For Each ob In .Item("OUT")
                    wk = dcOb(ob)

                    If wk Is Nothing Then
                        Stop
                    Else
                        With wk
                            If .Exists("fullname") Then
                                ky = .Item("fullname")
                                With rt
                                    If .Exists(ky) Then
                                        Stop
                                    Else
                                        .Add(ky, wk)
                                    End If : End With
                            Else
                                Stop
                            End If : End With
                    End If
                Next
            Else
                Stop
            End If : End With
        dVg3f1 = rt
        'send2clipBdWin10 ConvertToJson(dVg3f1(dcAiDocComponents(aiDocActive())), vbTab)
    End Function

    Public Function dVg3f2(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' dVg3f2 -    '
        '     NOTE the following:
        '     dcMapFSysVsVault maps the full file names
        '         from the supplied Dictionary's Keys
        '         to their Vault paths/names,
        '         and vice-versa
        '     dVg3f1 returns a Dictionary
        '         keyed to Vault paths/names
        '         which must be translated
        '         to full file names
        '     dcKeysInCommon will return a Dictionary
        '         also keyed to Vault paths/names
        '         containing matching entries
        '         from the results of each
        '         of the prior two
        '     the Dictionary returned is keyed
        '         to the FIRST value in each
        '         entry from dcKeysInCommon,
        '         mapping it to the SECOND value
        '     in this way, each model's full file path
        '         is mapped to its Vault data
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object

        rt = New Scripting.Dictionary

        With DcKeysInCommon(
        dcMapFSysVsVault(dc),
        dVg3f1(dc)
    )
            For Each ky In .Keys
                it = .Item(ky)
                rt.Add(it(0), it(1))
            Next : End With
        dVg3f2 = rt
    End Function

    Public Function dVg3f3(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' dVg3f3 -    given a Dictionary of Inventor Documents
        '             returns
        '
        Dim d2 As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object

        ' rt =
        'New Scripting.Dictionary
        ' d2 =
        'With
        rt = DcKeysCombined(dc, dVg3f2(dc))
        'For Each ky In .Keys
        '    Stop
        'Next: End With

        dVg3f3 = rt
    End Function

    ' END of module dvlVault
    '
    '
    Private Function dvlVault() As String

    End Function
End Module