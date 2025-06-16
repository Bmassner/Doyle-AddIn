Module dvl1

    ' Initial Purpose: Begin design of new Genius Properties Generator/Populator

    Public Function d1g0f0() As Object
        d1g0f0 = 0
    End Function

    Public Function d1g4f0(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary

        With dcRemapByPtNum(
        dcAiDocComponents(AiDoc)
    )
            For Each ky In .Keys
                rt.Add(ky, dcProps4genius(
                aiDocument(.Item(ky)), , 0
            ))
                rt.Add(ky, dcAiPropValsFromDc(
                dcOfPropsInAiDoc(
                    aiDocument(.Item(ky))
                )
            ))
            Next
        End With

        d1g4f0 = rt
        'Debug.Print(dumpLsKeyVal(d1g4f1(d1g4f0(InventorApp.ActiveDocument)), " - ")
        'send2clipBd ConvertToJson(d1g4f0(InventorApp.ActiveDocument), vbTab)
    End Function

    Public Function d1g4f1(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                rt.Add(ky, dcOb(.Item(ky)).Count)
            Next
        End With
        d1g4f1 = rt
    End Function

    Public Function d1g4f2(md As String, InvProperty As String) As String
        Dim vbc As VBIDE.VBComponent
        Dim rt As Object
        Dim ls As Object
        Dim mx As Long
        Dim dx As Long
        Dim ck As String
        Dim Doc As Inventor.Document

        vbc = Doc.VBAProject.InventorVBAComponents.Item(md).VBComponent
        With vbc.CodeModule
            ls = Split(.Lines(
            .ProcBodyLine(InvProperty, vbext_ProcKind.vbext_pk_Proc),
            .ProcCountLines(InvProperty, vbext_ProcKind.vbext_pk_Proc)
        ), vbCrLf)
            mx = UBound(ls)
            For dx = LBound(ls) To mx
                ck = Trim$(ls(dx))
                If Left$(ck, 1) = "'" Then
                    rt = rt & Mid$(ck, 2) & vbCrLf
                End If
            Next
        End With
        d1g4f2 = rt

        'Debug.Print(d1g4f2("dvl1", "d1g4f2")
        'Debug.Print(d1g4f2("zzCsv000", "zc0g0f1")
        'Debug.Print(d1g4f2("zzCsv000", "zc0g0f2")
    End Function

    Public Function d1g4f3(
    hdr As String, dlm As String
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ls As Object
        Dim mx As Long
        Dim dx As Long

        rt = New Scripting.Dictionary
        ls = Split(hdr, dlm)
        mx = UBound(ls)
        For dx = LBound(ls) To mx
            rt.Add(dx, ls(dx))
            rt.Add(ls(dx), dx)
        Next
        rt.Add("", dlm)
        d1g4f3 = rt
    End Function

    Public Function d1g4f4(
    dc As Scripting.Dictionary,
    tx As String
) As Scripting.Dictionary
        Dim hd As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ls As Object
        Dim mx As Long
        Dim dx As Long
        Dim dlm As String

        rt = New Scripting.Dictionary
        With dc
            hd = dcOb(.Item(""))
            With hd
                dlm = .Item("")

                ls = Split(tx, dlm)
                mx = UBound(ls)
                For dx = LBound(ls) To mx
                    If .Exists(dx) Then
                        rt.Add(.Item(dx), ls(dx))
                    End If
                    rt.Add(dx, ls(dx))
                Next
            End With

            dx = .Count
            Do While .Exists(dx)
                dx = 1 + dx
            Loop
            .Add(dx, rt)
        End With

        d1g4f4 = dc
    End Function

    Public Function d1g4f5(tx As String,
    dc As Scripting.Dictionary,
    Optional bk As String = vbCrLf
) As Scripting.Dictionary
        Dim ck As Long

        If Len(tx) > 0 Then
            ck = InStr(tx, bk)
            If ck > 0 Then
                d1g4f5 = d1g4f5(
                Mid$(tx, ck + Len(bk)),
                d1g4f4(dc, Left$(tx, ck - 1)),
            bk)
            Else
                d1g4f5 = d1g4f4(dc, tx)
            End If
        Else
            d1g4f5 = dc
        End If
    End Function

    Public Function d1g4f6(tx As String,
    Optional dlm As String = ",",
    Optional bk As String = vbCrLf
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim hdr As String
        Dim ck As Long

        rt = New Scripting.Dictionary
        ck = InStr(1, tx, bk)
        If ck > 0 Then
            hdr = Left$(tx, ck - 1)
            rt.Add("", d1g4f3(hdr, dlm))
            rt = d1g4f5(
            Mid$(tx, ck + Len(bk)),
            rt, bk
        )
        Else
            'handle straight text here?
        End If

        d1g4f6 = rt
        ''Debug.Print(ConvertToJson(d1g4f6(d1g4f2("zzCsv000", "zc0g0f2"), "|"), "  ")
    End Function

    Public Function d1g1f0() As Object
        Dim dc As Scripting.Dictionary

        dc = dcRemapByPtNum(
        dcAiDocComponents(aiDocActive())
    )
        Debug.Print(TxDumpLs(dc.Keys))

        d1g1f0 = dc.Keys
        'Debug.Print(txDumpLs(dcRemapByPtNum(dcAiDocComponents(InventorApp.ActiveDocument)).Keys)
    End Function

    Public Function d1g1f2(PartDesc As Long,
    Optional fc As Long = 2
) As String
        Dim ct As Long
        Dim dv As Long
        Dim rt As String

        If fc > PartDesc Then
            d1g1f2 = ""
        Else
            ct = 0
            dv = PartDesc
            Do Until dv Mod fc > 0
                ct = 1 + ct
                dv = dv \ fc
            Loop
            If ct > 0 Then
                rt = CStr(fc) & "," _
               & CStr(ct) & vbCrLf
            Else
                rt = ""
            End If
            d1g1f2 = rt & d1g1f2(dv, 1 + fc)
        End If
        '
        '
    End Function

    Public Function d1g1f3(PartDesc As Long,
    Optional tHdr As String = "Factor,Power",
    Optional fSep As String = ",",
    Optional lSep As String = vbCrLf
) As String
        Dim fc As Long
        Dim ct As Long
        Dim dv As Long
        Dim wk As String
        Dim rt As String

        rt = tHdr
        ct = d1g1f3b2(PartDesc)
        dv = PartDesc \ (2 ^ ct)
        If ct > 0 Then rt _
        = rt & lSep _
        & "2" & fSep _
        & CStr(ct)

        fc = 3
        Do Until fc > dv
            ct = 0
            Do Until dv Mod fc > 0
                ct = 1 + ct
                dv = dv \ fc
            Loop

            If ct > 0 Then rt _
            = rt & lSep _
            & CStr(fc) & fSep _
            & CStr(ct)

            fc = fc + 2
        Loop
        d1g1f3 = rt
        '
        'Debug.Print(d1g1f3(489168, "1&", " ^ ", " * ")
        '   Some odd behavior on this one.
        '   The result, as printed:
        '       1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
        '   can be processed in immediate mode
        '   to return the original value.
        '   However, there MUST be a space BEFORE
        '   the caret (^) in order for it to be
        '   interpreted correctly. Otherwise, it
        '   seems to be recognized as some sort
        '   of separator, as seen in these examples:
        '
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", "^", "*")
        'Debug.Print(1& * 2^, 4 * 3^, 2 * 43^, 1 * 79^, 1
        ' 2  12  86  79  1
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", "^", " * ")
        'Debug.Print(1& * 2^, 4 * 3^, 2 * 43^, 1 * 79^, 1
        ' 2  12  86  79  1
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", " ^", "*")
        'Debug.Print(1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
        '489168
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", "^ ", "*")
        'Debug.Print(1& * 2^, 4 * 3^, 2 * 43^, 1 * 79^, 1
        ' 2  12  86  79  1
        '
        '   It LOOKS like the caret is treated as some sort
        '   of type indicator, like !, #, % and & are used
        '   to indicate single, double, integer and long
        '   values. This would also seem to be supported
        '   by the presence of semicolons in the samples
        '   above. Those were NOT PRINTED in immediate mode!
        '   Instead, they were almost certainly added when
        '   copied into the editor. They will be left as
        '   they are, here, to show the interpreter does this.
        '
        '   Follow-up confirms speculation: the caret is
        '   indeed a type indicator, but ONLY in 64-bit VBA.
        '   It indicates a LongLong value. See references:
        '   https://stackoverflow.com/questions/51264287/
        '       vba-power-operator-not-working-as-expected-in-64-bit-vba
        '   https://docs.microsoft.com/en-us/office/vba/language/
        '       reference/user-interface-help/longlong-data-type
        '
        '   Further confirmation from testing in Excel VBA,
        '   which here at Doyle is still a 32-bit installation:
        '
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", "^ ", "*")
        'Debug.Print(1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
        '489168
        'Debug.Print("Debug.Print(" & d1g1f3(489168, "1&", "^", "*")
        'Debug.Print(1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
        '489168
        '
        '   These are the same examples which produced lists
        '   of numbers in Inventor VBA, and were "fixed"
        '   on pasting into this code before commenting.
        '   Actually, these strings WERE fixed, but only
        '   by inserting spaces where none were previously.
        '   In any case, this looks like mystery solved.
        '   On top of that, however, it would appear that
        '   Inventor VBA is a 64-bit implementation, which
        '   means it SHOULD support the LongLong data type.
        '   THAT might prove interesting to explore...
        '
        '
    End Function

    Public Function d1g1f3b2(PartDesc As Long) As Long
        Dim fc As Long
        Dim ct As Long
        Dim dv As Long
        Dim wk As String
        Dim rt As String

        ct = 0
        dv = PartDesc
        Do Until 1 And dv
            ct = 1 + ct
            dv = dv \ 2
        Loop

        d1g1f3b2 = ct
        '
        '
    End Function

    Public Function fcPrime(n As Long _
    , Optional rt As String = "" _
    , Optional ls As String = "BCEGKMQSW"
) As String
        Dim nx As Long
        Dim md As Long
        Dim fc As Long
        Dim ct As Long

        If n > 0 Then
            If n = 1 Then
                fcPrime = rt
            ElseIf Len(ls) > 0 Then
                nx = n
                fc = 31 And Asc(ls)
                md = nx Mod fc
                ct = 0
                Do Until md > 0
                    ct = 1 + ct
                    nx = nx \ fc
                    md = nx Mod fc
                Loop
                fcPrime = fcPrime(nx,
                rt & Chr(48 + ct),
                Mid$(ls, 2)
            )
            Else
                fcPrime = rt & "|" & CStr(n)
            End If
        Else
            fcPrime = ""
        End If
    End Function

    Public Function fcCommon(s0 As String, s1 As String) As String
        Dim n0 As Long
        Dim n1 As Long

        n0 = Len(s0) * Len(s1)
        If n0 > 0 Then
            n0 = Asc(s0)
            n1 = Asc(s1)
            fcCommon = Chr(IIf(n0 > n1, n1, n0)) _
            & fcCommon(Mid$(s0, 2), Mid$(s1, 2))
        Else
            fcCommon = ""
        End If
    End Function

    Public Function fcProduct(s As String _
    , Optional ls As String = "BCEGKMQSW"
) As Long
        If Len(ls) > 0 Then
            If Len(s) > 0 Then
                fcProduct = ((31 And Asc(ls)) _
               ^ (15 And Asc(s))) _
               * fcProduct(
                    Mid$(s, 2),
                    Mid$(ls, 2)
               )
            Else
                fcProduct = 1
            End If
        Else
            fcProduct = -1
        End If
    End Function

    Public Function fcMaxComm(n0 As Long, n1 As Long) As Long
        '
        ' fcMaxComm -- Return Greatest Common Factor
        '
        fcMaxComm = fcProduct(fcCommon(
        fcPrime(n0), fcPrime(n1)
    ))
        '
        'For m = 2 To 49: For n = 1 + m To 49: cf = fcMaxComm((m), (n)): Debug.Print(IIf(cf > 1, CStr(n) & "<" & CStr(cf) & ">" & CStr(m) & vbCrLf, "");: Next: Next
        '
    End Function

    Public Function gcfTest() As Long
        '
        ' gcfTest -- Test GCF Function fcMaxComm
        '
        Dim rt As Long
        Dim n0 As Long
        Dim n1 As Long
        Dim nd As Long
        Dim gf As Long

        rt = 0
        For n0 = 4 To 49
            nd = n0 - 1
            For n1 = 2 To nd
                gf = fcMaxComm(n0, n1)
                If (n0 Mod gf) _
            + (n1 Mod gf) _
            > 0 Then
                    rt = 0
                    Debug.Print("")
                End If
            Next
        Next
    End Function

    Public Function tbPrimesWithSquare(
    Optional ct As Long = 100000
    ) As Long(,)
        '
        ' tbPrimesWithSquare
        '
        ' Generate a table of primes
        ' and their corresponding squares
        '
        Dim dbg As Long
        Dim p(,) As Long
        Dim n As Long
        Dim mp As Long
        Dim dx As Long
        Dim mx As Long
        Dim nx As Long

        Dim d0 As Double
        Dim d1 As Double

        ReDim p(1, ct)
        mx = UBound(p, 2)
        p(0, 0) = 2 : p(1, 0) = 4
        nx = 1
        n = 3

        d0 = Timer
        Do
            dx = 0
            mp = 1
            Do
                mp = n Mod p(0, dx)
                If n > p(1, dx) Then
                    dx = dx + 1
                Else
                    dx = nx
                End If
            Loop While mp * p(0, dx) > 0

            If mp > 0 Then
                If p(0, dx) = 0 Then
                    p(0, dx) = n
                    On Error Resume Next
                    p(1, dx) = n * n
                    If Err.Number = 0 Then
                        nx = dx + 1
                    Else
                        nx = mx + 1
                    End If
                    On Error GoTo 0
                Else
                    Stop
                End If
            End If

            n = 1 + n
        Loop Until nx > mx
        d1 = Timer - d0

        dbg = 0 'Change to 1 for debug mode
        If dbg Then
            Debug.Print(1000 * d1) '- d0
            Stop
        End If

        tbPrimesWithSquare = p
    End Function

    Public Function d1g1f7() As Long
        Dim d0 As Double
        Dim d1 As Double
        Dim ur As VbMsgBoxResult

        d0 = Timer
        ur = MsgBox("", vbOKOnly, "")
        d1 = Timer - d0

        Stop
    End Function

    Public Function bcCtCommFac(dc As Scripting.Dictionary) As Long
        Dim ls As Object
        Dim rt As Long
        Dim mx As Long
        Dim dx As Long

        With dc
            If .Count > 0 Then
                ls = .Keys
                mx = UBound(ls)
                rt = CLng(.Item(ls(0)))
                dx = 1

                Do
                    rt = fcMaxComm(rt,
                    CLng(.Item(ls(dx)))
                )
                    If rt = 1 Then
                        dx = 1 + mx
                    Else
                        dx = 1 + dx
                    End If
                Loop Until dx > mx
            Else
                rt = 1
            End If
        End With

        bcCtCommFac = rt
        '
        '
    End Function

    Public Function dcBoltConn1byGCF(
    dc As Scripting.Dictionary,
    Optional fc As Long = 0
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ct As Long

        If fc > 0 Then
            rt = New Scripting.Dictionary
            With dc
                For Each ky In .Keys
                    ct = CLng(.Item(ky))
                    rt.Add(ky, ct \ fc)
                Next
            End With
        Else
            rt = dcBoltConn1byGCF(dc, bcCtCommFac(dc))
        End If

        dcBoltConn1byGCF = rt
        '
        'Debug.Print(dumpLsKeyVal(dcBoltConn1byGCF(dcOfBoltConn02(ad))): Debug.Print
        '
    End Function

    Public Function aiDocProp(
    AiDoc As Inventor.Document, propName As String,
    Optional prop As String = gnCustom
) As Inventor.Property
        '
        ' Proposed Name: aiDocProp
        '
        Dim rt As Inventor.Property

        If AiDoc Is Nothing Then
            aiDocProp = Nothing
        Else
            With AiDoc.Propertys
                If .PropertyExists(prop) Then
                    '.Item(prop).GetPropertyInfo()
                    aiDocProp = aiGetProp(
                    .Item(prop),
                    propName, 0
                )
                Else
                    aiDocProp = Nothing
                End If
            End With
        End If
    End Function

    Public Function aiDocPropVal(
    AiDoc As Inventor.Document, propName As String,
    Optional prop As String = gnCustom
) As Object
        '
        ' Proposed Name: aiDocPropVal
        '
        aiDocPropVal = aiPropVal(aiDocProp(
        AiDoc, propName, prop
    ))
    End Function

    Public Function d1g2f0() As Object
        '
        d1g2f0 = ""
    End Function

    Public Function d1g2f1(AiDoc As Inventor.Document) As Object
        '
        '
        '
        Dim PartDoc As Long
        Dim sc As Long

        PartDoc = 0 : sc = 0

        If InStr(1, AiDoc.FullFileName,
        "\Doyle_Vault\Designs\purchased\"
    ) > 0 Then PartDoc = PartDoc Or 1 : sc = sc + 1

        If InStr(1,
        "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
        "|" & CStr(aiDocPropVal(
            AiDoc, pnFamily, gnDesign
        )) & "|"
    ) > 0 Then PartDoc = PartDoc Or 2 : sc = sc + 1

        d1g2f1 = ""
    End Function

    Public Function d1g2f2(
    AiDoc As Inventor.AssemblyDocument
    ) As Scripting.Dictionary
        '
        '
        '
        'Dim dc As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim oc As Inventor.ComponentOccurrence
        'Dim rs As ADODB.Record ' Unused variable removed
        Dim ob As Inventor.Document
        'Dim ky As Object ' Unused variable removed
        'Dim pn As String ' Unused variable removed
        Dim PartDoc As Long
        Dim sc As Long
        Dim dx As Long
        'Dim bs As Inventor.BOMStructureEnum ' Unused variable removed

        rt = New Scripting.Dictionary
        dx = rt.Count
        'PartDoc = 0: sc = 0

        With AiDoc.ComponentDefinition
            For Each oc In .Occurrences
                With oc
                    ob = aiDocument(
                    .Definition.Document
                )
                    If .BOMStructure _
                    = BOMStructureEnum.kPhantomBOMStructure _
                Then
                        If .DefinitionDocumentType _
                        = DocumentTypeEnum.kAssemblyDocumentObject _
                    Then
                            With aiDocAssy(ob)
                                '.ComponentDefinition.BOMStructure
                                If .DocumentInterests.HasInterest(guidDesignAccl) Then
                                Else
                                End If
                            End With
                        Else
                        End If
                    Else
                    End If
                    With .Definition
                        With ob
                        End With
                    End With
                End With
            Next
        End With
        If InStr(1, AiDoc.FullFileName,
        "\Doyle_Vault\Designs\purchased\"
    ) > 0 Then
            PartDoc = PartDoc Or 1
            sc = sc + 1
        End If

        If InStr(1,
        "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
        "|" & CStr(aiDocPropVal(
            AiDoc, pnFamily, gnDesign
        )) & "|"
    ) > 0 Then
            PartDoc = PartDoc Or 2
            sc = sc + 1
        End If

        d1g2f2 = rt
    End Function

    Public Function d1g2f3(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        rt = New Scripting.Dictionary
        rt.Add(AiDoc.DocumentType, AiDoc)
        rt.Add(AiDoc.DocumentSubType.DocumentSubTypeID, AiDoc)
        d1g2f3 = rt
    End Function

    Public Function d1g3f0() As Object
        '
        '
        '
        d1g3f0 = ""
    End Function

    Public Function d1g3f1(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary
        '
        ' d1g3f1 --
        '
        ' Generate counts of components
        ' in supplied Assembly, adding
        ' a sub-Dictionary for any
        ' "phantom" component recognized
        ' as either a Bolted Connection,
        ' or an Assembly of entirely
        ' Content Center components.
        '
        ' (the latter case addresses
        ' an issue encountered with
        ' just such an Assembly)
        '
        Dim rt As Scripting.Dictionary
        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.Document
        Dim nm As String
        Dim bc As Scripting.Dictionary
        Dim ar As Object

        rt = New Scripting.Dictionary
        If ad Is Nothing Then 'we got nothing to work with
        Else
            For Each oc In ad.ComponentDefinition.Occurrences
                sd = aiDocument(oc.Definition.Document)
                nm = sd.FullDocumentName

                With rt
                    If .Exists(nm) Then
                        ar = .Item(nm)
                        Debug.Print("")
                        ar(1) = ar(1) + 1
                        .Item(nm) = ar
                        Debug.Print("")
                    Else
                        bc = Nothing
                        If oc.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                            With sd.DocumentInterests
                                If .HasInterest(guidDesignAccl) Then
                                    Debug.Print("FOUND Design Accelerator")
                                    Debug.Print("")
                                    bc = d1g3f1(sd) 'New Scripting.Dictionary
                                Else
                                    Debug.Print("FOUND Phantom Assembly")
                                    Debug.Print(vbTab & "NOT Design Accelerator")
                                    Debug.Print("")
                                    bc = dcIfDesignAccel(d1g3f1(sd))
                                    If bc Is Nothing Then
                                    Else
                                        Debug.Print(vbTab & "but ALL Members ARE Content Center")
                                        Debug.Print(vbTab & "so WILL Process as Such")
                                    End If
                                End If
                            End With

                            Debug.Print(vbTab & sd.FullDocumentName)
                            Debug.Print(vbTab & aiDocPropVal(sd, pnPartNum, gnDesign))
                        End If
                        rt.Add(nm, {sd, 1, bc})
                    End If
                End With
            Next
        End If

        d1g3f1 = rt
        'For Each oc In sd.ComponentDefinition.Occurrences: Debug.Print(oc.Name, "|", oc.RangeBox.MinPoint.X, "|", oc.RangeBox.MinPoint.Y, "|", oc.RangeBox.MinPoint.Z, "|", oc.RangeBox.MaxPoint.X, "|", oc.RangeBox.MaxPoint.Y, "|", oc.RangeBox.MaxPoint.Z: Next
    End Function

    Public Function d1g3f2(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim sk As Object
        Dim ar As Object
        Dim sa As Object
        Dim ct As Long

        rt = New Scripting.Dictionary
        With d1g3f1(ad)
            For Each ky In .Keys
                ar = .Item(ky)
                If ar(2) Is Nothing Then
                    rt.Add(ky, ar) '(1)
                Else
                    With dcOb(ar(2))
                        For Each sk In .Keys
                            sa = .Item(sk)
                            ct = ar(1) * sa(1)
                            With rt
                                If .Exists(sk) Then 'some already counted
                                    'so need to add to existing total

                                    'ct = ct + sa(1) '.Item(sk)
                                    sa(1) = ct + .Item(sk)(1)
                                    'got type mismatch here, and fixed
                                    'but not sure fix is correct

                                    .Item(sk) = sa 'ct
                                Else 'this is a whole new component
                                    'so just add its count to the list

                                    sa(1) = ct
                                    .Add(sk, sa) 'ct .Item(sk)
                                End If
                            End With
                        Next
                    End With
                End If
            Next
        End With

        d1g3f2 = rt
    End Function

    Public Function d1g3f3(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object

        rt = New Scripting.Dictionary
        With d1g3f2(ad)
            For Each ky In .Keys
                ar = .Item(ky)
                With aiDocument(obOf(ar(0))).Propertys
                    With .Item(gnDesign).Item(pnPartNum)
                        rt.Add(.Value, ar)
                    End With
                End With
            Next
        End With

        d1g3f3 = rt
    End Function

    Public Function d1g3f4(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object

        rt = New Scripting.Dictionary
        With d1g3f3(ad)
            For Each ky In .Keys
                ar = .Item(ky)
                rt.Add(ky, ar(1))
            Next
        End With

        d1g3f4 = rt
    End Function

    Public Function d1g3f5(
    ad As Inventor.AssemblyDocument,
    Optional incTop As Long = 0
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim sd As Inventor.AssemblyDocument
        Dim ky As Object
        'Dim ar As Object

        rt = New Scripting.Dictionary
        With dcRemapByPtNum(
        dcAiDocComponents(ad, , incTop, 1)
    )
            For Each ky In .Keys
                ad = aiDocAssy(obOf(.Item(ky)))
                If ad Is Nothing Then 'skip it
                Else
                    ''  Previous test, just for Bolted Connection
                    'With ad.DocumentInterests
                    'If .HasInterest(guidDesignAccl) Then
                    ''  Replaced with test for ALL Phantom (below)

                    With ad.ComponentDefinition
                        If .BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                            'Phantom -- don't add to Dictionary
                            Debug.Print("")
                        Else
                            rt.Add(ky, d1g3f4(ad))
                        End If
                    End With
                End If
            Next
        End With

        d1g3f5 = rt
    End Function

    Public Function d1g3f6(
    ad As Inventor.AssemblyDocument,
    Optional incTop As Long = 0
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object

        rt = New Scripting.Dictionary
        With d1g3f5(ad, incTop)
            For Each ky In .Keys
                rt.Add(ky & "|" & ky & "|1",
            vbCrLf & ky & "|" & DumpLsKeyVal(
                dcOb(.Item(ky)), "|",
                vbCrLf & ky & "|"
            ))
                'old key "[" & ky & "]"
                'replaced with pipe-delimited
                'record to fit in better
            Next
        End With

        d1g3f6 = rt
        'Debug.Print(dumpLsKeyVal(d1g3f6(aiDocument(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences.Item(1).Definition.Document)), "")
    End Function

    Public Function d1g3f7(
    ad As Inventor.AssemblyDocument,
    Optional incTop As Long = 0
) As String
        d1g3f7 = "Product|ItemCode|Qty" & vbCrLf _
    & DumpLsKeyVal(d1g3f6(ad, incTop), "")
        '
        'Debug.Print(d1g3f7(aiDocActive())
        '
    End Function

    Public Function dcOfBoltConnReLabeled(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim rd As Scripting.Dictionary
        Dim wd As Scripting.Dictionary
        Dim ky As Object
        Dim pn As String
        Dim fn As Object
        Dim ad As Inventor.Document
        Dim InvProperty As Inventor.Property

        rt = New Scripting.Dictionary

        With dc
            For Each ky In .Keys
                wd = dcOb(.Item(ky))
                rd = New Scripting.Dictionary

                pn = InputBox(
                Join({
                    "Part Number proposed",
                    "  for subassemblies",
                    Join(wd.Keys,
                        vbCrLf & "    "
                    ),
                    "",
                    "Modify as necessary,",
                    "then click OK to confirm.",
                    ""
                }),
                "Verify BC Part Number", CStr(ky)
            )
                With wd
                    For Each fn In .Keys
                        ad = aiDocument(
                        obOf(.Item(fn))
                    )
                        InvProperty = aiDocProp(ad,
                        pnPartNum, gnDesign
                    )
                        If InvProperty Is Nothing Then
                            'nothing to do
                        Else
                            InvProperty.Text = pn
                            Debug.Print("")

                            rd.Add(fn, ad)
                        End If
                    Next
                End With

                If rd.Count > 0 Then
                    rt.Add(pn, rd)
                End If
            Next
        End With

        dcOfBoltConnReLabeled = rt
        '
        'Debug.Print(txDumpLs(dcOfBoltConnReLabeled(dcOfBoltConnIn(aiDocAssy(aiDocActive()))).Keys)
        '
    End Function

    Public Function dcOfBoltConnIn(
    ad As Inventor.AssemblyDocument,
    Optional incTop As Long = 0
) As Scripting.Dictionary
        With dcAiDocComponents(ad, , incTop, 1)
        End With
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim wd As Scripting.Dictionary
        Dim sd As Inventor.AssemblyDocument
        Dim ky As Object
        Dim pn As String
        Dim dn As String
        'Dim ar As Object

        rt = New Scripting.Dictionary
        With dcAiDocComponents(ad, , incTop, 1)
            For Each ky In .Keys
                sd = aiDocAssy(obOf(.Item(ky)))
                pn = pnOfBoltConn(sd)

                If Len(pn) > 0 Then
                    dn = sd.FullDocumentName

                    With rt
                        If .Exists(pn) Then
                            wd = dcOb(.Item(pn))
                        Else
                            wd = New Scripting.Dictionary
                            .Add(pn, wd)
                        End If
                    End With

                    With wd
                        If .Exists(dn) Then
                            If obOf(.Item(dn)) Is sd Then
                                'should be good
                            Else 'not so sure
                                Stop
                            End If
                        Else
                            .Add(dn, sd)
                        End If
                    End With
                End If
            Next
        End With

        dcOfBoltConnIn = rt
        '
        'Debug.Print(txDumpLs(dcOfBoltConnIn(aiDocAssy(aiDocActive())).Keys)
        '
    End Function

    Public Function aiDocContentMember(
    ad As Inventor.PartDocument
) As Inventor.PartDocument
        If ad Is Nothing Then
            aiDocContentMember = ad
        ElseIf ad.ComponentDefinition.IsContentMember Then
            aiDocContentMember = ad
        Else
            aiDocContentMember = Nothing
        End If
    End Function

    Public Function dcIfDesignAccel(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcIfDesignAccel
        '
        ' Accepting a Dictionary of form
        ' generated by d1g3f1, verify
        ' that all Items represent Content
        ' Center components, and return
        ' same Dictionary if so.
        '
        ' If any Items are NOT Content Center
        ' components, return Nothing
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object

        rt = dc

        With dc
            For Each ky In .Keys
                ar = .Item(ky)
                Debug.Print("")
                If aiDocContentMember(
                aiDocPart(aiDocument(
                obOf(ar(0))
            ))) Is Nothing Then
                    rt = Nothing
                End If
            Next
        End With

        dcIfDesignAccel = rt
    End Function

    Public Function rsOfBoltConn(
    ad As Inventor.AssemblyDocument
) As ADODB.Record
        '
        ' rsOfBoltConn -- rsOfBoltConn
        '
        ' Return Record of Components
        ' of one supplied Assembly Document,
        ' provided it's a Bolted Connection.
        '
        ' Call rsOfBoltConnRedux against this function's
        ' resulting Record rt to condense it
        ' to definition of a single instance,
        ' with a count of each member indicating
        ' the number of instances.
        '
        ' (was going to call rsOfBoltConnRedux here and
        ' return THAT result, but realized
        ' this function's preprocessed result
        ' might prove useful in itself, so
        ' decided to return it directly
        ' after all)
        '
        Dim rt As ADODB.Record 'Scripting.Dictionary
        Dim pNum As ADODB.Field
        Dim fNam As ADODB.Field
        Dim zPos As ADODB.Field
        Dim xCen As ADODB.Field
        Dim yCen As ADODB.Field

        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.PartDocument
        Dim bc As Scripting.Dictionary
        Dim p0(2) As Double
        Dim p1(2) As Double

        rt = rsForBoltConn() 'New Scripting.Dictionary
        With rt.Fields
            pNum = .Item("pNum")
            fNam = .Item("fNam")
            zPos = .Item("zPos")
            xCen = .Item("xCen")
            yCen = .Item("yCen")
        End With

        If ad Is Nothing Then 'we got nothing to work with
        Else
            bc = Nothing
            If ad.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                With ad.DocumentInterests
                    If .HasInterest(guidDesignAccl) Then
                        bc = d1g3f1(ad)
                    Else
                        bc = dcIfDesignAccel(d1g3f1(ad))
                    End If
                End With

                If bc Is Nothing Then 'do nothing
                Else
                    For Each oc In ad.ComponentDefinition.Occurrences
                        With oc
                            sd = aiDocument(.Definition.Document)
                            With .RangeBox
                                .MinPoint.GetPointData(p0)
                                .MaxPoint.GetPointData(p1)
                            End With
                        End With


                        rt.AddNew
                        pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign)
                        fNam.Value = sd.FullDocumentName

                        zPos.Value = System.Math.Round(p0(2), 3)
                        'Debug.Print(FormatNumber(p0(2), 3), " ";
                        xCen.Value = System.Math.Round((p0(0) + p1(0)) / 2, 3)
                        'Debug.Print(FormatNumber((p0(0) + p1(0)) / 2, 3), " ";
                        yCen.Value = System.Math.Round((p0(1) + p1(1)) / 2, 3)
                        'Debug.Print(FormatNumber((p0(1) + p1(1)) / 2, 3), " ";
                        'Debug.Print

                        Debug.Print("")
                    Next

                    With rt
                        .Filter = ""
                        If .BOF Then
                            .AddNew
                            pNum.Value = "NONE"
                            fNam.Value = "No Hardware, or Not Bolted Connection!"
                            zPos.Value = 0
                            xCen.Value = 0
                            yCen.Value = 0
                        End If
                        .Sort = "zPos, pNum, yCen, xCen"
                    End With
                End If
            End If
        End If

        rsOfBoltConn = rt 'rsOfBoltConnRedux()
        'Debug.Print(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
    End Function

    Public Function dcOfBoltConn(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary 'ADODB.Record 'Scripting.Dictionary
        '
        ' dcOfBoltConn
        '
        ' Alternate implementation of rsOfBoltConn
        ' returning a Dictionary instead of
        ' a Record. However, this loses
        ' the benefit of a Record's Sort
        ' capability, and so is unlikely
        ' to prove as useful.
        '
        Dim dc As Scripting.Dictionary
        Dim k0 As String
        Dim ct As Long

        Dim rt As ADODB.Record 'Scripting.Dictionary
        Dim pNum As ADODB.Field
        Dim fNam As ADODB.Field
        Dim zPos As ADODB.Field
        Dim xCen As ADODB.Field
        Dim yCen As ADODB.Field

        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.PartDocument
        Dim bc As Scripting.Dictionary
        Dim p0(2) As Double
        Dim p1(2) As Double

        dc = New Scripting.Dictionary
        rt = rsForBoltConn()
        With rt.Fields
            pNum = .Item("pNum")
            fNam = .Item("fNam")
            zPos = .Item("zPos")
            xCen = .Item("xCen")
            yCen = .Item("yCen")
        End With

        If ad Is Nothing Then 'we got nothing to work with
        Else
            bc = Nothing
            If ad.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                With ad.DocumentInterests
                    If .HasInterest(guidDesignAccl) Then
                        bc = d1g3f1(ad)
                    Else
                        bc = dcIfDesignAccel(d1g3f1(ad))
                    End If
                End With

                If bc Is Nothing Then 'do nothing
                Else
                    For Each oc In ad.ComponentDefinition.Occurrences
                        With oc
                            sd = aiDocument(.Definition.Document)
                            With .RangeBox
                                .MinPoint.GetPointData(p0)
                                .MaxPoint.GetPointData(p1)
                            End With
                        End With

                        k0 = FormatNumber(p0(2), 3) _
                    & "|" & aiDocPropVal(
                        sd, pnPartNum, gnDesign
                    )
                        With dc
                            If .Exists(k0) Then
                                ct = 1 + .Item(k0)
                                .Item(k0) = ct
                            Else
                                .Add(k0, 1)
                            End If
                        End With

                        rt.AddNew
                        pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign)
                        fNam.Value = sd.FullDocumentName

                        zPos.Value = System.Math.Round(p0(2), 3)
                        'Debug.Print(FormatNumber(p0(2), 3), " ";
                        xCen.Value = System.Math.Round((p0(0) + p1(0)) / 2, 3)
                        'Debug.Print(FormatNumber((p0(0) + p1(0)) / 2, 3), " ";
                        yCen.Value = System.Math.Round((p0(1) + p1(1)) / 2, 3)
                        'Debug.Print(FormatNumber((p0(1) + p1(1)) / 2, 3), " ";
                        'Debug.Print

                        Debug.Print("")
                    Next

                    With rt
                        .Filter = ""
                        If .BOF Then
                            .AddNew
                            pNum.Value = "NONE"
                            fNam.Value = "No Hardware, or Not Bolted Connection!"
                            zPos.Value = 0
                            xCen.Value = 0
                            yCen.Value = 0
                        End If
                        .Sort = "zPos, pNum, yCen, xCen"
                    End With
                End If
            End If
        End If

        dcOfBoltConn = dc 'rt
        'Debug.Print(dcOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
    End Function

    Public Function rsForBoltConn() As ADODB.Record
        '
        ' rsForBoltConn -- rsForBoltConn
        '
        ' Generate an new, empty Record
        ' to gather data on Bolted Connection
        '
        Dim rt As ADODB.Record

        rt = New ADODB.Record
        With rt
            With .Fields
                .Append("zPos", DataTypeEnum.adDouble)
                .Append("pNum", DataTypeEnum.adVarChar, 63)
                .Append("fNam", DataTypeEnum.adVarChar, 255)
                '.Append "", adVarChar, 63
                .Append("xCen", DataTypeEnum.adDouble)
                .Append("yCen", DataTypeEnum.adDouble)
                '.Append "", adDouble
                '.Append "", adDouble
            End With
            .Open()
        End With
        rsForBoltConn = rt
    End Function

    Public Function rsOfBoltConnRedux(
    rs As ADODB.Record
) As ADODB.Record
        '
        ' rsOfBoltConnRedux
        '
        ' Condense supplied Record
        ' of Bolted Connection Assembly
        ' to summary of Components of
        ' ONE instance.
        '
        ' Include count of each member
        ' Component in Assembly, which
        ' should be the same for ALL
        ' Components, and reflect the
        ' total number of instances
        ' in the Assembly.
        '
        ' In most cases, this count
        ' should be just one, given
        ' the way Bolted Connections
        ' are generated and used here.
        ' However, some models might
        ' be found which use patterns
        ' or multiple holes, thus
        ' producing one BC Assembly
        ' defining multiple instances.
        ' A means to address this might
        ' therefore be required in future.
        '
        Dim rt As ADODB.Record
        Dim pNumIn As ADODB.Field
        Dim zPosIn As ADODB.Field
        Dim pNumOut As ADODB.Field
        Dim zPosOut As ADODB.Field
        Dim xCenOut As ADODB.Field

        Dim dc As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ky As Object

        Dim zp As Double
        Dim pn As String

        dc = New Scripting.Dictionary

        With rs
            With .Fields
                pNumIn = .Item("pNum")
                zPosIn = .Item("zPos")
            End With

            .Sort = "zPos"
            If Not .BOF Then
                Do Until .EOF
                    With dc
                        zp = zPosIn.Value
                        If .Exists(zp) Then
                            wk = .Item(zp)
                        Else
                            wk = New Scripting.Dictionary
                            .Add(zp, wk)
                        End If
                    End With

                    With wk
                        pn = pNumIn.Value
                        If .Exists(pn) Then
                            .Item(pn) = 1 + .Item(pn)
                        Else
                            .Add(pn, 1)
                        End If
                    End With
                    .MoveNext
                Loop
            End If
        End With

        rt = rsForBoltConn()
        With rt
            With .Fields
                pNumOut = .Item("pNum")
                zPosOut = .Item("zPos")
                xCenOut = .Item("xCen")
            End With

            With dc
                For Each ky In .Keys
                    wk = .Item(ky)

                    With wk
                        If .Count > 1 Then
                            Stop
                        Else
                            rt.AddNew
                            zPosOut.Value = CDbl(ky)
                            pNumOut.Value = .Keys(0)
                            xCenOut.Value = CDbl(.Items(0))
                        End If
                    End With
                Next
            End With

            .Filter = ""
            .Sort = "zPos, xCen"
        End With

        rsOfBoltConnRedux = rt
        'Debug.Print(rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)).GetString
    End Function

    Public Function rsOfBoltConnRedux02(
    rs As ADODB.Record
) As ADODB.Record
        '
        ' rsOfBoltConnRedux02
        '
        ' Condense supplied Record
        ' of Bolted Connection Assembly
        ' to summary of Components of
        ' ONE instance.
        '
        ' Include count of each member
        ' Component in Assembly, which
        ' should be the same for ALL
        ' Components, and reflect the
        ' total number of instances
        ' in the Assembly.
        '
        ' In most cases, this count
        ' should be just one, given
        ' the way Bolted Connections
        ' are generated and used here.
        ' However, some models might
        ' be found which use patterns
        ' or multiple holes, thus
        ' producing one BC Assembly
        ' defining multiple instances.
        ' A means to address this might
        ' therefore be required in future.
        '
        Dim rt As ADODB.Record
        Dim pNumIn As ADODB.Field
        Dim zPosIn As ADODB.Field
        Dim pNumOut As ADODB.Field
        Dim zPosOut As ADODB.Field
        Dim xCenOut As ADODB.Field

        Dim dc As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ky As Object

        Dim zp As Double
        Dim pn As String

        dc = New Scripting.Dictionary

        With rs
            With .Fields
                pNumIn = .Item("pNum")
                zPosIn = .Item("zPos")
            End With

            .Sort = "zPos"
            If Not .BOF Then
                Do Until .EOF
                    With dc
                        zp = zPosIn.Value
                        If .Exists(zp) Then
                            wk = .Item(zp)
                        Else
                            wk = New Scripting.Dictionary
                            .Add(zp, wk)
                        End If
                    End With

                    With wk
                        pn = pNumIn.Value
                        If .Exists(pn) Then
                            .Item(pn) = 1 + .Item(pn)
                        Else
                            .Add(pn, 1)
                        End If
                    End With
                    .MoveNext
                Loop
            End If
        End With

        rt = rsForBoltConn()
        With rt
            With .Fields
                pNumOut = .Item("pNum")
                zPosOut = .Item("zPos")
                xCenOut = .Item("xCen")
            End With

            With dc
                For Each ky In .Keys
                    wk = .Item(ky)

                    With wk
                        If .Count > 1 Then
                            Stop
                        Else
                            rt.AddNew
                            zPosOut.Value = CDbl(ky)
                            pNumOut.Value = .Keys(0)
                            xCenOut.Value = CDbl(.Items(0))
                        End If
                    End With
                Next
            End With

            .Filter = ""
            .Sort = "zPos, xCen"
        End With

        rsOfBoltConnRedux02 = rt
        'Debug.Print(rsOfBoltConnRedux02(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)).GetString
    End Function

    Public Function bcPtNumFromRS(
    rs As ADODB.Record
) As String
        bcPtNumFromRS = bcPtNumFromRSv2(rs)
    End Function

    Public Function bcPtNumFromRSv1(
    rs As ADODB.Record
) As String
        '
        ' bcPtNumFromRSv1
        '
        ' Generate a uniquely identifying
        ' Part Number from supplied Record
        ' Given a "Bolted Connection",
        '
        Dim pNumIn As ADODB.Field
        Dim xCenIn As ADODB.Field
        Dim rt As String
        Dim pn As String
        Dim ft As Object
        Dim ct As Long

        With rs
            With .Fields
                pNumIn = .Item("pNum")
                xCenIn = .Item("xCen")
            End With

            .Sort = "zPos"
            If .BOF Or .EOF Then
                rt = ""
            Else
                .Sort = "zPos"
                pn = pNumIn.Value
                rt = "BC" & Mid$(
                pn, 3, Len(pn) - 4
            ) & Right$(pn, 2)
                ct = xCenIn.Value

                For Each ft In {
                "zPos <= 0",
                "zPos > 0"
            }
                    rt = rt & "-"
                    .Filter = ft
                    If Not .BOF Then
                        .Sort = "zPos"
                        Do Until .EOF
                            If ct <> xCenIn.Value Then
                                Stop
                            End If
                            pn = pNumIn.Value
                            rt = rt & Left$(pn, 2)
                            .MoveNext
                        Loop
                    End If
                Next

                If ct > 1 Then
                    rt = rt & Format$(ct, "-X00")
                End If
            End If
        End With

        bcPtNumFromRSv1 = rt
        'Debug.Print(bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)))
    End Function

    Public Function bcPtNumFromRSv2(
    rs As ADODB.Record
) As String
        '
        ' bcPtNumFromRSv2
        '
        ' Generate a uniquely identifying
        ' Part Number from supplied Record
        ' Given a "Bolted Connection",
        '
        Dim pNumIn As ADODB.Field
        Dim xCenIn As ADODB.Field
        Dim rt As String
        Dim pn As String
        Dim ft As Object
        Dim ct As Long

        With rs
            With .Fields
                pNumIn = .Item("pNum")
                xCenIn = .Item("xCen")
            End With

            '.Sort = "zPos"
            .Filter = ""
            If .BOF Or .EOF Then
                rt = ""
            Else
                .Sort = "zPos"
                pn = pNumIn.Value
                rt = "BC" & Right$(pn, 1) & Mid$(
                pn, 3, Len(pn) - 4
            ) '& Right$(pn, 2)
                ct = xCenIn.Value

                For Each ft In {
                "zPos <= 0|zPos",
                "zPos > 0|zPos desc"
            }
                    rt = rt & "-"
                    .Filter = Left$(ft, InStr(ft, "|") - 1)
                    If Not .BOF Then
                        .Sort = Mid$(ft, InStr(ft, "|") + 1) '"zPos"
                        rt = rt & Left$(pNumIn.Value, 2)
                        .MoveNext
                        Do Until .EOF
                            If ct <> xCenIn.Value Then
                                Stop
                            End If
                            'pn = pNumIn.Text
                            'rt = rt & Left$(pn, 2)
                            rt = rt & Left$(pNumIn.Value, 1)
                            .MoveNext
                        Loop
                    End If
                Next

                If ct > 1 Then
                    rt = rt & Format$(ct, "-X00")
                End If
            End If
        End With

        If Len(rt) > 23 Then
            Stop
        End If
        bcPtNumFromRSv2 = rt
        'Debug.Print(bcPtNumFromRSv2(rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)))
    End Function

    Public Function pnOfBoltConn(
    ad As Inventor.AssemblyDocument
) As String
        pnOfBoltConn = bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(ad)))
        'Debug.Print(pnOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)
    End Function

    Public Function dcOfBoltConn02(
    ad As Inventor.AssemblyDocument
) As Scripting.Dictionary
        '
        ' dcOfBoltConn02
        '
        ' Second variation on dcOfBoltConn
        ' returning a Dictionary of Component
        ' quantities, keyed on Item Number.
        '
        Dim rt As Scripting.Dictionary
        Dim pNum As String
        Dim ct As Long

        Dim oc As Inventor.ComponentOccurrence
        Dim sd As Inventor.PartDocument
        Dim bc As Scripting.Dictionary

        rt = New Scripting.Dictionary

        If ad Is Nothing Then 'we got nothing to work with
        Else
            bc = Nothing
            If ad.ComponentDefinition.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                With ad.DocumentInterests
                    If .HasInterest(guidDesignAccl) Then
                        bc = d1g3f1(ad)
                    Else
                        bc = dcIfDesignAccel(d1g3f1(ad))
                    End If
                End With

                If bc Is Nothing Then 'do nothing
                Else
                    For Each oc In ad.ComponentDefinition.Occurrences
                        sd = aiDocument(oc.Definition.Document)

                        pNum = aiDocPropVal(sd,
                        pnPartNum, gnDesign
                    )
                        With rt
                            If .Exists(pNum) Then
                                ct = 1 + .Item(pNum)
                                .Item(pNum) = ct
                            Else
                                .Add(pNum, 1)
                            End If
                        End With

                        Debug.Print("")
                    Next
                End If
            End If
        End If

        dcOfBoltConn02 = rt
        'Debug.Print(dcOfBoltConn02(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
    End Function

    Public Function rsFiltered(
    rs As ADODB.Record,
    Optional flText As String = ""
) As ADODB.Record
        rs.Filter = flText
        rsFiltered = rs
    End Function

    Public Function rsFromGnsSql(
    sqlText As String
) As ADODB.Record
        '
        '
        '
        Dim rt As ADODB.Record

        With CnGnsDoyle()
            rt = .Execute(sqlText)
            If rt Is Nothing Then
                Stop
            End If
            rsFromGnsSql = rt
        End With
    End Function

    Public Function rsAiPurch01fromDict(
    dc As Scripting.Dictionary
) As ADODB.Record
        '
        '
        '
        rsAiPurch01fromDict = rsFromGnsSql(
        sqlSelAiPurch01fromDict(dc)
    )
    End Function

    Public Function rsAiPurch01fromAssy(
    AiDoc As Inventor.Document
) As ADODB.Record
        '
        '
        '
        rsAiPurch01fromAssy _
    = rsFromGnsSql(
        sqlSelAiPurch01fromAssy(AiDoc)
    )
    End Function

    Public Function rsAiPdParts01fromAssy(
    AiDoc As Inventor.Document
) As ADODB.Record
        '
        '
        '
        rsAiPdParts01fromAssy _
    = rsFromGnsSql(
        sqlSelAiPdParts01fromAssy(AiDoc)
    )
        'Debug.Print(rsAiPdParts01fromAssy(aiDocActive()).GetString(adClipString)
    End Function

    Public Function dcAiPurch01fromAdoRs(
    rs As ADODB.Record
) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim fdItem As ADODB.Field
        Dim fdType As ADODB.Field
        Dim fdFmly As ADODB.Field

        rt = New Scripting.Dictionary
        With rs
            If Not .BOF Then
                .Filter = ""

                With .Fields
                    fdItem = .Item("Item")
                    fdType = .Item("Type")
                    fdFmly = .Item("Family")
                End With

                Do Until .EOF
                    rt.Add(fdItem.Value, {
                    fdType.Value, fdFmly.Value
                })
                    .MoveNext
                Loop

                .Close
            End If
            dcAiPurch01fromAdoRs = rt
        End With
    End Function

    Public Function dcAiPurch01fromDict(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        '
        '
        dcAiPurch01fromDict _
    = dcAiPurch01fromAdoRs(
        rsAiPurch01fromDict(dc)
    )
    End Function

    Public Function dcAiPurch01fromAssy(
    AiDoc As Inventor.Document
) As Scripting.Dictionary
        '
        '
        '
        dcAiPurch01fromAssy _
    = dcAiPurch01fromAdoRs(
        rsAiPurch01fromAssy(AiDoc)
    )
    End Function

End Module