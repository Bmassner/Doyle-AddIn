Module libDcGeneral

    Public Function DcCopy(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcCopy -- return a new Dictionary
        '     copying the contents of the
        '     one supplied, including COPIES
        '     of any Dictionary Objects within.
        '
        Dim rt As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim bx As Object
        Dim ky As Object

        rt = New Scripting.Dictionary

        If dc Is Nothing Then 'skip transfer
            'cuz there ain't Nothing in Nothing!
        Else 'we MIGHT have stuff to move
            With dc : For Each ky In .Keys
                    bx = { .Item(ky)}
                    ck = dcOb(obOf(bx(0)))

                    If ck Is Nothing Then
                        rt.Add(ky, bx(0))
                    Else
                        rt.Add(ky, DcCopy(ck))
                    End If
                Next : End With
        End If

        DcCopy = rt
    End Function

    Public Function DcWith(ky As Object, it As Object,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        If dc Is Nothing Then
            DcWith = DcWith(ky, it,
        New Scripting.Dictionary)
        Else
            With dc
                If .Exists(ky) Then .Remove(ky)
                .Add(ky, it)
            End With
            DcWith = dc
        End If
    End Function

    Public Function NuDcPopulator(
    Optional Dict As Scripting.Dictionary = Nothing,
    Optional Opts As Long = 0
) As DcPopulator
        With New DcPopulator
            NuDcPopulator = .UseDictionary(Dict, Opts)
        End With
        'Debug.Print(dumpLsKeyVal(nuDcPopulator().Setting("A", "B").Setting("C", "D").Dictionary)
    End Function

    Public Function DcItemIfPresent(
    dc As Scripting.Dictionary, ky As Object,
    Optional vtMissing As VbVarType = 0
) As Object
        If dc Is Nothing Then
            DcItemIfPresent = noVal(vtMissing)
        Else
            With dc
                If .Exists(ky) Then
                    If TypeOf .Item(ky) Is Object Then
                        DcItemIfPresent = .Item(ky)
                    Else
                        DcItemIfPresent = .Item(ky)
                    End If
                Else
                    If vtMissing = vbObject Then
                        DcItemIfPresent = noVal(vtMissing)
                    Else
                        DcItemIfPresent = noVal(vtMissing)
                    End If
                End If
            End With
        End If
    End Function

    Public Function DcInDc(dcKey As String,
    inDc As Scripting.Dictionary
) As Scripting.Dictionary
        DcInDc = dcOb(obOf(
        DcItemIfPresent(
        inDc, dcKey, vbObject
    )))
    End Function

    Public Function DcInDcMk(
    ky As Object, dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcInDcMk --
        '
        Dim rt As Scripting.Dictionary

        With dc
            If .Exists(ky) Then
                rt = dcOb(.Item(ky))
            Else
                rt = New Scripting.Dictionary
                .Add(ky, rt)
            End If
        End With

        DcInDcMk = rt
    End Function

    Public Function DcOfKeys2match(
    ls As Object
) As Scripting.Dictionary
        '
        ' dcOfKeys2match -- generate a Dictionary
        '     mapping a supplied Key, or Array
        '     of Keys to itself or themselves
        '     '
        '     primary purpose is to provide
        '     a 'reference' Dictionary of Keys
        '     to be sought in other Dictionaries
        '     '
        '     (formerly d4g4f2)
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary

        If TypeOf ls Is Array Then
            For Each ky In ls
                rt.Add(ky, ky)
            Next
        Else
            rt.Add(ls, ls)
        End If

        DcOfKeys2match = rt
    End Function

    Public Function DcKeysInCommon(
    d0 As Scripting.Dictionary,
    d1 As Scripting.Dictionary,
    Optional pk As Long = 0
) As Scripting.Dictionary
        '
        ' dcKeysInCommon -- return intersection
        '     of two Dictionary Objects based on
        '     matching keys. Use optional pk value
        '     to select which Dictionary's Items
        '     to return in result:
        '
        '     0 returns an array of Items from both
        '       this is the default
        '     1 returns only Items from the first
        '     2 returns only Items from the second
        '
        ' NOTE that if either Dictionary Object
        '     is not supplied (is Nothing), then
        '     an empty Dictionary is returned,
        '     just as if an empty Dictionary
        '     had been supplied.
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ls As Object

        If d0 Is Nothing Then
            rt = DcKeysInCommon(
            New Scripting.Dictionary, d1
        )
        ElseIf d1 Is Nothing Then
            rt = DcKeysInCommon(
            d0, New Scripting.Dictionary
        )
        Else
            rt = New Scripting.Dictionary
            With d0 : For Each ky In .Keys
                    If d1.Exists(ky) Then
                        ls = {
                .Item(ky),
                d1.Item(ky)
            }
                        rt.Add(ky, {ls, ls(0), ls(1)}(pk))
                    End If : Next : End With
        End If
        DcKeysInCommon = rt
    End Function

    Public Function DcKeysMissing(
    dcWith As Scripting.Dictionary,
    dcWout As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcKeysMissing -- return difference
        ' of first Dictionary Object minus
        ' those keys found in the second.
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dcWith
            For Each ky In .Keys
                If dcWout.Exists(ky) Then
                    'don't include this one
                Else
                    rt.Add(ky, .Item(ky))
                End If
            Next
        End With
        DcKeysMissing = rt
    End Function

    Public Function DcKeysCombined(
    d0 As Scripting.Dictionary,
    d1 As Scripting.Dictionary,
    Optional pk As Long = 0
) As Scripting.Dictionary
        '
        ' dcKeysCombined -- return union
        ' of two Dictionary Objects. For
        ' keys in both, use optional pk
        ' value to select which Dictionary's
        ' Items to return in result:
        '
        '     0 returns an array of Items from both
        '       this is the default
        '     1 returns only Items from the first
        '     2 returns only Items from the second
        '
        Dim rt As Scripting.Dictionary
        Dim ob As Scripting.Dictionary
        Dim ky As Object
        Dim ls As Object

        If pk > 1 Then
            rt = DcKeysCombined(d1, d0, 1)
        Else
            rt = DcKeysInCommon(d0, d1, pk)

            For Each ls In {d0, d1}
                ob = ls
                With DcKeysMissing(ob, rt) 'd0
                    For Each ky In .Keys
                        rt.Add(ky, .Item(ky))
                    Next : End With : Next

            'With dcKeysMissing(d1, rt)
            'For Each ky In .Keys
            '    rt.Add(ky, .Item(ky)
            'Next: End With
        End If

        ' rt = dcKeysInCommon()

        If d0 Is Nothing Then
            Stop
        ElseIf d1 Is Nothing Then
            Stop
        Else
            ' NOTE[2023.04.10.1449]
            ' need to review what's going on here
            ' this LOOKS like an effor to include items
            ' left out of earlier operation, however,
            ' it's not clear this stage isn't redundant.
            With d0 : For Each ky In .Keys
                    If d1.Exists(ky) Then
                        ls = {
                .Item(ky), d1.Item(ky)
            }
                        If rt.Exists(ky) Then
                            '          If 'ConvertToJson({ls, ls(0), ls(1)}(pk)) <> ConvertToJson(rt.Item(ky)) Then
                            Stop
                            ' End If'
                        Else
                            rt.Add(ky, {ls, ls(0), ls(1)}(pk))
                        End If
                    End If
                Next : End With
        End If
        DcKeysCombined = rt
    End Function

    Public Function DcOfIdent(
    src As Object
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary

        If TypeOf src Is Array Then
            For Each ky In src
                rt.Add(ky, ky)
            Next
        Else
            rt.Add(src, src)
        End If

        DcOfIdent = rt
    End Function

    Public Function DcTransposed(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' Transpose Key values of supplied
        ' Dictionary with matching Item values.
        '
        ' As written, will ONLY work against
        ' a Dictionary whose Item values,
        ' like its Keys, are unique.
        '
        Dim rt As Scripting.Dictionary
        Dim Fm As Object

        rt = New Scripting.Dictionary
        With dc
            For Each Fm In .Keys
                If rt.Exists(.Item(fm)) Then
                    Stop
                Else
                    rt.Add(.Item(Fm), Fm)
                End If
            Next
        End With
        DcTransposed = rt
        'Debug.Print(dumpLsKeyVal(dcTransposed(dcOfRgAddresses(dcByFormulaOnly(dcOfRgByFormulaR1C1(wsNamed(chosenWorkbook(), "StandardItems"))))), "|")
    End Function

    Public Function DcTransGrouped(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcTransGrouped
        '     (derived from dcTransposed)
        '
        ' generate new Dictionary "tramsposing"
        ' Key values with matching Item values
        ' in supplied Dictionary.
        '
        ' because more than one Key might map to
        ' the same Item, the returned Dictionary
        ' maps each (Item) Key to a Dictionary of
        ' the Keys which originally mapped to it.
        '
        ' each of these Dictionaries in turn
        ' maps the original Key back to its
        ' corresponding Item once more,
        ' since, HEY, it might as well!
        '
        ' obviously, an effor to work around
        ' the main limitation of the original
        ' dcTransposed, which can only work
        ' against a Dictionary whose Items,
        ' like its Keys, are unique.
        '
        Dim rt As Scripting.Dictionary
        Dim it As Scripting.Dictionary
        Dim ky As Object
        Dim ar As Object

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                ar = { .Item(ky)}
                If Not rt.Exists(ar(0)) Then
                    rt.Add(ar(0), New Scripting.Dictionary)
                End If
                dcOb(rt.Item(ar(0))).Add(ky, ar(0))
            Next : End With
        DcTransGrouped = rt
    End Function

    Public Function DcDepth(
    dc As Scripting.Dictionary
) As Long
        '
        ' this function extracts the "depth"
        ' of the supplied Dictionary object,
        ' that is, how many "levels" of
        ' Dictionary objects it contains,
        ' counting the supplied Dictionary
        ' itself. When an actual Dictionary
        ' is supplied, the value returned
        ' will be at least 1. It will only
        ' be zero when Nothing is supplied.
        '
        Dim rt As Long
        Dim ck As Long
        Dim ky As Object

        If dc Is Nothing Then
            DcDepth = 0
        Else
            rt = 0

            With dc : For Each ky In .Keys
                    ck = DcDepth(dcOb(obOf(.Item(ky))))
                    If ck > rt Then rt = ck
                Next : End With

            DcDepth = 1 + rt
        End If
    End Function

    Public Function DcFlattenDown(
    Dict As Scripting.Dictionary,
    Optional DownTo As Long = 1
) As Scripting.Dictionary
        '
        ' this function partially "flattens"
        ' a hierarchy of Dictionary objects
        ' (a Dictionary of Dictionaries,
        '  potentially of more Dictionaries)
        ' starting from the top, working
        ' down to 'DownTo' levels
        '
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim sk As Object

        If Dict Is Nothing Then
            DcFlattenDown = Nothing
        Else
            rt = New Scripting.Dictionary

            With Dict
                For Each ky In .Keys
                    it = { .Item(ky)}

                    If DownTo > 0 Then
                        sd = DcFlattenDown(dcOb(obOf(it(0))), DownTo - 1)
                    Else
                        sd = Nothing
                    End If

                    If sd Is Nothing Then
                        rt.Add(ky, it(0))
                    Else
                        With sd
                            For Each sk In .Keys
                                rt.Add(ky & "." & sk, .Item(sk))
                            Next
                        End With
                    End If
                Next
            End With

            DcFlattenDown = rt
        End If
    End Function

    Public Function DcFlattenUp(
    Dict As Scripting.Dictionary,
    Optional DownFrom As Long = 0
) As Scripting.Dictionary
        '
        ' this function partially "flattens"
        ' a hierarchy of Dictionary objects
        ' (a Dictionary of Dictionaries,
        '  potentially of more Dictionaries)
        ' starting BELOW the top, skipping
        ' DownFrom levels before "flattening"
        ' Dictionaries below and at that level
        '
        Dim rt As Scripting.Dictionary
        Dim sd As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim sk As Object

        If Dict Is Nothing Then
            DcFlattenUp = Nothing
        Else
            rt = New Scripting.Dictionary

            With Dict
                For Each ky In .Keys
                    it = { .Item(ky)}
                    sd = dcOb(obOf(it(0))) 'dcFlattenUp(, DownFrom - 1)

                    If sd Is Nothing Then
                        rt.Add(ky, it(0))
                    Else
                        If DownFrom > 0 Then
                            rt.Add(ky, DcFlattenUp(sd, DownFrom - 1))
                        Else
                            With DcFlattenUp(sd, 0)
                                For Each sk In .Keys
                                    rt.Add(ky & "." & sk, .Item(sk))
                                Next
                            End With
                        End If
                    End If
                Next
            End With

            DcFlattenUp = rt
        End If
    End Function

    Public Function DcOfDcRekeyedSecToPri(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcOfDcRekeyedSecToPri - take Dictionary of Dictionaries
        '     and return Dictionary of Dictionaries
        '     with Secondary Keys promoted to Primary,
        '     and Primary demoted to Secondary
        '
        Dim rt As Scripting.Dictionary
        Dim sb As Scripting.Dictionary
        Dim kp As Object
        Dim ks As Object
        Dim ar As Object

        rt = New Scripting.Dictionary

        With dc : For Each kp In .Keys
                sb = dcOb(.Item(kp))
                If sb Is Nothing Then 'we got a problem
                    'Stop
                    Debug.Print("")  'Breakpoint Landing
                Else
                    With sb : For Each ks In .Keys
                            ar = { .Item(ks)}
                            With rt
                                If Not .Exists(ks) Then
                                    .Add(ks, New Scripting.Dictionary)
                                End If

                                With dcOb(.Item(ks))
                                    If .Exists(kp) Then 'another problem
                                        Stop
                                    Else
                                        .Add(kp, ar(0))
                                    End If : End With
                            End With
                        Next : End With : End If
            Next : End With

        DcOfDcRekeyedSecToPri = rt
    End Function

    Public Function DcCmpTextOf2items(
    id0 As String, id1 As String,
    it0 As Object, it1 As Object
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim dc0 As Scripting.Dictionary
        Dim dc1 As Scripting.Dictionary
        Dim ob0 As Object
        Dim ob1 As Object
        Dim tx0 As String
        Dim tx1 As String
        Dim ck As Long
        Dim c2 As Long

        ck = IIf(it0 Is Nothing, 0, 1) _
   + IIf(it1 Is Nothing, 0, 2) _
   + IIf(it0 Is Nothing, 4, 0) _
   + IIf(it1 Is Nothing, 8, 0)
        If ck And 1 Then ck = ck + IIf(it0 Is Nothing, 4, 0)
        If ck And 2 Then ck = ck + IIf(it1 Is Nothing, 8, 0)
        ' UPDATE[2021.12.09]:
        '     Added trap for Nothing Objects
        '     to treat them as Empty as well

        c2 = ck And 12 '$11XX
        If c2 Then 'at least one side
            'is Empty/Nothing
            With NuDcPopulator()
                If c2 = 12 Then 'both
                    'are equally Empty
                    rt = .Setting(
                    "==", ""
                ).Dictionary()
                ElseIf c2 = 8 Then 'it's just it1
                    'and id0 needs processed
                    'If ck And 1 Then
                    '     rt = dcCmpTextOf2dc( _
                    '        dcOb(it0), .Dictionary() _
                    '    )
                    'Else
                    rt = .Setting(
                        id0, it0
                    ).Dictionary()
                    'End If
                Else 'it0 is the empty one
                    'and id1 needs processed
                    'If ck And 2 Then
                    '     rt = dcCmpTextOf2dc( _
                    '        .Dictionary(), dcOb(it1) _
                    '    )
                    'Else
                    rt = .Setting(
                        id1, it1
                    ).Dictionary()
                    'End If
                End If
            End With
        Else 'both sides have data
            If ck = 3 Then 'PROBABLY a couple
                'of comparable Dictionaries.
                'compare these recursively
                rt = DcCmpTextOf2dc(
                dcOb(it0), dcOb(it1)
            )
            ElseIf ck = 0 Then 'NO Dictionaries
                'just a couple of String
                'or (hopefully) String-
                'compatible values.
                'compare them directly.
                tx0 = CStr(it0)
                tx1 = CStr(it1)

                ' rt = New Scripting.Dictionary
                'With rt
                With NuDcPopulator()
                    If tx0 = tx1 Then 'match found
                        rt = .Setting(
                        "==", tx0
                    ).Dictionary()
                        '.Add("==", tx0
                    Else 'mismatched
                        rt = .Setting(
                        id0, tx0
                    ).Setting(
                        id1, tx1
                    ).Dictionary()
                        '.Add(id0, tx0
                        '.Add(id1, tx1
                    End If
                End With
            Else 'we've got a Dictionary AND a String!
                'can't compare them in any way
                'just add each on its own side
                With NuDcPopulator()
                    rt = .Setting(
                    id0, it0
                ).Setting(
                    id1, it1
                ).Dictionary()
                End With
                'ElseIf (ck And 14) = 6 Then '%011X
                '    'covers cases 6 and 7
                '    'excludes 4 and 5
                '    'along with 8~F
                '    'Stop
                '     rt = dcCmpTextOf2dc( _
                '        New Scripting.Dictionary, dcOb(it1) _
                '    )
                'ElseIf (ck And 13) = 9 Then '%10X1
                '    'covers cases 9 and B
                '    'excludes 8 and A
                '    'along with 0~7
                '     rt = dcCmpTextOf2dc( _
                '        dcOb(it0), New Scripting.Dictionary _
                '    )
                '    'ElseIf ck And 4 Then 'it0 is missing
                '    '    If ck And 2 Then 'it1 is a Dictionary
                '    '    Else
                '    '        .Add(id1, CStr(it1)
                '    '    End If
                '    'ElseIf ck = 8 Then
                '    '    .Add(id0, CStr(it0)
                '    Else 'either one or both members are Empty,
                '        ' Nothing, or of incompatible form.
                '        ' They cannot be compared directly,
                '        ' but must, if present, be separately
                '        ' included as they are.
                '
                '        ' NOTE
                '        'Debug.Print("First Item ";
                '        If (ck And 4) = 0 Then 'it0 is present
                '            'Debug.Print("is " & TypeName(it0)
                '            If ck And 1 Then
                '                .Add(id0, it0
                '            Else
                '                .Add(id0, CStr(it0)
                '            End If
                '            Debug.Print("") 'Breakpoint Landing
                '        Else
                '            'Debug.Print("NOT present!"
                '            Debug.Print("") 'Breakpoint Landing
                '        End If
                '        'Stop
                '
                '        'Debug.Print("Second Item ";
                '        If (ck And 8) = 0 Then 'it1 is present
                '            'Debug.Print("is " & TypeName(it1)
                '            If ck And 2 Then
                '                .Add(id1, it1
                '            Else
                '                .Add(id1, CStr(it1)
                '            End If
                '            Debug.Print("") 'Breakpoint Landing
                '        Else
                '            'Debug.Print("NOT present!"
                '            Debug.Print("") 'Breakpoint Landing
                '        End If
                '        'Stop
                '    End If
            End If
        End If

        DcCmpTextOf2items = rt
    End Function

    Public Function DcCmpTextOf2dc(
    dc0 As Scripting.Dictionary,
    dc1 As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim qi As Scripting.Dictionary
        Dim nm0 As String
        Dim nm1 As String
        Dim tx0 As String
        Dim tx1 As String
        Dim ky As Object
        Dim nm As String

        nm0 = "src0"
        nm1 = "src1"

        If dc0 Is Nothing Then
            rt = DcCmpTextOf2dc(
            New Scripting.Dictionary, dc1
        )
        ElseIf dc1 Is Nothing Then
            rt = DcCmpTextOf2dc(dc0,
            New Scripting.Dictionary
        )
        Else
            rt = New Scripting.Dictionary

            With dc0 'add all from wb0
                'and matching from wb1
                For Each ky In .Keys
                    'tx0 = CStr(.Item(ky))

                    ' qi = New Scripting.Dictionary
                    'rt.Add(ky, qi

                    If dc1.Exists(ky) Then 'check for match
                        rt.Add(ky, DcCmpTextOf2items(
                        nm0, nm1, .Item(ky), dc1.Item(ky)
                    ))
                        'tx1 = CStr(dc1.Item(ky))

                        'If tx0 = tx1 Then 'match found
                        '    qi.Add("==", tx0
                        'Else 'mismatched
                        '    qi.Add(nm0, tx0
                        '    qi.Add(nm1, tx1
                        'End If
                    Else 'no match
                        'qi.Add(nm0, tx0
                        rt.Add(ky, DcCmpTextOf2items(
                        nm0, nm1, .Item(ky), vbEmpty
                    ))
                    End If
                Next
            End With

            With dc1 'add any missed from wb1
                For Each ky In .Keys
                    If rt.Exists(ky) Then 'skip it
                        'picked up first round
                    Else 'missed it before
                        'so add it now

                        'tx1 = CStr(.Item(ky))
                        '
                        ' qi = New Scripting.Dictionary
                        'qi.Add(nm1, tx1
                        'rt.Add(ky, qi
                        rt.Add(ky, DcCmpTextOf2items(
                        nm0, nm1, vbEmpty, .Item(ky)
                    ))
                    End If
                Next
            End With
        End If

        DcCmpTextOf2dc = rt
    End Function

    Public Function DcCmpTextOf2subDc(
    dc As Scripting.Dictionary,
    k0 As String, k1 As String
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        rt = DcCmpTextOf2dc(
        DcInDc(k0, dc),
        DcInDc(k1, dc)
    )
        Debug.Print("")
        DcCmpTextOf2subDc = rt
    End Function

    Public Function DcWBQbyCmpResult(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ck As Scripting.Dictionary
        Dim gp As Scripting.Dictionary
        Dim ky As Object
        Dim gk As Object

        rt = New Scripting.Dictionary

        With dc
            For Each ky In .Keys
                ck = .Item(ky)

                If ck.Count > 1 Then
                    gk = "!="
                ElseIf ck.Count < 1 Then
                    Stop 'because SOMETHING went wrong
                    gk = "**"
                Else
                    gk = ck.Keys(0)
                End If

                With rt
                    If .Exists(gk) Then
                        gp = .Item(gk)
                    Else
                        gp = New Scripting.Dictionary
                        rt.Add(gk, gp)
                    End If

                    gp.Add(ky, ck)
                End With
            Next
        End With

        DcWBQbyCmpResult = rt
    End Function

End Module