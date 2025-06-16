Module dvlAiNameValMap
    '
    ' dvlAiNameValMap -- functions to streamline
    '     translation of data from Dictionary
    '     Objects to Inventor NameValueMap
    '     Objects, and vice-versa.
    '
    ' NOTE: these functions MIGHT be supplanted by
    '     their addition to or implementation in
    '     Class Module dcPopulator
    '
    Dim inventorApp As Inventor.Application
    Public Function dc2aiNameValMap(dc As Scripting.Dictionary,
    Optional mp As Inventor.NameValueMap = Nothing
) As Inventor.NameValueMap
        Dim rt As Inventor.NameValueMap
        Dim ky As Object
        Dim it As Object
        Dim nm As String
        Dim ck As VbMsgBoxResult

        If mp Is Nothing Then
            With InventorApp.TransientObjects
                rt = dc2aiNameValMap(
            dc, .CreateNameValueMap
        )   'NameValueMap cannot
                'be created with New
            End With
        Else
            rt = mp
            With dc : For Each ky In .Keys
                    nm = CStr(ky)
                    it = { .Item(ky)}

                    If it(0).GetType Is GetType(Object) Then
                        ' Object handling not
                        ' implemented as yet.
                        ' A general solution is
                        ' not likely possible.

                        ' UPDATE[2022.07.05.1319]
                        ' it appears that a NameValueMap
                        ' CAN include another NameValueMap
                        ' as a Value, thus enabling multi-
                        ' level NameValueMaps. Whether
                        ' other Objects can be so contained
                        ' is likely a question best left
                        ' unexplored for now, but it seems
                        ' at least sub-Dictionaries and
                        ' NameValueMaps can be processed.
                        If TypeOf it(0) Is Scripting.Dictionary Then
                            rt.Add(nm, dc2aiNameValMap(obOf(it(0))))
                        ElseIf TypeOf it(0) Is Inventor.NameValueMap Then
                            rt.Add(nm, it(0))
                        Else
                        End If
                    Else
                        On Error Resume Next

                        Err.Clear()
                        rt.Add(nm, it(0))

                        If Err.Number Then
                            ck = MsgBox(Join({
                        "Key """ & nm,
                        """ Value (" & CStr(it(0)) & ")",
                        "could not be .",
                        "The Key will not",
                        "be assigned.",
                        "",
                        "Click OK to continue",
                        "(Cancel to debug)"
                    }, vbCrLf),
                        vbOKCancel + vbExclamation,
                        "Assignment Error!"
                    )
                            If ck = vbCancel Then
                                Stop 'to Debug
                            End If
                        End If

                        On Error GoTo 0
                    End If
                Next : End With
        End If

        dc2aiNameValMap = rt
    End Function

    Public Function dcFromAiNameValMap(
    mp As Inventor.NameValueMap,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim nm As String
        Dim mx As Long
        Dim dx As Long

        If dc Is Nothing Then
            rt = dcFromAiNameValMap(
            mp, New Scripting.Dictionary
        )
        Else
            rt = dc 'mp
            With mp 'dc
                mx = .Count
                For dx = 1 To mx
                    nm = .Name(dx)

                    rt.Add(nm, itemForDcOr(.Item(nm)))
                Next
            End With
        End If

        dcFromAiNameValMap = rt
    End Function

    Public Function itemForDcOr(it As Object,
    Optional tp As Long = 0
) As Object
        '
        ' itemForDcOr -- given item it, return
        '     transformation according to type,
        '     and type of result desired for
        '     Dictionary and NameValueMap Objects,
        '     according to value of tp:
        '         NameValueMap for tp = 1
        '         Dictionary for any other
        '         value, including default 0
        '     all other types of item are returned
        '     as is, including Objects other than
        '     Dictionary and NameValueMap
        '
        Dim rt As Object
        Dim ck As Object
        Dim mx As Long
        Dim dx As Long
        Dim dc As Scripting.Dictionary

        If TypeOf it Is Array Then
            mx = UBound(it)
            If mx < LBound(it) Then
                rt = {}
            Else
                ReDim rt(mx)

                For dx = 0 To mx
                    ck = {itemForDcOr(it(dx), tp)}
                    If TypeOf ck(0) Is Object Then
                        rt(dx) = ck(0)
                    Else
                        rt(dx) = ck(0)
                    End If
                Next
            End If
        ElseIf TypeOf it Is Inventor.NameValueMap Then
            rt = dcFromAiNameValMap(obOf(it))
        ElseIf TypeOf it Is Scripting.Dictionary Then
            dc = it
            With New DcPopulator
                For Each ck In dc.Keys
                    .Setting(ck, itemForDcOr(dc.Item(ck), tp))
                Next

                If tp = 1 Then 'NameValMap wanted
                    rt = .Count()
                Else 'assume Dictionary
                    rt = .Dictionary()
                End If
            End With
            rt = it
            If TypeName(rt) = "Object" Then
                itemForDcOr = rt
            Else
                itemForDcOr = rt
            End If

            If TypeName(rt) = "Object" Then
                itemForDcOr = rt
            Else
                itemForDcOr = rt
            End If
        End If
    End Function

    Public Function nuAiNameValMap(
    Optional message As Inventor.Document = Nothing
) As Inventor.NameValueMap
        If message Is Nothing Then
            With InventorApp.TransientObjects
                nuAiNameValMap = .CreateNameValueMap
            End With
        Else
            nuAiNameValMap = dc2aiNameValMap(message)
        End If
    End Function
End Module