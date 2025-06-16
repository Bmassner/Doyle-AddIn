Class DcPopulator
    Private dc As Scripting.Dictionary

    Private Function CheckDC(
    Optional dcIn As Scripting.Dictionary = Nothing,
    Optional Opts As Long = 0
) As Scripting.Dictionary
        ''
        ''
        If dcIn Is Nothing Then
            If dc Is Nothing Then
                dc = New Scripting.Dictionary
            End If
        Else
            If dc Is Nothing Then
                dc = dcIn
            Else
                Stop
                If Opts And 1 Then
                    'Planning Key Value Replacement here
                End If
            End If
        End If

        CheckDC = dc
        ''
        ''
    End Function

    Public Function UseDictionary(
    Dict As Scripting.Dictionary,
    Optional Opts As Long = 0
) As DcPopulator
        ''
        ''
        With CheckDC(Dict, Opts)
        End With

        UseDictionary = Me
    End Function

    Public Function Setting(
    Key As Object, Item As Object
) As DcPopulator
        With CheckDC()
            If .Exists(Key) Then .Remove(Key)
            .Add(Key, Item)
        End With

        Setting = Me
    End Function

    Public Function Count() As Long
        Count = CheckDC().Count
    End Function

    Public Function Dictionary(
    Optional Dict As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        Dictionary = CheckDC(Dict)
    End Function

    Public Function Matching(
    KeySet As Object
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        If Not IsNothing(KeySet) Then
            rt = New Scripting.Dictionary

            With CheckDC()
                For Each ky In KeySet
                    If .Exists(ky) Then
                        If rt.Exists(ky) Then
                            'don't need to match it twice!
                        Else
                            rt.Add(ky, .Item(ky))
                        End If
                    End If
                Next : End With
        Else
            rt = Matching({KeySet})
        End If

        Matching = rt
    End Function

    Public Function Exists(Key As Object) As Boolean
        With Dictionary()
            Exists = .Exists(Key)
        End With
    End Function

    Public Function Item(Key As Object) As Object
        With Dictionary()
            If .Exists(Key) Then
                Item = .Item(Key)
            Else
                Item = Nothing
            End If
        End With
    End Function

    '
    ' OPTIONAL section for Inventor.NameValueMap
    ' the following functions will ONLY compile
    ' within the Autodesk Inventor VBA environment,
    ' or other environment which supports the same
    ' NameValueMap classes and structures.
    '
    ' It should be disabled or deleted for use
    ' outside of any such environment.
    '
    '    Public Function UsingNameValMap(
    '    NVMap As Inventor.NameValueMap,
    '    Optional Opts As Long = 0
    ') As dcPopulator
    '        ''
    '        ''
    '        UsingNameValMap = UseDict( _
    '    dcFromAiNameValMap(NVMap), Opts)
    'End Function

    '    Public Function NameValMap(
    '    Optional NVMap As Inventor.NameValueMap = Nothing
    ') As Inventor.NameValueMap
    '        NameValMap = dc2aiNameValMap(Dictionary(), NVMap)
    '    End Function
    '
    ' END of OPTIONAL Inventor.NameValueMap section
    '
    '

End Class