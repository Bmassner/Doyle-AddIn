Imports Collection = VBA.Collection

Module AssyComponentsCollector
    '
    ' Purpose of this module is to provide an alternate
    ' method of collecting assembly components
    ' using the native VBA Collection instead of
    ' the Scripting Runtime's Dictionary.
    ' Though less powerful/convenient, it does avoid
    ' the need for a reference to the Scripting Runtime.
    '

    Public Function CollectItem(Item As Object,
        Optional Key As Object = Nothing,
        Optional coll As Collection = Nothing
    ) As Collection
        Dim rt As Collection

        If coll Is Nothing Then
            rt = New Collection
        Else
            rt = coll
        End If

        On Error Resume Next
        With Err()
            rt.Add(Item, Key)
            If .Number Then
                If .Number = 457 Then
                    If TypeOf Item Is Object Then
                        If TypeOf rt.Item(Key) Is Object Then
                            If Item Is rt.Item(Key) Then
                                ' OK! Same Object!
                            Else
                                Stop 'Different Objects!
                            End If
                        Else
                            Stop 'Object vs non-Object
                        End If
                    ElseIf TypeOf rt.Item(Key) Is Object Then
                        Stop 'Object vs non-Object
                    Else
                        If Item = rt.Item(Key) Then
                            ' OK! Equal Values!
                        Else
                            Stop 'Different Values!
                        End If
                    End If
                Else
                    Stop
                End If
            End If
        End With
        On Error GoTo 0

        CollectItem = rt
    End Function

    Public Function CollectComponents(
    AiDoc As Inventor.Document,
    Optional coll As Collection = Nothing
) As Collection
        Dim aiDType As Inventor.DocumentTypeEnum
        Dim aiOcc As Inventor.ComponentOccurrence
        Dim rt As Collection

        If coll Is Nothing Then
            rt = CollectComponents(AiDoc, New Collection)
        Else
            rt = coll
            aiDType = AiDoc.DocumentType
            If aiDType = DocumentTypeEnum.kAssemblyDocumentObject Then
                With aiDocAssy(AiDoc).ComponentDefinition
                    For Each aiOcc In .Occurrences
                        If aiOcc.Definition.Document Is AiDoc Then 'skip it
                            'Otherwise, we'll dive into
                            'a bottomless pit of recursion.
                        Else
                            rt = CollectComponents(
                            aiOcc.Definition.Document, rt
                        )
                        End If
                    Next
                End With
            ElseIf aiDType = DocumentTypeEnum.kPartDocumentObject Then
                rt = CollectItem(AiDoc, AiDoc.FullFileName, rt)
            Else
                Stop 'cuz we dunno what to do with this one.
            End If
        End If

        CollectComponents = rt
    End Function

    Public Function ActiveDocsComponents(aiApp As Inventor.Application) As Collection
        ActiveDocsComponents = CollectComponents(aiApp.ActiveDocument)
    End Function

    Public Function strActiveDocsComponents(aiApp As Inventor.Application) As String
        Dim AiDoc As Inventor.Document
        Dim rt As String

        rt = ""
        For Each AiDoc In ActiveDocsComponents(aiApp)
            rt = rt & vbCrLf & AiDoc.FullFileName
        Next
        strActiveDocsComponents = rt
    End Function
    'Debug.Print(strActiveDocsComponents(ThisApplication)

End Module