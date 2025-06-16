Class wkgCls0
    Dim inventorApp As Inventor.Application
    Private dcWkg As Scripting.Dictionary
    Private dcFiled As Scripting.Dictionary
    Private dcIndex As Scripting.Dictionary

    Private Fm As FmIfcTest05A

    Private Sub Class_Initialize()
        '
        dcWkg = New Scripting.Dictionary
        dcFiled = New Scripting.Dictionary
        dcIndex = New Scripting.Dictionary

        Fm = New FmIfcTest05A
    End Sub

    Private Sub Class_Terminate()
        '
        Fm = Nothing

        dcWkg.RemoveAll()
        dcFiled.RemoveAll()
        dcIndex.RemoveAll()

        dcWkg = Nothing
        dcFiled = Nothing
        dcIndex = Nothing
    End Sub

    Public Function Itself() As wkgCls0
        Itself = Me
    End Function

    Public Function UseDict(
    Optional AiDoc As Inventor.Document = Nothing
) As wkgCls0
        If AiDoc Is Nothing Then
            Return Me
        Else
            Return Collect(AiDoc)
        End If
    End Function

    Public Function Process(
    Optional AiDoc As Inventor.Document = Nothing
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        If AiDoc Is Nothing Then
            ' THIS is where we start processing
            ' Inventor Documents collected
            ' in the internal Dictionary

            ' at the moment, we simply pull up
            ' the standard form and present it
            ' with the current collection
            ' of Inventor Documents
            With Fm.UseDict(dcWkg)
                .Show(1)
                Stop
            End With

            rt = DcCopy(dcWkg)
            ' Right now, all this does
            ' is copy out the internal
            ' Dictionary as is.
        Else
            rt = Collect(AiDoc).Process()
        End If

        Process = rt
        'Debug.Print(nu_wkgCls0().Collect().Process().Count
    End Function

    Public Function Collect(
    Optional AiDoc As Inventor.Document = Nothing
) As wkgCls0 'Scripting.Dictionary
        ',optional dcWkg as Scripting.Dictionary=nothing
        '
        ' Method Function Collect
        '
        '     given a valid Inventor Document
        '     (usually an assembly), gather
        '     any and all Parts in it into
        '     the internal Dictionary dcWkg,
        '     and return a copy.
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim ThisDocument As Inventor.Document
        If AiDoc Is Nothing Then
            Collect(InventorApp.ActiveDocument)

        ElseIf AiDoc Is ThisDocument Then
            'ElseIf AiDoc.DocumentType =DocumentTypeEnum.kAssemblyDocumentObject Then
            'ElseIf AiDoc.DocumentType =DocumentTypeEnum.kPartDocumentObject Then
        Else
            dcWkg =
            dcAiDocGrpsByForm(
            dcRemapByPtNum(
            dcAiDocComponents(
            AiDoc, , 0
        )))
            'With dcWkg: For Each ky In .Keys
            '    If dcFiled.Exists(ky) Then
            '        Stop 'going to have to deal
            '        'with merging two Dictionaries
            '    Else
            '        dcFiled.Add(ky, .Item(ky)
            '    End If
            '
            'If dcWkg.Exists(ky) Then
            '    If .Item(ky) Is dcWkg.Item(ky) Then
            '        Stop 'should be okay
            '        'but want to be sure
            '        'this can even happen
            '    Else
            '        Stop 'to deal with part number conflict
            '        'there may be several possibilities to deal with
            '        'which will likely require a separate function.
            '    End If
            'Else
            '    dcWkg.Add(ky, .Item(ky)
            'End If
            'Next: End With
        End If

        Collect = Me 'dcCopy(dcWkg)
    End Function

End Class