Class KyPickAiDocContentCtr

    Inherits KyPick

    Private pk As KyPick

    Private Sub Class_Initialize()
        pk = New KyPick
    End Sub

    Public Overloads Function DcFor(Item As Object) As Scripting.IDictionary
        Dim ob As Inventor.Document

        ob = aiDocument(obOf(Item))
        If ob Is Nothing Then
            DcFor = pk.DcFor(0)
        Else
            If ob.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                If aiDocPart(ob).ComponentDefinition.IsContentMember Then
                    DcFor = pk.DcFor(ob)
                Else
                    DcFor = pk.DcFor(0)
                End If
            Else
                DcFor = pk.DcFor(0)
            End If
        End If
    End Function

    Public Overloads Function WithInDc(
    Dict As Scripting.Dictionary
    ) As KyPick
        pk = pk.WithInDc(Dict)
        WithInDc = Me
    End Function

    Public Overloads Function WithOutDc(
    Dict As Scripting.Dictionary
    ) As KyPick
        pk = pk.WithOutDc(Dict)
        WithOutDc = Me
    End Function

    Public Overloads Function DcIn() As Scripting.Dictionary
        DcIn = pk.DcIn
    End Function

    Public Overloads Function DcOut() As Scripting.Dictionary
        DcOut = pk.DcOut
    End Function

    Public Overloads Function AfterScanning(
    dSrc As Scripting.Dictionary
    ) As KyPick
        AfterScanning = KyPick_AfterScanning(dSrc)
    End Function

    Private Function KyPick_AfterScanning(dSrc As Scripting.IDictionary) As KyPick
        Dim ky As Object

        With dSrc
            For Each ky In .Keys
                With DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If
                End With
            Next
        End With
        KyPick_AfterScanning = Me
    End Function

    '
    ' KyPick Implementation code follows
    '
    Private Function KyPick_DcFor(Item As Object) As Scripting.IDictionary
        KyPick_DcFor = DcFor(Item)
    End Function

    Private Function KyPick_DcIn() As Scripting.IDictionary
        KyPick_DcIn = DcIn()
    End Function

    Private Function KyPick_DcOut() As Scripting.IDictionary
        KyPick_DcOut = DcOut()
    End Function

    Public Overloads Function Itself() As KyPick
        Itself = Me
    End Function

    Private Function KyPick_Itself() As KyPick
        KyPick_Itself = Me.Itself
    End Function

    Private Function KyPick_WithInDc(Dict As Scripting.IDictionary) As KyPick
        KyPick_WithInDc = WithInDc(Dict)
    End Function

    Private Function KyPick_WithOutDc(Dict As Scripting.IDictionary) As KyPick
        KyPick_WithOutDc = WithOutDc(Dict)
    End Function
End Class