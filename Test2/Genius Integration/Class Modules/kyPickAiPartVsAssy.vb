Class kyPickAiPartVsAssy

    Inherits KyPick

    Private pk As KyPick

    Private Sub Class_Initialize()
        pk = New KyPick
    End Sub

    Public Function dcFor(Item As Object) As Scripting.IDictionary
        Dim ob As Inventor.Document

        ob = aiDocument(obOf(Item))
        If ob Is Nothing Then
            dcFor = pk.DcFor(0)
        Else
            If ob.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                dcFor = pk.DcFor(ob)
            Else
                dcFor = pk.DcFor(0)
            End If
        End If
    End Function

    Public Function WithInDc(
    Dict As Scripting.Dictionary
) As KyPick
        pk = pk.WithInDc(Dict)
        WithInDc = Me
    End Function

    Public Function WithOutDc(
    Dict As Scripting.Dictionary
) As KyPick
        pk = pk.WithOutDc(Dict)
        WithOutDc = Me
    End Function

    Public Function dcIn() As Scripting.Dictionary
        dcIn = pk.DcIn
    End Function

    Public Function dcOut() As Scripting.Dictionary
        dcOut = pk.DcOut
    End Function

    Public Function AfterScanning(
    dSrc As Scripting.Dictionary
) As KyPick
        AfterScanning = kyPick_AfterScanning(dSrc)
    End Function

    Private Function kyPick_AfterScanning(dSrc As Scripting.IDictionary) As KyPick
        Dim ky As Object

        With dSrc
            For Each ky In .Keys
                With dcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If
                End With
            Next
        End With
        kyPick_AfterScanning = Me
    End Function

    '
    ' kyPick Implementation code follows
    '
    Private Function kyPick_DcFor(Item As Object) As Scripting.IDictionary
        kyPick_DcFor = dcFor(Item)
    End Function

    Private Function kyPick_DcIn() As Scripting.IDictionary
        kyPick_DcIn = dcIn()
    End Function

    Private Function kyPick_DcOut() As Scripting.IDictionary
        kyPick_DcOut = dcOut()
    End Function

    Public Function Itself() As KyPick
        Itself = Me
    End Function

    Private Function kyPick_Itself() As KyPick
        kyPick_Itself = Me.Itself
    End Function

    Private Function kyPick_WithInDc(Dict As Scripting.IDictionary) As KyPick
        kyPick_WithInDc = WithInDc(Dict)
    End Function

    Private Function kyPick_WithOutDc(Dict As Scripting.IDictionary) As KyPick
        kyPick_WithOutDc = WithOutDc(Dict)
    End Function
End Class