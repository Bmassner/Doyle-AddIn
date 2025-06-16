Class KyPickInGenius

    Inherits KyPick

    'Private Const sqlA As String = "select ItemId from vgMfiItems where Item='"
    Private Const sqlA As String = "select count(ItemId) as ct from vgMfiItems where Item='"

    Private pk As KyPick
    Private cn As ADODB.Connection
    'Private cm As ADODB.Command

    Private Sub Class_Initialize()
        pk = New KyPick
        cn = CnGnsDoyle()
        'dcGeniusItems
        ' cm = New ADODB.Command
        'With cm
        '.CreateParameter(
        'End With
    End Sub

    Public Function dcFor(Item As Object) As Scripting.IDictionary
        Dim ob As Inventor.Document
        Dim pn As String

        ob = aiDocument(obOf(Item))
        If ob Is Nothing Then
            dcFor = pk.DcFor(0)
        Else
            If ob.DocumentType = DocumentTypeEnum.kPartDocumentObject _
        Or ob.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                pn = ob.Propertys.Item(gnDesign).Item(pnPartNum).Text
                If Len(pn) > 0 Then
                    With cn.Execute(sqlA & pn & "'")
                        'If .BOF Or .EOF Then
                        If .Fields("ct").Value > 0 Then
                            dcFor = pk.DcFor(ob)
                        Else
                            dcFor = pk.DcFor(0)
                        End If
                    End With
                Else
                    dcFor = pk.DcFor(0)
                End If
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