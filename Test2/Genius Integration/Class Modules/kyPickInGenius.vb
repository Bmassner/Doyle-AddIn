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

    Public Overloads Function DcFor(Item As Object) As Scripting.IDictionary
        Dim ob As Inventor.Document
        Dim pn As String

        ob = aiDocument(obOf(Item))
        If ob Is Nothing Then
            DcFor = pk.DcFor(0)
        Else
            If ob.DocumentType = DocumentTypeEnum.kPartDocumentObject _
        Or ob.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                pn = ob.Propertys.Item(GnDesign).Item(PnPartNum).Text
                If Len(pn) > 0 Then
                    With cn.Execute(sqlA & pn & "'")
                        'If .BOF Or .EOF Then
                        If .Fields("ct").Value > 0 Then
                            DcFor = pk.DcFor(ob)
                        Else
                            DcFor = pk.DcFor(0)
                        End If
                    End With
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