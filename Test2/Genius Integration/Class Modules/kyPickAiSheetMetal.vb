Class KyPickAiSheetMetal

    Inherits KyPick

    Private pk As KyPick
    Private Const txVersion As String = "KyPickAiSheetMetal v0.0.0.1 [2022.03.08.1332]"
    ' prior Versions
    '     ""
    '
    ' KyPick Implementation code follows
    '

    Private Function KyPick_Itself() As KyPick
        KyPick_Itself = Me
    End Function


    Private Function KyPick_WithInDc(
    Dict As Scripting.IDictionary
) As KyPick
        pk = pk.WithInDc(Dict)
        KyPick_WithInDc = Me
    End Function

    Private Function KyPick_WithOutDc(
    Dict As Scripting.IDictionary
) As KyPick
        pk = pk.WithOutDc(Dict)
        KyPick_WithOutDc = Me
    End Function


    Private Function KyPick_AfterScanning(
    dSrc As Scripting.IDictionary
) As KyPick
        Dim ky As Object

        With dSrc : For Each ky In .Keys
                With KyPick_DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If
                End With
            Next : End With
        KyPick_AfterScanning = Me
    End Function


    Private Function KyPick_DcIn() As Scripting.IDictionary
        KyPick_DcIn = DcIn()
    End Function

    Private Function KyPick_DcOut() As Scripting.IDictionary
        KyPick_DcOut = DcOut()
    End Function


    Private Function KyPick_DcFor(
    Item As Object
) As Scripting.IDictionary
        Dim ob As Inventor.PartDocument '.Document

        ob = aiDocPart(aiDocument(obOf(Item)))
        If ob Is Nothing Then
            KyPick_DcFor = pk.DcFor(0)
        Else
            KyPick_DcFor = g0f1(
            ob.ComponentDefinition
        )
            'If ob.DocumentType =DocumentTypeEnum.kPartDocumentObject Then
            '    If aiDocPart(ob).SubType = guidSheetMetal Then
            '         KyPick_DcFor = pk.dcFor(ob)
            '    Else
            '         KyPick_DcFor = pk.dcFor(0)
            '    End If
            'Else
            '     KyPick_DcFor = pk.dcFor(0)
            'End If
        End If
    End Function
    '
    '
    ' General Class handling code follows
    '

    Private Sub Class_Initialize()
        pk = New KyPick
    End Sub
    '
    '
    ' KyPickAiSheetMetal Class
    ' implementation code follows
    '

    Public Overloads Function Itself() As KyPick
        Itself = Me
    End Function


    Public Overloads Function WithInDc(
    Dict As Scripting.Dictionary
    ) As KyPick
        WithInDc = KyPick_WithInDc(Dict)
    End Function

    Public Overloads Function WithOutDc(
    Dict As Scripting.Dictionary
    ) As KyPick
        WithOutDc = KyPick_WithOutDc(Dict)
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


    Public Overloads Function DcFor(Item As Object) As Scripting.IDictionary
        DcFor = KyPick_DcFor(Item)
    End Function
    '
    '
    ' Internal support code follows
    '

    Private Function G0f0(
    ob As Inventor.PartDocument
) As Scripting.Dictionary
        If ob Is Nothing Then
            G0f0 = pk.DcFor(0)
        Else
            G0f0 = g0f1(ob.ComponentDefinition)
        End If
    End Function

    Private Function g0f1(
    ob As Inventor.PartComponentDefinition
) As Scripting.Dictionary
        If TypeOf ob Is Inventor.SheetMetalComponentDefinition Then
            g0f1 = pk.DcFor(ob.Document)
        Else
            g0f1 = pk.DcFor(0)
        End If
    End Function
    '
    '
    ' Version code follows
    '

    Public Function Version() As String
        Version = txVersion
    End Function
    '
    ' End of Module
    '
End Class