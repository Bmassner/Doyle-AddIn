Class kyPickAiSheetMetal

    Inherits KyPick

    Private pk As KyPick
    Private Const txVersion As String = "kyPickAiSheetMetal v0.0.0.1 [2022.03.08.1332]"
    ' prior Versions
    '     ""
    '
    ' kyPick Implementation code follows
    '

    Private Function kyPick_Itself() As KyPick
        kyPick_Itself = Me
    End Function


    Private Function kyPick_WithInDc(
    Dict As Scripting.IDictionary
) As KyPick
        pk = pk.WithInDc(Dict)
        kyPick_WithInDc = Me
    End Function

    Private Function kyPick_WithOutDc(
    Dict As Scripting.IDictionary
) As KyPick
        pk = pk.WithOutDc(Dict)
        kyPick_WithOutDc = Me
    End Function


    Private Function kyPick_AfterScanning(
    dSrc As Scripting.IDictionary
) As KyPick
        Dim ky As Object

        With dSrc : For Each ky In .Keys
                With kyPick_DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If
                End With
            Next : End With
        kyPick_AfterScanning = Me
    End Function


    Private Function kyPick_DcIn() As Scripting.IDictionary
        kyPick_DcIn = dcIn()
    End Function

    Private Function kyPick_DcOut() As Scripting.IDictionary
        kyPick_DcOut = dcOut()
    End Function


    Private Function kyPick_DcFor(
    Item As Object
) As Scripting.IDictionary
        Dim ob As Inventor.PartDocument '.Document

        ob = aiDocPart(aiDocument(obOf(Item)))
        If ob Is Nothing Then
            kyPick_DcFor = pk.DcFor(0)
        Else
            kyPick_DcFor = g0f1(
            ob.ComponentDefinition
        )
            'If ob.DocumentType =DocumentTypeEnum.kPartDocumentObject Then
            '    If aiDocPart(ob).SubType = guidSheetMetal Then
            '         kyPick_DcFor = pk.dcFor(ob)
            '    Else
            '         kyPick_DcFor = pk.dcFor(0)
            '    End If
            'Else
            '     kyPick_DcFor = pk.dcFor(0)
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
    ' kyPickAiSheetMetal Class
    ' implementation code follows
    '

    Public Function Itself() As KyPick
        Itself = Me
    End Function


    Public Function WithInDc(
    Dict As Scripting.Dictionary
) As KyPick
        WithInDc = kyPick_WithInDc(Dict)
    End Function

    Public Function WithOutDc(
    Dict As Scripting.Dictionary
) As KyPick
        WithOutDc = kyPick_WithOutDc(Dict)
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


    Public Function dcFor(Item As Object) As Scripting.IDictionary
        dcFor = kyPick_DcFor(Item)
    End Function
    '
    '
    ' Internal support code follows
    '

    Private Function g0f0(
    ob As Inventor.PartDocument
) As Scripting.Dictionary
        If ob Is Nothing Then
            g0f0 = pk.DcFor(0)
        Else
            g0f0 = g0f1(ob.ComponentDefinition)
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