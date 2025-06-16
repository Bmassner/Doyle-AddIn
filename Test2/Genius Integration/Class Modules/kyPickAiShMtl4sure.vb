Class KyPickAiShMtl4sure


    Inherits KyPick

    Private pk As KyPick
    Private Const txVersion As String = "kyPickAiShMtl4sure v0.0.0.1 [2022.03.08.1336]"
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
        Dim dcI As Scripting.Dictionary
        Dim dcO As Scripting.Dictionary
        Dim dCk As Scripting.Dictionary
        Dim ky As Object

        With pk.AfterScanning(dSrc)
            dcI = .DcIn
            dcO = .DcOut
        End With

        With dcI : For Each ky In .Keys
                dCk = kyPick_DcFor(.Item(ky))
                If dCk Is dcI Then 'should be okay
                    'don't need to do anything here
                    Debug.Print("") 'Breakpoint Landing
                Else
                    With dCk
                        If .Exists(ky) Then
                            Stop
                        Else
                            .Add(ky, dcI.Item(ky))
                        End If : End With
                End If
            Next : End With
        pk = pk.WithInDc(DcKeysMissing(dcI, dcO))

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
        End If
    End Function
    '
    '
    ' General Class handling code follows
    '

    Private Sub Class_Initialize()
        pk = New kyPickAiSheetMetal 'kyPick
    End Sub
    '
    '
    ' kyPickAiShMtl4sure Class
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
            g0f1 = g0f2(ob)
        Else
            g0f1 = pk.DcFor(0)
        End If
    End Function

    Private Function g0f2(
    ob As Inventor.SheetMetalComponentDefinition
) As Scripting.Dictionary
        Dim smThk As Double
        Dim fpHgt As Double
        Dim dfRns As Double

        If ob Is Nothing Then
            g0f2 = pk.DcFor(0)
        Else
            With ob
                If .HasFlatPattern Then 'MIGHT be
                    ''  check stated thickness...
                    smThk = .Thickness.Text
                    'Debug.Print("Thickness: " & CStr(smThk)

                    ''  against flat pattern height
                    With nuAiBoxData().UsingBox(
                .FlatPattern.RangeBox
            )
                        'Debug.Print(.Dump()
                        Debug.Print("") 'Breakpoint Landing
                        fpHgt = .SpanZ
                    End With

                    dfRns = fpHgt - smThk
                    If Math.Abs(dfRns) < 0.001 Then 'close enough
                        ''  assume it's valid
                        g0f2 = pk.DcFor(ob.Document)
                    Else 'something's screwy here
                        ''  assume likely not
                        'Stop
                        g0f2 = pk.DcFor(0)
                    End If
                Else 'likely not
                    g0f2 = pk.DcFor(0)
                End If : End With
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