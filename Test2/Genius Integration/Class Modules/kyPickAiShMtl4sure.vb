Class KyPickAiShMtl4sure


    Inherits KyPick

    Private pk As KyPick
    Private Const txVersion As String = "KyPickAiShMtl4sure v0.0.0.1 [2022.03.08.1336]"
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
        Dim dcI As Scripting.Dictionary
        Dim dcO As Scripting.Dictionary
        Dim dCk As Scripting.Dictionary
        Dim ky As Object

        With pk.AfterScanning(dSrc)
            dcI = .DcIn
            dcO = .DcOut
        End With

        With dcI : For Each ky In .Keys
                dCk = KyPick_DcFor(.Item(ky))
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
        End If
    End Function
    '
    '
    ' General Class handling code follows
    '

    Private Sub Class_Initialize()
        pk = New KyPickAiSheetMetal 'KyPick
    End Sub
    '
    '
    ' KyPickAiShMtl4sure Class
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