Class AiBoxData
    Dim inventorApp As Inventor.Application
    Private Const f01 As String = "#,##0.000"
    'Private Const f02 As String = "#,##0.0000 '"

    Private bx As Inventor.Box
    Private mn As Inventor.Point
    Private mx As Inventor.Point

    Private sc As Double

    Private Sub Class_Initialize()
        sc = 1.0#
    End Sub

    Public Property Box(ThisBox As Box) As Inventor.Box
        Get
            Return bx
        End Get
        Set
            bx = ThisBox
            mn = bx.MinPoint
            mx = bx.MaxPoint
        End Set
    End Property

    Public Property Box() As Inventor.Box
        Get
            Box = bx
        End Get
        Set(value As Inventor.Box)
            bx = value
            mn = bx.MinPoint
            mx = bx.MaxPoint
        End Set
    End Property

    Public Function UsingBox(
        ThisOne As Inventor.Box
    ) As AiBoxData
        Me.Box = ThisOne
        UsingBox = Me
    End Function

    Public Function UsingOrBox(
    ThisOne As Inventor.OrientedBox
) As AiBoxData
        bx = InventorApp.TransientGeometry.CreateBox()

        With ThisOne
            bx.Extend(InventorApp.TransientGeometry.CreatePoint(.DirectionOne.Length, .DirectionTwo.Length, .DirectionThree.Length))
        End With

        'Me.Box = ThisOne
        UsingOrBox = Me.UsingBox(bx)
        'Debug.Print(nuAiBoxData().UsingInches.UsingOrBox(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences.Item(1).OrientedMinimumRangeBox).Dump()
    End Function

    Public Function UsingBoxOb(
    ThisOne As Object
) As AiBoxData
        If ThisOne Is Nothing Then
            UsingBoxOb = Me
        ElseIf TypeOf ThisOne Is Inventor.Box Then
            UsingBoxOb = UsingBox(ThisOne)
        ElseIf TypeOf ThisOne Is Inventor.OrientedBox Then
            UsingBoxOb = UsingOrBox(ThisOne)
        Else
            UsingBoxOb = Me
        End If
    End Function

    Public Function UsingModel(
       ThisOne As Inventor.Document,
       Optional Oriented As Long = 0
    ) As AiBoxData
        UsingModel _
            = UsingPart(aiDocPart(ThisOne), Oriented
            ).UsingAssy(aiDocAssy(ThisOne), Oriented
        )
    End Function

    Public Function UsingPart(
       ThisOne As Inventor.PartDocument,
       Optional Oriented As Long = 0
    ) As AiBoxData
        If ThisOne Is Nothing Then
            UsingPart = Me
        Else
            With ThisOne.ComponentDefinition
                UsingPart = UsingBoxOb(IIf(Oriented = 0,
                .RangeBox, .OrientedMinimumRangeBox
            ))
            End With
        End If
    End Function

    Public Function UsingAssy(
       ThisOne As Inventor.AssemblyDocument,
       Optional Oriented As Long = 0
    ) As AiBoxData
        If ThisOne Is Nothing Then
            UsingAssy = Me
        Else
            With ThisOne.ComponentDefinition
                UsingAssy = UsingBoxOb(IIf(Oriented = 0,
                .RangeBox, .OrientedMinimumRangeBox
            ))
            End With
            UsingAssy = UsingBox(
        ThisOne.ComponentDefinition.RangeBox
        )
        End If
    End Function

    Public Function SortingDims(
        Optional ThisBox As Inventor.Box = Nothing
    ) As AiBoxData
        If ThisBox Is Nothing Then
            If bx Is Nothing Then
                SortingDims = Me
            Else
                SortingDims = SortingDims(bx)
            End If
        Else
            Me.Box = aiBoxSortDown(ThisBox)
            SortingDims = Me
        End If
    End Function

    Private Function Span(
        ptMin As Double,
        ptMax As Double
    ) As Double
        Span = sc * (ptMax - ptMin)
    End Function

    Public Function SpanX() As Double
        SpanX = Span(mn.X, mx.X)
    End Function

    Public Function SpanY() As Double
        SpanY = Span(mn.Y, mx.Y)
    End Function

    Public Function SpanZ() As Double
        SpanZ = Span(mn.Z, mx.Z)
    End Function

    Public Function SpansXYZ() As Double()
        Dim rt(2) As Double

        rt(0) = SpanX()
        rt(1) = SpanY()
        rt(2) = SpanZ()

        SpansXYZ = rt
    End Function

    Public Function SpansOrdered() As Double()
        SpansOrdered = sort3dimsUp(SpanX, SpanY, SpanZ)
    End Function

    Public Function UsingInches(Optional Yes As Long = 1) As AiBoxData
        If Yes Then sc = 1 / 2.54 Else sc = 1
        UsingInches = Me
    End Function

    Public Function Dump(Optional Form As Long = 0) As String
        Dump = ""
        'ConvertToJson(nuDcPopulator().Setting("X SPAN", Format$(me.SpanX, "#,##0.0000 '")).Setting("Y SPAN", Format$(me.SpanY, "#,##0.0000 '")).Setting("Z SPAN", Format$(me.SpanZ, "#,##0.0000 '")).Dictionary,vbTab)
        'ConvertToJson(nuDcPopulator().Setting("X SPAN", System.Math.Round(me.SpanX,4)).Setting("Y SPAN", System.Math.Round(me.SpanY,4)).Setting("Z SPAN", System.Math.Round(me.SpanZ,4)).Dictionary,vbTab)
        Select Case Form
            Case 67518582
                With NuDcPopulator().Setting("X SPAN", Format$(Me.SpanX, "#,##0.0000 '")).Setting("Y SPAN", Format$(Me.SpanY, "#,##0.0000 '")).Setting("Z SPAN", Format$(Me.SpanZ, "#,##0.0000 '")).Dictionary
                    '
                End With
            Case Else
                Dump = "X SPAN" & vbTab & "Y SPAN" & vbTab & "Z SPAN" _
                & vbCrLf & Format(Me.SpanX, f01) _
                & vbTab & Format(Me.SpanY, f01) _
                & vbTab & Format(Me.SpanZ, f01)
        End Select
    End Function

    Public Function Dictionary(
        Optional Form As Long = 3
    ) As Scripting.Dictionary
        '
        ' Dictionary -- return Dictionary of dimensions
        '     keyed according to Form, a sum of:
        '     1 - "X", "Y", "Z", per Model
        '     2 - magnitudes "Min", "Mid", "Max"
        '         (note that sorting keys in descending order
        '          produces values sorted in ascending order)
        '     3 - BOTH sets of keys (1 + 2)
        '
        ' REV[2022.08.31.1444] Method Dictionary
        ' added to Class to support extraction
        ' of Dictionary Object for data export
        ' (see dcGnsPtProps_Rev20220830_inProg)
        '
        Dim rt As Scripting.Dictionary
        Dim dm() As Double

        If (Form And 3) = 0 Then
            rt = Dictionary(3)
        Else
            rt = New Scripting.Dictionary

            With rt
                If Form And 1 Then 'add XYZ entries
                    .Add("X", SpanX())
                    .Add("Y", SpanY())
                    .Add("Z", SpanZ())
                End If

                If Form And 2 Then 'add Min, Mid, Max entries
                    dm = SpansOrdered()
                    .Add("Min", dm(0))
                    .Add("Mid", dm(1))
                    .Add("Max", dm(2))
                End If
            End With
        End If

        Dictionary = rt
    End Function

End Class