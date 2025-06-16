Class KyPick

    Private dcGrpIn As Scripting.Dictionary
    Private dcGrpOut As Scripting.Dictionary

    Private Sub Class_Initialize()
        dcGrpIn = New Scripting.Dictionary
        dcGrpOut = New Scripting.Dictionary
    End Sub

    Public Function Itself() As KyPick
        Itself = Me
    End Function

    Public Function WithInDc(
    Dict As Scripting.Dictionary
) As KyPick
        dcGrpIn = DcNewIfNone(Dict)
        WithInDc = Me
    End Function

    Public Function WithOutDc(
    Dict As Scripting.Dictionary
) As KyPick
        dcGrpOut = DcNewIfNone(Dict)
        WithOutDc = Me
    End Function

    Public Function AfterScanning(
    dSrc As Scripting.Dictionary
) As KyPick
        Dim ky As Object

        With dSrc : For Each ky In .Keys
                With DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If
                End With
            Next : End With
        AfterScanning = Me
    End Function

    Public Function DcIn() As Scripting.Dictionary
        DcIn = dcGrpIn
    End Function

    Public Function DcOut() As Scripting.Dictionary
        DcOut = dcGrpOut
    End Function

    Public Function DcFor(
    Item As Object
) As Scripting.Dictionary
        If TypeOf Item Is Object Then
            DcFor = dcGrpIn
        Else
            DcFor = dcGrpOut
        End If
    End Function

End Class