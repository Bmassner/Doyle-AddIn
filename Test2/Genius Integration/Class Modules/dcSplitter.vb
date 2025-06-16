Class DcSplitter
    'Private dcGrpIn As Scripting.Dictionary
    'Private dcGrpOut As Scripting.Dictionary
    Private dcPicker As KyPick

    ' Sample Usage:
    'Debug.Print(txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).OutGroup().Keys)
    'Debug.Print(txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).WithSel(New kyPickAiDocWithRM, 1).SubScanning().OutGroup().Keys)

    Private Sub Class_Initialize()
        ' dcGrpIn = New Scripting.Dictionary
        ' dcGrpOut = New Scripting.Dictionary
        dcPicker = New KyPick
    End Sub

    Public Function WithInDc(
    Dict As Scripting.Dictionary
) As DcSplitter
        dcPicker = dcPicker.WithInDc(Dict)
        WithInDc = Me
    End Function

    Public Function WithOutDc(
    Dict As Scripting.Dictionary
) As DcSplitter
        dcPicker = dcPicker.WithOutDc(Dict)
        WithOutDc = Me
    End Function

    Public Function WithSel(Selector As KyPick,
    Optional KeepData As Long = 0
) As DcSplitter
        If KeepData = 0 Then
            dcPicker = Selector
        Else
            dcPicker = Selector.WithInDc(
            dcPicker.DcIn).WithOutDc(
            dcPicker.DcOut
        )
        End If
        WithSel = Me
    End Function

    Public Function InGroup() As Scripting.Dictionary
        InGroup = dcPicker.DcIn
    End Function

    Public Function OutGroup() As Scripting.Dictionary
        OutGroup = dcPicker.DcOut
    End Function

    Public Function Scanning(SrcDict As Scripting.Dictionary) As DcSplitter
        Dim ky As Object

        With SrcDict
            For Each ky In .Keys
                With dcPicker.DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, SrcDict.Item(ky))
                    End If
                End With
            Next
        End With
        Scanning = Me
    End Function

    Public Function SubScanning(Optional WantOut As Long = 0) As DcSplitter
        Dim dcSub As Scripting.Dictionary

        If WantOut = 0 Then
            dcSub = dcPicker.DcIn
        Else
            dcSub = dcPicker.DcOut
        End If
        dcPicker = dcPicker.WithInDc(New Scripting.Dictionary).WithOutDc(New Scripting.Dictionary)

        SubScanning = Scanning(dcSub)
    End Function

End Class