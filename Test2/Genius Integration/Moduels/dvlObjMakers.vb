Module dvlObjMakers

    Private Const txVersion As String = "dvlObjMakers REV[2022.03.16.0930]"
    Dim ThisApplication As Inventor.Application
    Public Function nu_wkgCls0(
    Optional AiDoc As Inventor.Document = Nothing
) As wkgCls0
        With New wkgCls0
            nu_wkgCls0 = .UseDict(AiDoc)
        End With
    End Function

    Public Function nu_gnsIfcAiDoc() As GnsIfcAiDoc
        nu_gnsIfcAiDoc = New GnsIfcAiDoc
    End Function

    Public Function nuILogicIfc(
    Optional message As Inventor.Document = Nothing
) As iLogicIfc
        With New iLogicIfc 'ifcVault
            If .RuleSource Is message Then
                If message Is Nothing Then
                    nuILogicIfc =
            .WithRulesIn(message)
                Else
                    nuILogicIfc = .Itself
                End If
            ElseIf message Is Nothing Then
                nuILogicIfc = .Itself
            Else
                nuILogicIfc =
        .WithRulesIn(message)
            End If
        End With
        'Debug.Print(txDumpLs(dcOb(nuILogicIfc().Apply("ilRuleText", nuAiNameValMap()).Item("OUT")).Keys)
        'Debug.Print(nuSelectorV2().WithList(nuILogicIfc().Apply("ilRuleText", nuAiNameValMap).Item("OUT")).GetReply()
        '{mod from Debug.Print(nuSelectorV2().WithList(nuIfcVault().iLogCall("ilRuleText", New Scripting.Dictionary).Item("OUT")).GetReply()}
    End Function

    Public Function nu_fmIfcTest04A(
    Optional About As Scripting.Dictionary = Nothing
) As FmIfcTest04A
        With New FmIfcTest04A
            nu_fmIfcTest04A _
        = .UseDict(About)
        End With
        'Debug.Print(nu_fmIfcTest04A().SeeUser().Version()
    End Function

    Public Function nu_fmIfcMatlQty01() As FmIfcMatlQty01
        nu_fmIfcMatlQty01 = New FmIfcMatlQty01
        'Debug.Print(nu_fmIfcMatlQty01().SeeUser().Version()
    End Function

    Public Function nu_FmGetList() As FmGetList
        nu_FmGetList = New FmGetList
    End Function

    Public Function newFmTest0() As FmTest0
        newFmTest0 = New FmTest0
    End Function
    'Debug.Print(newFmTest0().ft0g0f0(aiDocument(ThisApplication.ActiveDocument).Thumbnail)

    Public Function newFmTest1() As FmTest1
        newFmTest1 = New FmTest1
    End Function

    Public Function newFmTest2() As FmTest2
        newFmTest2 = New FmTest2
    End Function

    Public Function nuAiBoxData() As AiBoxData
        ''  Using "blank" version at this point
        nuAiBoxData = nuAiBoxDataRC0()
    End Function
    'Debug.Print(nuAiBoxData().UsingInches().Sorted(aiDocPart(aiDocActive()).ComponentDefinition.RangeBox).Dump(0)

    '
    ' TESTING SECTION
    '

    Public Function tstFmTest1()
        Dim ky As Object
        Dim nm As String

        'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-047.ipt"
        nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-072.ipt"
        'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-050.ipt"
        'nm = "C:\Doyle_Vault\Designs\doyle\(29) Field Loader Conveyor\29-051.ipt"

        With newFmTest1()
            If .AskAbout(
            ThisApplication.Documents.ItemByName(nm)
        ) = vbYes Then
                With .ItemData
                    For Each ky In .Keys
                        Debug.Print(ky, .Item(ky))
                    Next
                    Stop
                End With
            Else
            End If
        End With
    End Function

    '
    ' VERSION /ObjectSECTION
    '

    Public Function nuAiBoxDataRC1(arg1 As Object,
    Optional UseInches As Long = -1
) As AiBoxData
        Dim ob As Object
        Dim rt As AiBoxData

        If UseInches < 0 Then
            If arg1 Is Nothing Then
            ElseIf TypeOf arg1 Is Object Then
            Else
            End If
        Else
        End If

        With New AiBoxData
            rt = .UsingInches(UseInches)
        End With

        nuAiBoxDataRC1 = rt
    End Function

    Public Function nuAiBoxDataRC0() As AiBoxData
        nuAiBoxDataRC0 = New AiBoxData
    End Function

    '
    ' END of MODULE dvlObjMakers
    '
    Public Function dvlObjMakers() As String
        dvlObjMakers = txVersion
    End Function
End Module