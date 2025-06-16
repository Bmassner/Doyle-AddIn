Class kyPickAiDocPurchased

    Inherits KyPick

    Private pk As KyPick
    '
    '
    ' kyPick Implementation code follows
    '

    Private Function kyPick_Itself() As KyPick
        kyPick_Itself = Me.Itself
    End Function


    Private Function kyPick_WithInDc(
    Dict As Scripting.IDictionary
) As KyPick
        kyPick_WithInDc = WithInDc(Dict)
    End Function

    Private Function kyPick_WithOutDc(
    Dict As Scripting.IDictionary
) As KyPick
        kyPick_WithOutDc = WithOutDc(Dict)
    End Function


    Private Function kyPick_AfterScanning(
    dSrc As Scripting.IDictionary
) As KyPick
        Dim ky As Object

        With dSrc : For Each ky In .Keys
                With dcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If : End With
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
        kyPick_DcFor = dcFor(Item)
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
    ' kyPickAiDocPurchased Class
    ' implementation code follows
    '

    Public Function Itself() As KyPick
        Itself = Me
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


    Public Function AfterScanning(
    dSrc As Scripting.Dictionary
) As KyPick
        AfterScanning = kyPick_AfterScanning(dSrc)
    End Function


    Public Function dcIn() As Scripting.Dictionary
        dcIn = pk.DcIn
    End Function

    Public Function dcOut() As Scripting.Dictionary
        dcOut = pk.DcOut
    End Function


    Public Function dcFor(Item As Object) As Scripting.IDictionary
        Dim ck As Inventor.BOMStructureEnum
        Dim ob As Inventor.Document
        Dim InvProperty As Inventor.Property
        ' REV[2022.03.08.1021]
        '     Added BOMStructureEnum variable ck
        '     to collect BOMStructureEnum for each
        '     relevant Document type, and consolidate
        '     BOMStructureEnum check to one block
        '     following Doc type accommodation.

        ob = aiDocument(obOf(Item))

        If ob Is Nothing Then
            ck = BOMStructureEnum.kDefaultBOMStructure
        ElseIf ob.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            ck = aiDocPart(ob).ComponentDefinition.BOMStructure
        ElseIf ob.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ck = aiDocAssy(ob).ComponentDefinition.BOMStructure
        Else
            ck = BOMStructureEnum.kDefaultBOMStructure
        End If

        If ck = BOMStructureEnum.kPurchasedBOMStructure Then
            dcFor = pk.DcFor(ob)
        Else
            ' REV[2022.03.08.1038]
            '     Additional checks on Item
            '     Family and File Location
            '     NOTE that this is more of
            '     a "soft" identification
            '     of likely purchased parts,
            '     and might or might not be
            '     appropriate to apply.
            With ob
                InvProperty = .Propertys.Item(
                gnDesign).Item(pnFamily
            )
                If InStr(1, ob.FullFileName,
                "\Doyle_Vault\Designs\purchased\"
            ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                "|" & InvProperty.Text & "|"
            ) > 0 Then
                    dcFor = pk.DcFor(ob)
                Else
                    dcFor = pk.DcFor(0)
                End If
            End With
        End If
    End Function

End Class