Class KyPickAiDocPurchased

    Inherits KyPick

    Private pk As KyPick
    '
    '
    ' KyPick Implementation code follows
    '

    Private Function KyPick_Itself() As KyPick
        KyPick_Itself = Me.Itself
    End Function


    Private Function KyPick_WithInDc(
    Dict As Scripting.IDictionary
) As KyPick
        KyPick_WithInDc = WithInDc(Dict)
    End Function

    Private Function KyPick_WithOutDc(
    Dict As Scripting.IDictionary
) As KyPick
        KyPick_WithOutDc = WithOutDc(Dict)
    End Function


    Private Function KyPick_AfterScanning(
    dSrc As Scripting.IDictionary
) As KyPick
        Dim ky As Object

        With dSrc : For Each ky In .Keys
                With DcFor(.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dSrc.Item(ky))
                    End If : End With
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
        KyPick_DcFor = DcFor(Item)
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
    ' KyPickAiDocPurchased Class
    ' implementation code follows
    '

    Public Overloads Function Itself() As KyPick
        Itself = Me
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


    Public Overloads Function AfterScanning(
    dSrc As Scripting.Dictionary
    ) As KyPick
        AfterScanning = KyPick_AfterScanning(dSrc)
    End Function


    Public Overloads Function DcIn() As Scripting.Dictionary
        DcIn = pk.DcIn
    End Function

    Public Overloads Function DcOut() As Scripting.Dictionary
        DcOut = pk.DcOut
    End Function


    Public Overloads Function DcFor(Item As Object) As Scripting.IDictionary
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
            DcFor = pk.DcFor(ob)
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
                GnDesign).Item(PnFamily
            )
                If InStr(1, ob.FullFileName,
                "\Doyle_Vault\Designs\purchased\"
            ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                "|" & InvProperty.Text & "|"
            ) > 0 Then
                    DcFor = pk.DcFor(ob)
                Else
                    DcFor = pk.DcFor(0)
                End If
            End With
        End If
    End Function

End Class