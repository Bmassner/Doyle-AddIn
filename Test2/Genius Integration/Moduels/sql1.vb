Module sql1
    Dim inventorApp As Inventor.Application
    Public Function sqlListValues(
    dc As Scripting.Dictionary,
    Optional sep As String = "', '",
    Optional pfx As String = "('",
    Optional sfx As String = "')"
) As String
        '
        '
        '
        sqlListValues = pfx & Join(dc.Keys, sep) & sfx
        '"')," & vbCrLf & vbTab & "('"
    End Function

    Public Function sqlValSelFromDict(
    dc As Scripting.Dictionary
) As String
        '
        '
        '
        sqlValSelFromDict = sqlListValues(dc,
        "')," & vbCrLf & vbTab & "('",
        "(select pn from (VALUES ('",
        "')" & vbCrLf & ") as a(pn))"
    )
        'With dc
        'sqlValSelFromDict = "(select pn from (VALUES " _
        '& vbCrLf & vbTab & "('" _
        '& Join(.Keys, "')," & vbCrLf & vbTab & "('") _
        '& "')" & vbCrLf & ") as a(pn))" _
        '& ""
        'End With
    End Function

    Public Function sqlValsFromDict(
    dc As Scripting.Dictionary,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        '
        '
        sqlValsFromDict = sqlListValues(dc,
        "')," & vbCrLf & vbTab & "('",
        "(values ('", "')" & vbCrLf & ") as " _
        & lsName & "(" & fdName & ")"
    )
    End Function

    Public Function sqlValsFromAssy(
    AiDoc As Inventor.Document,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        '
        '
        Dim dc As Scripting.Dictionary
        Dim ck As String

        dc = dcRemapByPtNum(dcAiDocComponents(AiDoc))

        ck = sqlListValues(dc,
        "')," & vbCrLf & vbTab & "('",
        "(values ('", "')" & vbCrLf & ") as " _
        & lsName & "(" & fdName & ")"
    )
        If sqlValsFromDict(dc) = ck Then
            Stop
        End If
        sqlValsFromAssy = ck
    End Function

    Public Function q1g0x0(AiDoc As Inventor.Document) As String
        '
        ' SQL text function naming convention
        '     q1 - "q" for "query", with module number
        '     g1 - "g" for "group" (typical usage)
        '     x1 - "x" for "text" (stands out better than "t")
        '
        q1g0x0 = "-- SQL text begins here" _
    & vbCrLf & "" _
    & vbCrLf & "-- SQL text ends here"
        '
    End Function

    Public Function q1g1x1(
    AiDoc As Inventor.Document,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        ' SQL text function naming convention
        '     q1 - "q" for "query", with module number
        '     g1 - "g" for "group" (typical usage)
        '     x1 - "x" for "text" (stands out better than "t")
        '
        q1g1x1 = "from vgMfiItems i inner join" _
        & vbCrLf & sqlValsFromAssy(
            AiDoc, lsName, fdName
        ) & vbCrLf & "on i.Item = " _
        & lsName & "." & fdName
    End Function

    Public Function q1g1x1v2(
    AiDoc As Inventor.Document,
    Optional gnsTbl As String = "vgMfiItems",
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        ' SQL text function naming convention
        '     q1 - "q" for "query", with module number
        '     g1 - "g" for "group" (typical usage)
        '     x1 - "x" for "text" (stands out better than "t")
        '

        'REV[2021.08.18] (REVERSED)
        '   changed inner join to right (outer) join
        '   to pick up Inventor Items not (yet) in Genius
        '   REVERSED -- since all returned fields are null,
        '   no information is returned for missing Items.
        q1g1x1v2 = "from " & gnsTbl & " i inner join" _
        & vbCrLf & sqlValsFromAssy(
            AiDoc, lsName, fdName
        ) & vbCrLf & "on i.Item = " _
        & lsName & "." & fdName
    End Function

    Public Function q1g1x1v3(
    dc As Scripting.Dictionary,
    Optional gnsTbl As String = "vgMfiItems",
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        ' SQL text function naming convention
        '     q1 - "q" for "query", with module number
        '     g1 - "g" for "group" (typical usage)
        '     x1 - "x" for "text" (stands out better than "t")
        '

        'REV[2021.08.18] (REVERSED)
        '   changed inner join to right (outer) join
        '   to pick up Inventor Items not (yet) in Genius
        '   REVERSED -- since all returned fields are null,
        '   no information is returned for missing Items.
        q1g1x1v3 = "from " & gnsTbl & " i inner join" _
        & vbCrLf & sqlValsFromDict(
            dc, lsName, fdName
        ) & vbCrLf & "on i.Item = " _
        & lsName & "." & fdName
    End Function

    Public Function q1g1x2(
    AiDoc As Inventor.Document,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        q1g1x2 = "select" & " i." &
    IIf(False, "*", Join({"Item" _
        , "Family", "Description1", "Description3", "Unit" _
        , "Thickness", "Width", "Length" _
        , "Height", "Diameter", "Weight" _
        , "Specification1", "Specification2", "Specification3" _
        , "Specification4", "Specification5", "Specification6" _
        , "Specification7", "Specification8", "Specification9", ", i." & vbCrLf & q1g1x1(AiDoc, lsName, fdName), "", "", ", " & lsName & "."}))

        'send2clipBd ConvertToJson(dcRecSetDcDx4json( _
        DcFromAdoRS(CnGnsDoyle().Execute(
        q1g1x2(InventorApp.ActiveDocument)), vbTab)
    End Function

    Public Function q1g1x2v2(
    dc As Scripting.Dictionary,
    Optional gnsTbl As String = "vgMfiItems",
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        q1g1x2v2 = "select" & " i." &
    IIf(False, "*", Join({"Item" _
        , "Family", "Description1", "Description3", "Unit" _
        , "Thickness", "Width", "Length" _
        , "Height", "Diameter", "Weight" _
        , "Specification1", "Specification2", "Specification3" _
        , "Specification4", "Specification5", "Specification6" _
        , "Specification7", "Specification8", "Specification9", ", i." & vbCrLf & q1g1x1v3(dc, gnsTbl, lsName, fdName), "", "", ", " & lsName & "."}))
        '
        '
        'send2clipBdWin10 ConvertToJson(dcRecDcDx4json( _
        dcRecSetDcDx4json(DcFromAdoRS(
            CnGnsDoyle().Execute(q1g1x2v2(aiDocActive())), vbTab))
        'Debug.Print(txDumpLs(dcKeysMissing( _
        dcRemapByPtNum(
        dcAiDocComponents(aiDocActive())
    )
        dcOb(dcRecSetDcDx4json(DcFromAdoRS(
        CnGnsDoyle().Execute(q1g1x2v2(aiDocActive()))
    )).Item("Item")).Keys()
    End Function

    Public Function q1g2x1(
    AiDoc As Inventor.Document,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        '
        ' SQL text function naming convention
        '     q1 - "q" for "query", with module number
        '     g1 - "g" for "group" (typical usage)
        '     x1 - "x" for "text" (stands out better than "t")
        '
        q1g2x1 = "from vgIcoBillOfMaterials b inner join" _
        & vbCrLf & sqlValsFromAssy(
            AiDoc, lsName, fdName
        ) & vbCrLf & "on i.Item = " _
        & lsName & "." & fdName
    End Function

    Public Function q1g2x2(
    AiDoc As Inventor.Document,
    Optional lsName As String = "ls",
    Optional fdName As String = "it"
) As String
        q1g2x2 = "select" & " b." &
    IIf(False, "*", Join({
        "Product", "ItemOrder", "Item",
        "QuantityInConversionUnit",
        "ConversionUnit",
        "ItemType", "Reserved" _
    , ", b." & vbCrLf _
    & q1g2x1(AiDoc, lsName, fdName)}))
        ' _
        '

        '& ", " & lsName & "." & join({ _
        '    "", "" _
        '), ", " & lsName & ".") _
        '
        '
        ''send2clipBd ConvertToJson(dcRecSetDcDx4json( _
        '    DcFromAdoRS(CnGnsDoyle().Execute(
        '        q1g2x2(InventorApp.ActiveDocument)
        '    )) _
        '), vbTab)
    End Function

    Public Function sqlSelAiPurch01fromTextV01(txList As String) As String
        '
        '
        '
        sqlSelAiPurch01fromTextV01 = "-- " _
    & vbCrLf & "with" _
        & vbCrLf & vbTab & "ls as " & txList _
    & vbCrLf & "select" _
        & vbCrLf & vbTab & "ls.pn, it.Type, it.Family" _
    & vbCrLf & "from" _
        & vbCrLf & vbTab & "ls inner join vgMfiItems as it" _
        & vbCrLf & vbTab & "on ls.pn = it.Item" _
    & vbCrLf & "where" _
        & vbCrLf & vbTab & "it.Type = 'R'" _
        & vbCrLf & vbTab & "or it.Family in (" _
        & "'D-HDWR', 'D-PTO', 'D-PTS', 'R-PTO', 'R-PTS'" _
        & ")" _
    & "" _
    & ""
    End Function

    Public Function sqlSelAiPurch01fromTextV02(
    txList As String
) As String
        '
        '
        '
        Dim n0 As Integer
        Dim n1 As Long
        Dim s0 As String
        Dim s1 As String
        Dim a0 As Object

        n0 = InStr(1, txList, "'")
        s0 = Mid$(txList, 1 + n0)
        n0 = InStr(1, s0, "'")
        n1 = InStr(1 + n0, s0, "'") '- n0 + 1
        If n1 > 0 Then
            n1 = n1 - n0 + 1
            s1 = Mid$(s0, n0, n1)
            a0 = Split(s0, s1)
            n0 = UBound(a0)
            a0(n0) = Split(a0(n0), "'")(0)
        Else
            a0 = {Left$(s0, n0 - 1)}
        End If
        's0 = Join(a0, "', '")

        Debug.Print("")

        sqlSelAiPurch01fromTextV02 = "-- " _
        & vbCrLf & "select" _
            & vbCrLf & vbTab & "it.Item, it.Type, it.Family" _
        & vbCrLf & "from" _
            & vbCrLf & vbTab & "vgMfiItems as it" _
        & vbCrLf & "where" _
            & vbCrLf & vbTab & "it.Item in ('" _
                & Join(a0, "', '") & "')" _
            & vbCrLf & vbTab & "and (it.Type = 'R'" _
            & vbCrLf & vbTab & "or it.Family in ('" _
                & Join({"D-HDWR",
                    "D-PTO", "D-PTS",
                    "R-PTO", "R-PTS" _
                , "', '" _
        & "'))" _
    & "" _
    & ""})
        '
        '& vbCrLf & "with" _
        '& vbCrLf & vbTab & "ls as " & txList _
        '
    End Function

    Public Function sqlSelAiPdParts01fromTextV01(
    txList As String
) As String
        '
        '
        '
        Dim n0 As Integer
        Dim n1 As Long
        Dim s0 As String
        Dim s1 As String
        Dim a0 As Object

        n0 = InStr(1, txList, "'")
        s0 = Mid$(txList, 1 + n0)
        n0 = InStr(1, s0, "'")
        n1 = InStr(1 + n0, s0, "'") - n0 + 1
        s1 = Mid$(s0, n0, n1)
        a0 = Split(s0, s1)
        n0 = UBound(a0)
        a0(n0) = Split(a0(n0), "'")(0)
        's0 = Join(a0, "', '")

        Debug.Print("")

        sqlSelAiPdParts01fromTextV01 = "-- " _
        & vbCrLf & "select" _
            & vbCrLf & vbTab & Join({
                "bm.Product", "bm.Item",
                "PartDesc.Type pType", "PartDesc.Family pFam",
                "it.Type iType", "it.Family iFam" _
            , ", " _
    & vbCrLf & "from" _
        & vbCrLf & vbTab & "vgIcoBillOfMaterials bm Inner Join" _
        & vbCrLf & vbTab & "vgMfiItems PartDesc On bm.Product = PartDesc.Item" _
        & vbCrLf & vbTab & "Left Join vgMfiItems it On bm.Item = it.Item" _
    & vbCrLf & "where" _
        & vbCrLf & vbTab & "bm.Product in " _
            & txList _
        & "" _
    & "" _
    & ""})
        '
        '(' Join(a0, "', '") ')
        '& vbCrLf & "with" _
        '& vbCrLf & vbTab & "ls as " & txList _
        '
    End Function

    Public Function sqlSelTestFromText(
    txList As String
) As String
        '
        '
        '

        sqlSelTestFromText _
    = sqlSelAiPdParts01fromTextV01(txList)
    End Function

    Public Function sqlSelAiPurch01fromText(
    txList As String
) As String
        '
        '
        '

        sqlSelAiPurch01fromText _
    = sqlSelAiPurch01fromTextV02(txList)
    End Function

    Public Function sqlSelAiPdParts01fromText(
    txList As String
) As String
        '
        '
        '

        sqlSelAiPdParts01fromText _
    = sqlSelAiPdParts01fromTextV01(txList)
    End Function

    Public Function sqlSelAiPurch01fromDict(
    dc As Scripting.Dictionary
) As String
        '
        '
        '
        sqlSelAiPurch01fromDict _
    = sqlSelAiPurch01fromText(Join(Split(
        sqlValSelFromDict(dc),
        vbCrLf), vbCrLf & vbTab
    ))
    End Function

    Public Function sqlSelAiPdParts01fromDict(
    dc As Scripting.Dictionary
) As String
        '
        '
        '
        sqlSelAiPdParts01fromDict _
    = sqlSelAiPdParts01fromText(
        sqlListValues(dc)
    )
    End Function

    Public Function sqlSelAiPurch01fromAssy(
    AiDoc As Inventor.Document
) As String
        '
        '
        '
        sqlSelAiPurch01fromAssy _
    = sqlSelAiPurch01fromDict(
        dcRemapByPtNum(
        dcAiDocComponents(AiDoc)
        )
    )
    End Function

    Public Function sqlSelAiPdParts01fromAssy(
    AiDoc As Inventor.Document
) As String
        '
        '
        '
        sqlSelAiPdParts01fromAssy _
    = sqlSelAiPdParts01fromDict(
        dcRemapByPtNum(
        dcAiDocComponents(AiDoc)
        )
    )
    End Function

End Module