Option Explicit On

Module lib0

    ' Measurement Unit Conversion Factors
    Public Const CvArSqCm2SqFt As Double = 0.00107639
    Public Const CvMassKg2LbM As Double = 2.20462
    Public Const CvLenIn2cm As Double = 2.54

    ' GUIDs for Inventor object types
    Public Const GuidRegPart As String = "{4D29B490-49B2-11D0-93C3-7E0706000000}"
    Public Const GuidSheetMetal As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Public Const GuidDesignAccl As String = "{BB8FE430-83BF-418D-8DF9-9B323D3DB9B9}"
    Public Const GuidPipingSgmt As String = "{4D39D5F1-0985-4783-AA5A-FC16C288418C}"
    Public Const GuidILogicAdIn As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
    Public Const GuidRegAssy As String = "{E60F81E1-49B3-11D0-93C3-7E0706000000}"
    Public Const GuidWeldment As String = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}"

    ' GUIDs for Inventor property sets
    Public Const GuidPrSumm As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}" ' Summary Information
    Public Const GuidPrDocu As String = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}" ' Document Summary Information
    Public Const GuidPrTrkg As String = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}" ' Design Tracking Properties
    Public Const GuidPrUser As String = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" ' User Defined Properties
    Public Const GuidPrCLib As String = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}" ' Content Library Component Properties
    Public Const GuidPrCCtr As String = "{CEAAEE65-91D8-444E-ACBA-BE54A5FB9D4D}" ' ContentCenter

    ' Property Names
    Public Const GnDesign As String = "Design Tracking Properties"
    Public Const PnMaterial As String = "Material"
    Public Const PnPartNum As String = "Part Number"
    Public Const PnStockNum As String = "Stock Number"
    Public Const PnFamily As String = "Cost Center"
    Public Const PnDesc As String = "Description"
    Public Const PnCatWebLink As String = "Catalog Web Link"

    Public Const GnCustom As String = "Inventor User Defined Properties"
    Public Const PnMass As String = "GeniusMass"
    Public Const PnRawMaterial As String = "RM"
    Public Const PnRmQty As String = "RMQTY"
    Public Const PnRmUnit As String = "RMUNIT"
    Public Const PnArea As String = "Extent_Area"
    Public Const PnLength As String = "Extent_Length"
    Public Const PnWidth As String = "Extent_Width"
    Public Const PnThickness As String = "Thickness"

    ''' <summary>
    ''' Returns the local VBA project for the given Inventor document.
    ''' </summary>
    Public Function VbProjectLocal(doc As Inventor.Document) As Object
        If doc Is Nothing Then Return Nothing
        Return doc.VBAProject.VBProject
    End Function

    ''' <summary>
    ''' Returns a new, open ADODB connection to the Genius Doyle database.
    ''' </summary>
    Public Function CnGnsDoyle() As Object
        Dim conn As Object
        conn = CreateObject("ADODB.Connection")
        With conn
            .Provider = "SQLOLEDB"
            .CursorLocation = 3 ' adUseClient
            .Open("Data Source=DOYLE-ERP02", "GeniusReporting", "geniusreporting")
            .DefaultDatabase = "DoyleDB"
        End With
        Return conn
    End Function

    ''' <summary>
    ''' Returns a dictionary mapping Inventor DocumentTypeEnum values to their names.
    ''' </summary>
    Public Function DcIvDocTypeEnum() As Object
        Dim dc As Object
        dc = CreateObject("Scripting.Dictionary")
        dc.Add(12289, "kUnknownDocumentObject")
        dc.Add(12290, "kSATFileDocumentObject")
        dc.Add(12291, "kPresentationDocumentObject")
        dc.Add(12292, "kPartDocumentObject")
        dc.Add(12293, "kNoDocument")
        dc.Add(12294, "kForeignModelDocumentObject")
        dc.Add(12295, "kDrawingDocumentObject")
        dc.Add(12296, "kDesignElementDocumentObject")
        dc.Add(12297, "kAssemblyDocumentObject")
        Return dc
    End Function

    ''' <summary>
    ''' Dumps the contents of a list or dictionary as a string.
    ''' </summary>
    Public Function TxDumpLs(ls As Object, Optional bk As String = vbCrLf) As String
        If IsArray(ls) Then
            Dim bs As Long, mx As Long, dx As Long
            bs = LBound(ls)
            mx = UBound(ls)
            If bs > mx Then Return ""
            Dim rt() As String
            ReDim rt(mx - bs)
            For dx = bs To mx
                rt(dx - bs) = TxDumpLs(ls(dx))
            Next
            Return Join(rt, bk)
        ElseIf TypeName(ls) = "Dictionary" Then
            Return TxDumpLs(ls.Keys)
        Else
            Return CStr(ls)
        End If
    End Function

    ''' <summary>
    ''' Prints the contents of a list or dictionary to the debug window.
    ''' </summary>
    Public Sub LsDump(ls As Object, Optional bk As String = vbCrLf)
        Debug.Print(TxDumpLs(ls, bk))
    End Sub

    ''' <summary>
    ''' Dumps key-value pairs from a dictionary as a delimited string.
    ''' </summary>
    Public Function DumpLsKeyVal(dc As Object,
                                Optional dlmField As String = ",",
                                Optional dlmLine As String = vbCrLf,
                                Optional nullTxt As String = "<null>",
                                Optional emptyTx As String = "<empty>") As String
        If dc Is Nothing Then Return ""
        Dim d2 As Object
        d2 = CreateObject("Scripting.Dictionary")
        Dim ky As Object
        For Each ky In dc.Keys
            Dim vl As Object
            vl = dc.Item(ky)
            Dim valStr As String
            If vl Is Nothing Then
                valStr = nullTxt
            ElseIf TypeName(vl) = "String" And vl = "" Then
                valStr = emptyTx
            ElseIf TypeOf vl Is Object Then
                valStr = "<ob:" & TypeName(vl) & ">"
            Else
                valStr = CStr(vl)
            End If
            d2.Add(CStr(ky) & dlmField & valStr, ky)
        Next
        DumpLsKeyVal = Join(d2.Keys, dlmLine)
    End Function

    ''' <summary>
    ''' Returns a string consisting of a repeated pattern.
    ''' </summary>
    Public Function Repeat(Count As Long, Text As String) As String
        If Count <= 0 Or Len(Text) = 0 Then Return ""
        Repeat = New String(CChar(Left(Text, 1)), Count)
        If Len(Text) > 1 Then Repeat = Replace(Repeat, Left(Text, 1), Text)
    End Function

    ''' <summary>
    ''' Returns a block of text with repeated lines and characters.
    ''' </summary>
    Public Function TxBlk(Lines As Long, Chars As Long, Optional Use As String = "+") As String
        If Lines <= 0 Or Chars <= 0 Then Return ""
        Dim lineStr As String
        lineStr = New String(CChar(Use), Chars)
        Dim arr() As String
        ReDim arr(Lines - 1)
        Dim i As Long
        For i = 0 To Lines - 1
            arr(i) = lineStr
        Next
        TxBlk = Join(arr, vbCrLf)
    End Function

    ''' <summary>
    ''' Sets a document's BOM structure to Purchased, based on type.
    ''' </summary>
    Public Function MkAiDocPurchased(AiDoc As Object) As VbMsgBoxResult
        If TypeName(AiDoc) = "PartDocument" Then
            MkAiDocPurchased = MkAiPartPurchased(AiDoc)
        ElseIf TypeName(AiDoc) = "AssemblyDocument" Then
            MkAiDocPurchased = MkAiAssyPurchased(AiDoc)
        Else
            MkAiDocPurchased = vbNo
        End If
    End Function

    ''' <summary>
    ''' Sets a PartDocument's BOM structure to Purchased.
    ''' </summary>
    Public Function MkAiPartPurchased(AiDoc As Object) As VbMsgBoxResult
        If AiDoc Is Nothing Then
            MkAiPartPurchased = vbNo
            Exit Function
        End If
        On Error Resume Next
        AiDoc.ComponentDefinition.BOMStructure = 51970 ' BOMStructureEnum.kPurchasedBOMStructure
        If Err.Number = 0 Then
            MkAiPartPurchased = vbOK
        Else
            MkAiPartPurchased = vbAbort
        End If
        On Error GoTo 0
    End Function

    ''' <summary>
    ''' Sets an AssemblyDocument's BOM structure to Purchased.
    ''' </summary>
    Public Function MkAiAssyPurchased(AiDoc As Object) As VbMsgBoxResult
        If AiDoc Is Nothing Then
            MkAiAssyPurchased = vbNo
            Exit Function
        End If
        On Error Resume Next
        AiDoc.ComponentDefinition.BOMStructure = 51970 ' BOMStructureEnum.kPurchasedBOMStructure
        If Err.Number = 0 Then
            MkAiAssyPurchased = vbOK
        Else
            MkAiAssyPurchased = vbAbort
        End If
        On Error GoTo 0
    End Function

    ''' <summary>
    ''' Returns a Scripting.Dictionary, optionally using an existing one.
    ''' </summary>
    Public Function DcTemplate0A(Optional dc As Object = Nothing) As Object
        If dc Is Nothing Then
            DcTemplate0A = CreateObject("Scripting.Dictionary")
        Else
            DcTemplate0A = dc
        End If
    End Function

    ''' <summary>
    ''' Copies text to clipboard (OBSOLETE).
    ''' </summary>
    Public Function Send2ClipBd_OBSOLETE(src As Object) As Object
        With CreateObject("MSForms.DataObject")
            .SetText(
            .PutInClipboard)
        End With
        Send2ClipBd_OBSOLETE = src
    End Function

End Module