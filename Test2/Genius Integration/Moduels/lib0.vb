Module lib0

    ' Measurement Unit Conversion Factors
    Public Const cvArSqCm2SqFt As Double = 0.00107639
    ' 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
    '
    ' /  1ft | 1in    \2     2                2
    '( ------+-------- ) * cm  = 0.00107639 ft
    ' \ 12in | 2.54cm /
    Public Const cvMassKg2LbM As Double = 2.20462
    Public Const cvLenIn2cm As Double = 2.54

    '
    Public Const guidRegPart As String = "{4D29B490-49B2-11D0-93C3-7E0706000000}"
    Public Const guidSheetMetal As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Public Const guidDesignAccl As String = "{BB8FE430-83BF-418D-8DF9-9B323D3DB9B9}"
    Public Const guidPipingSgmt As String = "{4D39D5F1-0985-4783-AA5A-FC16C288418C}"
    Public Const guidILogicAdIn As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
    '
    Public Const guidRegAssy As String = "{E60F81E1-49B3-11D0-93C3-7E0706000000}"
    Public Const guidWeldment As String = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}"

    '
    Public Const guidPrSumm As String = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}" 'Summary Information (Inventor Summary Information)
    Public Const guidPrDocu As String = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}" 'Document Summary Information (Inventor Document Summary Information)
    Public Const guidPrTrkg As String = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}" 'Design Tracking Properties (Design Tracking Properties)
    Public Const guidPrUser As String = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" 'User Defined Properties (Inventor User Defined Properties)
    Public Const guidPrCLib As String = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}" 'Content Library Component Properties (Content Library Component Properties)
    Public Const guidPrCCtr As String = "{CEAAEE65-91D8-444E-ACBA-BE54A5FB9D4D}" 'ContentCenter (ContentCenter)
    'Public Const guidPr____  As String = "{00000000-0000-0000-0000-000000000000}" 'Display Name (Name)
    '

    Public Const gnDesign As String = "Design Tracking Properties"
    Public Const pnMaterial As String = "Material"          '
    Public Const pnPartNum As String = "Part Number"       '
    Public Const pnStockNum As String = "Stock Number"      '
    Public Const pnFamily As String = "Cost Center"       '
    Public Const pnDesc As String = "Description"       '
    Public Const pnCatWebLink As String = "Catalog Web Link"  '

    Public Const gnCustom As String = "Inventor User Defined Properties"
    Public Const pnMass As String = "GeniusMass"    '
    Public Const pnRawMaterial As String = "RM"            '
    Public Const pnRmQty As String = "RMQTY"         '
    Public Const pnRmUnit As String = "RMUNIT"        '(replaces "RMUOM")
    '                                                       '
    Public Const pnArea As String = "Extent_Area"   '
    Public Const pnLength As String = "Extent_Length" '
    Public Const pnWidth As String = "Extent_Width"  '
    '
    Public Const pnThickness As String = "Thickness"     '
    '

    Public Function VbProjectLocal() As VBIDE.VBProject
        Dim Doc As Inventor.Document
        VbProjectLocal = Doc.VBAProject.VBProject
    End Function

    Public Function CnGnsDoyle() As ADODB.Connection
        Dim rt As ADODB.Connection
        ' NOTE[2021.12.08]:
        '     Might consider make rt a Static Object.
        '     If it can be created and opened just once
        '     during a run, this could potentially save
        '     a LOT of overhead from repeated open/close
        '     operations, and might save a little load
        '     on the server, as well.

        rt = New ADODB.Connection
        With rt
            .Provider = "SQLOLEDB" '"SQLNCLI11"
            .CursorLocation = CursorLocationEnum.adUseClient
            .Open("Data Source=DOYLE-ERP02", "GeniusReporting", "geniusreporting")
            .DefaultDatabase = "DoyleDB"
            '.Close
        End With
        CnGnsDoyle = rt
    End Function

    Public Function DcIvObjTypeEnum() As Scripting.Dictionary
        Dim dc As Scripting.Dictionary
        Dim en As Inventor.ObjectTypeEnum

        dc = New Scripting.Dictionary

        With dc
            en = ObjectTypeEnum.k3dAViewObject
            en = ObjectTypeEnum.kAliasFreeformFeatureObject
            en = ObjectTypeEnum.kAliasFreeformFeatureProxyObject
            en = ObjectTypeEnum.kAliasFreeformFeaturesObject
            en = ObjectTypeEnum.kAnalysisManagerObject
            en = ObjectTypeEnum.kAnalyticEdgeWorkAxisDefObject
            en = ObjectTypeEnum.kAngleConstraintObject
            en = ObjectTypeEnum.kAngleConstraintProxyObject
            en = ObjectTypeEnum.kAngleExtentObject
            en = ObjectTypeEnum.kAngleiMateDefinitionObject
            '.Add(kUnknownDocumentObject, "kUnknownDocumentObject"
            '.Add(kSATFileDocumentObject, "kSATFileDocumentObject"
            '.Add(kPresentationDocumentObject, "kPresentationDocumentObject"
            '.Add(kPartDocumentObject, "kPartDocumentObject"
            '.Add(kNoDocument, "kNoDocument"
            '.Add(kForeignModelDocumentObject, "kForeignModelDocumentObject"
            '.Add(kDrawingDocumentObject, "kDrawingDocumentObject"
            '.Add(kDesignElementDocumentObject, "kDesignElementDocumentObject"
            '.Add(kAssemblyDocumentObject, "kAssemblyDocumentObject"
        End With

        DcIvObjTypeEnum = dc
    End Function
    '
    '=====

    Public Function DcIvDocTypeEnum() As Scripting.Dictionary
        Dim dc As Scripting.Dictionary
        Dim en As Inventor.DocumentTypeEnum

        dc = New Scripting.Dictionary

        With dc
            .Add(DocumentTypeEnum.kUnknownDocumentObject, "kUnknownDocumentObject")
            .Add(DocumentTypeEnum.kSATFileDocumentObject, "kSATFileDocumentObject")
            .Add(DocumentTypeEnum.kPresentationDocumentObject, "kPresentationDocumentObject")
            .Add(DocumentTypeEnum.kPartDocumentObject, "kPartDocumentObject")
            .Add(DocumentTypeEnum.kNoDocument, "kNoDocument")
            .Add(DocumentTypeEnum.kForeignModelDocumentObject, "kForeignModelDocumentObject")
            .Add(DocumentTypeEnum.kDrawingDocumentObject, "kDrawingDocumentObject")
            .Add(DocumentTypeEnum.kDesignElementDocumentObject, "kDesignElementDocumentObject")
            .Add(DocumentTypeEnum.kAssemblyDocumentObject, "kAssemblyDocumentObject")
        End With

        DcIvDocTypeEnum = dc
    End Function
    '
    '=====

    Public Function TxDumpLs(ls As Object,
    Optional bk As String = vbCrLf
    ) As String
        Dim rt() As String
        Dim mx As Long
        Dim bs As Long
        Dim dx As Long

        If IsArray(ls) Then
            bs = LBound(ls)
            mx = UBound(ls)
            If bs > mx Then
                TxDumpLs = ""
                Exit Function
            Else
                ReDim rt(0 To mx - bs)
                For dx = bs To mx
                    rt(dx - bs) = TxDumpLs(ls(dx))
                Next
                TxDumpLs = Join(rt, bk)
                Exit Function
            End If
        ElseIf TypeOf ls Is Object Then
            If TypeOf ls Is Scripting.Dictionary Then
                TxDumpLs = TxDumpLs(ls.Keys)
                Exit Function
            End If
        End If
        TxDumpLs = CStr(ls)
    End Function

    Public Sub LsDump(ls As Object,
    Optional bk As String = vbCrLf
)
        Debug.Print(TxDumpLs(ls, bk))
    End Sub

    '
    ' The following is copied over from the Excel project file libExt.xlsm
    ' to provide a means of dumping Key-Value pairs from a Dictionary.
    '

    Public Function DumpLsKeyVal(dc As Scripting.Dictionary _
    , Optional dlmField As String = "," _
    , Optional dlmLine As String = vbCrLf _
    , Optional nullTxt As String = "<null>" _
    , Optional emptyTx As String = "<empty>"
) As String
        Dim d2 As Scripting.Dictionary
        Dim ky As Object
        Dim vl As Object
        'Dim rt As String

        'rt = ""
        If dc Is Nothing Then
            DumpLsKeyVal = ""
        Else
            d2 = New Scripting.Dictionary
            With dc
                For Each ky In .Keys
                    'rt = rt & ky & "," & .Item(ky) & vbCrLf

                    vl = { .Item(ky)}
                    ''  Any values which have
                    ''  no direct String conversion
                    ''  are replaced with String defaults
                    ''  (   user supplied, or see
                    ''      Function declaration above )
                    If (vl(0)) Is Nothing Then vl = nullTxt
                    If (vl(0)) Is Nothing Then vl = emptyTx
                    'If IsMissing(vl) Then vl = ""
                    'If IsError(vl) Then vl = ""
                    If TypeOf (vl(0)) Is Object Then
                        If vl(0) Is Nothing Then
                            vl = "<ob:Nothing>"
                        Else
                            vl = "<ob:" & TypeName(vl(0)) & ">"
                        End If
                    End If
                    If TypeOf vl Is Array Then
                        If TypeOf {vl(0)} Is Array Then
                            vl = "<array>"
                        Else
                            vl = vl(0)
                        End If
                    End If

                    d2.Add(Join({ky, vl}, dlmField), ky)
                    ''  Note that Key value might also
                    ''  require String default replacement,
                    ''  as well. Won't address this unless
                    ''  and until it becomes an issue.
                    ''
                    ''  If it DOES, the defaulting process
                    ''  will probably be broken out
                    ''  into its own function
                Next
            End With
            DumpLsKeyVal = Join(d2.Keys, dlmLine)
        End If
    End Function

End Module