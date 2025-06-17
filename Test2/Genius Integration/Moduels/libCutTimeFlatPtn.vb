Module libCutTimeFlatPtn
    Public Sub showPerimeterInches(ThisApplication As Inventor.Application, oDoc As Inventor.Document)
        Dim rt As Double
        rt = fpPerimeter(ThisApplication, oDoc)

        If rt > 0 Then
            'MsgBox("Total length of all loops on face: " & rt & " cm"
            MsgBox("Total length of all loops on face: " & (rt / CvLenIn2cm) & " in")
        ElseIf rt < 0 Then
            MsgBox("No Valid Flat Pattern Found")
        Else
            MsgBox("Flat Pattern Has No Measurable Perimeter")
        End If
    End Sub

    Public Function fpPerimeter(ThisApplication As Inventor.Application,
    oDoc As Inventor.Document,
    Optional ld As Double = 0
    ) As Double
        Dim oFace As Inventor.Face
        Dim oEdge As Inventor.Edge
        Dim rt As Double
        Dim ct As Double

        ct = 0
        If TypeOf oDoc Is Inventor.PartDocument Then
            oFace = smPartFlatPatternTopFace(ThisApplication, oDoc)
            If oFace Is Nothing Then 'we can't get a perimeter
                '' Simple 'error' indicator
                rt = -1
            Else
                rt = 0
                For Each oEdge In oFace.Edges
                    rt = rt + edgeLength(oEdge)
                    ct = 1 + ct
                Next
            End If
        Else
            rt = -1
        End If
        fpPerimeter = rt '+ ct * ld
        ' Holding off on "lead-in" calculation
        ' ct does NOT count closed loops,
        ' but individual edges, resulting in
        ' a MUCH larger count than appropriate!
        '
        ' ld is optional "lead-in" from piercing point
        ' to cutting line. A non-zero value here
        ' adds this lead-in length for every hole,
        ' and may produce more accurate estimates
        ' of cutting length as reported by SigmaNest
    End Function
    Public Function fpPerimeterInch(ThisApplication As Inventor.Application,
    oDoc As Inventor.Document,
    Optional ld As Double = 0
) As Double
        Dim rt As Double

        rt = fpPerimeter(ThisApplication, oDoc) ', ld / cvLenIn2cm)
        If rt > 0 Then
            fpPerimeterInch = rt / CvLenIn2cm
        Else
            fpPerimeterInch = rt
        End If
    End Function
    Public Function edgeLength(ed As Inventor.Edge) As Double
        Dim mn As Double
        Dim mx As Double
        Dim lg As Double

        With ed.Evaluator
            .GetParamExtents(mn, mx)
            .GetLengthAtParam(mn, mx, lg)
        End With
        edgeLength = lg
    End Function

    Public Function smPartFlatPatternTopFace(oDoc As PartDocument, Thisapplication As Inventor.Application) As Inventor.Face
        If oDoc Is Nothing Then
            smPartFlatPatternTopFace = Nothing
        Else
            smPartFlatPatternTopFace = fpTopFaceIfShtMetal(Thisapplication, oDoc.ComponentDefinition)
        End If
    End Function

    Public Function fpTopFaceIfShtMetal(ThisApplication As Inventor.Application,
    oDef As Inventor.ComponentDefinition
) As Inventor.Face
        If TypeOf oDef Is Inventor.SheetMetalComponentDefinition Then
            fpTopFaceIfShtMetal = smcdFlatPtnTopFace(ThisApplication, oDef)
        Else
            fpTopFaceIfShtMetal = Nothing
        End If
    End Function

    Public Function smcdFlatPtnTopFace(ThisApplication As Inventor.Application,
    oDef As Inventor.SheetMetalComponentDefinition
) As Inventor.Face
        smcdFlatPtnTopFace = fpTopFace(ThisApplication, oDef.FlatPattern)
    End Function

    Public Function fpTopFace(fp As Inventor.FlatPattern, ThisApplication As Inventor.Application) As Inventor.Face
        Dim oZAxis As Inventor.UnitVector
        Dim oFace As Inventor.Face
        Dim rt As Inventor.Face

        oZAxis = ThisApplication.TransientGeometry.CreateUnitVector(0, 0, 1)

        For Each oFace In fp.Body.Faces
            ' Only looking until we find a match
            If rt Is Nothing Then
                With oFace
                    ' Only interested in planar faces
                    If .SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                        With aiPlane(.Geometry)
                            ' Only interested in faces that have z-direction normal
                            If .Normal.IsParallelTo(oZAxis) Then
                                ' Look for the face with Z = 0
                                If .RootPoint.Z <= 0.0000001 Then
                                    rt = oFace
                                End If
                            End If
                        End With
                    End If
                End With
            End If
        Next

        fpTopFace = rt
    End Function

End Module