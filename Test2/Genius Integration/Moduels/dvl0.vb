Module dvl0
    Dim ThisApplication As Inventor.Application
    Public Function d0g0f0(
    cd As Inventor.SheetMetalComponentDefinition,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        'New Sheet Metal Part processing function
        ' Reference function d0g0f0
        ''
        Dim rt As Scripting.Dictionary
        ''
        Dim PartDoc As Inventor.PartDocument
        Dim PropertySet As Inventor.Property
        Dim ThicknessProp As Inventor.Parameter
        ''
        Dim ck As VbMsgBoxResult
        Dim ec As Long
        Dim ed As String
        'Dim op As Boolean
        Dim Veiw1 As Inventor.View
        Dim Veiw2 As Inventor.View
        ''
        Dim dLength As Double
        Dim dWidth As Double
        Dim dArea As Double
        Dim strWidth As String
        Dim strLength As String
        Dim strArea As String
        ''
        Dim dHeight As Double
        Dim dfHtThk As Double
        Dim strDVNs As String
        ''

        If dc Is Nothing Then
            d0g0f0 = d0g0f0(cd,
        New Scripting.Dictionary)
        Else
            rt = dc
            If cd Is Nothing Then
                'Stop
            Else
                PartDoc = cd.Document
                PropertySet = PartDoc.Propertys.Item(gnCustom)
                Veiw1 = d0g1f0(PartDoc)
                'op = PartDoc.Open

                On Error Resume Next
                Err.Clear()
                ThicknessProp = cd.Thickness
                ec = Err.Number
                ed = Err.Description
                On Error GoTo 0

                If ec = 0 Then 'we're good so far
                    If Not cd.HasFlatPattern Then
                        ck = newFmTest2().AskAbout(PartDoc, ,
                        "NO FLAT PATTERN!" & vbCrLf &
                        "Try to generate one?"
                    )
                        If ck = vbYes Then
                            'If Not op Then Stop
                            'Want to see if forcing an Unfold
                            'causes an unopened document to open.
                            'Also want to see how Open Property
                            'relates to a referenced Document
                            'not (yet) separately opened.

                            Err.Clear()
                            cd.Unfold()

                            If Err.Number = 0 Then
                                If cd.HasFlatPattern Then
                                    cd.FlatPattern.ExitEdit()
                                End If
                            Else
                                Stop 'Couldn't make Flat Pattern
                            End If
                            Err.Clear()

                            Veiw2 = d0g1f0(PartDoc)
                            If Not Veiw2 Is Nothing Then
                                If Veiw1 Is Nothing Then Veiw2.Close()
                            End If
                        Else
                        End If
                    End If

                    If cd.HasFlatPattern Then
                        With cd.FlatPattern
                            'First, make sure it's VALID
                            With .Body.RangeBox
                                ' Check height against thickness
                                ' Valid flat pattern should return
                                ' zero or VERY minimal difference
                                dHeight = .MaxPoint.Z - .MinPoint.Z
                                dfHtThk = Math.Abs(dHeight - ThicknessProp.Text)

                                ' Get the extent of the face.
                                ' Extract the width, length and area from the range.
                                dLength = .MaxPoint.X - .MinPoint.X
                                dWidth = .MaxPoint.Y - .MinPoint.Y
                                dArea = dLength * dWidth
                            End With
                            'Stop
                            '
                            ' At this point, we should have enough
                            ' to check at least a few things,
                            ' and possibly pick out stock.
                            '
                            If dfHtThk > 0.01 Then
                                'Stop 'and prep for machined (non sheet metal) specs
                                ' Pretty sure dimension values
                                ' come through in centimeters
                                ' so try converting them here
                                'sort3dimsUp
                                d0g1f4(cd)

                                rt = d0g1f3(PartDoc,
                                sort3dimsUp(
                                    dHeight / cvLenIn2cm,
                                    dWidth / cvLenIn2cm,
                                    dLength / cvLenIn2cm
                                ),
                            rt)
                            Else
                                'Stop 'and prep to verify sheet metal processing
                            End If

                            If dArea > 0 Then 'this one's a longshot, BUT!
                                ' an invalid flat pattern SHOULD have no geometry,
                                ' which means it SHOULD have no area to speak of.
                                ' '
                                ' One would think this obvious, in retrospect,
                                ' but one would not be surprised to be proven wrong.
                                ' Again.

                                With PartDoc

                                    ' Convert values into document units.
                                    ' This will result in strings that are identical
                                    ' to the strings shown in the Extent dialog.
                                    With .UnitsOfMeasure
                                        strWidth = .GetStringFromValue(dWidth,
                                        .GetStringFromType(.LengthUnits))
                                        strLength = .GetStringFromValue(dLength,
                                        .GetStringFromType(.LengthUnits))
                                        strArea = .GetStringFromValue(dArea,
                                        .GetStringFromType(.LengthUnits) & "^2")

                                        If dfHtThk > 0.01 Then
                                            strDVNs = .GetStringFromValue(
                                            dfHtThk, .GetStringFromType(.LengthUnits))
                                            'Debug.Print(Join({"OFFTHK", _
                                            aiDocument(.Document.FullFileName(Format(dHeight, "0.0000") & Format(.prThickness.Text, "0.0000") & Format(dHeight - .prThickness.Text, "0.0000")))
                                            'Stop
                                        Else
                                            strDVNs = ""
                                        End If
                                    End With
                                End With
                            Else 'we don't have a valid FlatPattern
                                If MsgBox(Join({
                                "The flat pattern for this",
                                "part has no features,",
                                "and is likely not valid.",
                                "",
                                "Pause here to review?",
                                "(Click 'NO' to just keep going)"
                            }, vbCrLf), vbYesNo,
                                "Invalid Flat Pattern"
                            ) = vbYes Then
                                    Stop 'and let the user look into it
                                End If
                                Debug.Print(aiDocument(.Document).FullDocumentName)
                            End If
                        End With

                        ' Add area to custom property 
                        '                         rt = dcWithProp(aiProp, pnRmQty, dArea * cvArSqCm2SqFt, rt)

                        ' Add Width to custom property 
                        '                         rt = dcWithProp(aiProp, pnWidth, strWidth, rt)

                        ' Add Length to custom property 
                        '                         rt = dcWithProp(aiProp, pnLength, strLength, rt)

                        ' Add AreaDescription to custom property 
                        '                         rt = dcWithProp(aiProp, pnArea, strArea, rt)

                        If Len(strDVNs) > 0 Then
                            '                             rt = dcWithProp(aiProp, "OFFTHK", strWidth, rt)
                        End If
                    Else
                    End If
                Else
                    Stop
                End If
                On Error GoTo 0
            End If
            d0g0f0 = rt
        End If

        '
        '
        '
    End Function
    'For Each dc In aiDocAssy(aiDocActive).ComponentDefinition.Occurrences: Debug.Print(aiDocument(aiCompOcc(obOf(dc)).Definition.Document).Open, aiDocument(aiCompOcc(obOf(dc)).Definition.Document).FullDocumentName: Next
    'Looks like Open property will NOT distinguish documents in tab list from those not
    'All entries came up True

    Public Function d0g1f0(rf As Inventor.Document) As Inventor.View
        Dim rt As Inventor.View
        Dim vw As Inventor.View

        rt = Nothing
        For Each vw In ThisApplication.Views
            If vw.Document Is rf Then rt = vw
        Next
        d0g1f0 = rt
    End Function

    Public Function d0g1f1(
    rf As Inventor.PartDocument
) As Inventor.SheetMetalComponentDefinition
        If rf Is Nothing Then
            d0g1f1 = Nothing
        Else
            d0g1f1 = aiCompDefShtMetal(
            rf.ComponentDefinition
        )
        End If
    End Function

    Public Function noVal(Optional ByVal vt As VbVarType = vbEmpty) As Object
        If vt And vbArray Then
            noVal = {}
        Else
            Select Case vt
                Case vbString : noVal = ""
                Case vbLong : noVal = CLng(0)
                Case vbVariant : noVal = vbEmpty

                Case vbInteger : noVal = CInt(0)
                Case vbSingle : noVal = CSng(0)
                Case vbDouble : noVal = CDbl(0)
                Case vbDecimal : noVal = CDec(0)
                Case vbCurrency : noVal = CType(0, Decimal)
                Case vbBoolean : noVal = CBool(0)
                Case vbByte : noVal = CByte(0)

                Case vbEmpty : noVal = vbEmpty
                Case vbNull : noVal = vbNull

                Case vbObject : noVal = Nothing

                Case vbDate : Stop 'noVal = Empty
                Case VbVarType.vbError : Stop 'noVal = Empty
                Case VbVarType.vbDataObject : Stop 'noVal = Empty
                Case vbUserDefinedType : Stop 'noVal = Empty
            End Select
        End If
    End Function

    Public Function pt3d(
    Optional d0 As Double = 0#,
    Optional d1 As Double = 0#,
    Optional d2 As Double = 0#
) As Double()
        Dim rt(2) As Double

        rt(0) = d0
        rt(1) = d1
        rt(2) = d2

        pt3d = rt
    End Function

    Public Function sort3dimsUp(
    d0 As Double, d1 As Double, d2 As Double
) As Double()
        Dim rt() As Double

        If d1 < d0 Then
            rt = sort3dimsUp(d1, d0, d2)
        ElseIf d2 < d1 Then
            rt = sort3dimsUp(d2, d0, d1)
        Else
            ReDim rt(2)
            rt(0) = d0
            rt(1) = d1
            rt(2) = d2
        End If
        sort3dimsUp = rt
    End Function

    Public Function sort3dimsDn(
    d0 As Double, d1 As Double, d2 As Double
) As Double()
        Dim rt() As Double

        If d1 > d0 Then
            rt = sort3dimsDn(d1, d0, d2)
        ElseIf d2 > d1 Then
            rt = sort3dimsDn(d2, d0, d1)
        Else
            ReDim rt(2)
            rt(0) = d0
            rt(1) = d1
            rt(2) = d2
        End If
        sort3dimsDn = rt
    End Function

    Public Function aiBoxDims(
    RefBox As Inventor.Box
) As Double()
        Dim rt() As Double
        Dim mx As Inventor.Point
        Dim mn As Inventor.Point

        With RefBox
            mx = .MaxPoint
            mn = .MinPoint
        End With

        rt(0) = mx.X - mn.X
        rt(1) = mx.Y - mn.Y
        rt(2) = mx.Z - mn.Z

        aiBoxDims = rt
    End Function

    Public Function aiBoxSortDown(
    RefBox As Inventor.Box
) As Inventor.Box
        Dim rt As Inventor.Box
        Dim mx As Inventor.Point
        Dim mn As Inventor.Point

        With ThisApplication.TransientGeometry
            rt = .CreateBox()
            With RefBox
                mx = .MaxPoint
                mn = .MinPoint
                rt.PutBoxData(pt3d(),
            sort3dimsDn(
                mx.X - mn.X,
                mx.Y - mn.Y,
                mx.Z - mn.Z
            ))
            End With
        End With

        aiBoxSortDown = rt
    End Function

    Public Function dcSteelType2Spec6() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary

        rt = New Scripting.Dictionary
        With rt
            .Add("Steel, Mild", "MS")
            .Add("Stainless Steel", "SS")
            .Add("Stainless Steel, Austenitic", "SS")
        End With

        dcSteelType2Spec6 = rt
    End Function

    Public Function steelSpec6(
    stl As String, Optional Ask As Long = 0
) As String
        Dim rt As String
        'With dcSteelType2Spec6()
        '    If .Exists(stl) Then
        '        steelSpec6 = .Item(stl)
        '    Else
        '        steelSpec6 = ""
        '    End If
        'End With

        Select Case stl
            Case "Stainless Steel"
                rt = "SS"
            Case "Stainless Steel, Austenitic"
                rt = "SS"
            Case "Stainless Steel 304"
                rt = "SS"
            Case "Steel, Mild"
                rt = "MS"
            Case "Rubber"
                rt = ""  'LG
            Case "Rubber, Silicone"
                rt = ""  'LG
            Case "UHMW, White"
                rt = ""  'LG
            Case Else
                If Ask Then
                    Debug.Print("=== UNKNOWN MATERIAL ===")
                    Debug.Print("   (" & stl & ")")
                    Debug.Print("Please supply a code for Specification 6,")
                    Debug.Print("if applicable, on the line below, and")
                    Debug.Print("press [ENTER] or [RETURN] to modify.")
                    Debug.Print("Press [F5] when ready to continue.")
                    Debug.Print("rt  = """" '<-( place code between double quotes )")
                    Stop
                Else
                    rt = ""
                End If
        End Select

        steelSpec6 = rt
    End Function

    Public Function d0g1f3(
    rf As Inventor.PartDocument, dm() As Double,
    Optional dc As Scripting.Dictionary = Nothing
        ) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        ''
        Dim aspect As Double
        Dim offSqr As Double
        Dim length As Double
        ''
        Dim rmType As String
        Dim rmSpc6 As String
        Dim rmItem As String
        Dim rmUnit As String
        Dim rmQty As String

        If dc Is Nothing Then
            rt = d0g1f3(rf, dm,
            New Scripting.Dictionary
        )
        Else
            rt = dc

            aspect = dm(1) / dm(0)
            offSqr = aspect - 1.0#
            length = dm(2)
            rmType = rf.Propertys(gnDesign).Item(pnMaterial).Text
            rmSpc6 = steelSpec6(rmType)

            'Debug.Print(".Add(""" & rmType & """, """""
            Debug.Print("Material: " _
            & rmType & " (" _
            & rmSpc6 & ")")
            '
            Debug.Print("Cross Section: " _
            & Format$(dm(1), "0.000") _
            & " X " _
            & Format$(dm(0), "0.000"))
            '
            Debug.Print("Length: " _
            & Format$(dm(2), "0.000"))
            '

            If offSqr < 0.01 Then
                Debug.Print("Likely Square or System.Math.Round")
                'Stop 'probably square or round
            ElseIf offSqr > 20 Then
                Debug.Print("Likely Sheet or Plate")
                'Stop 'might be SOME sort of sheet/plate
            Else
                Debug.Print("Likely Rectangular, Uneven?")
                'Stop 'likely rectangular
            End If

            rmItem = rmType & "???" 'User has to help from info given
            rmQty = Format$(dm(2), "0.000")
            rmUnit = "IN"

            dc.Add("RM", rmItem)
            dc.Add("RMQTY", rmQty)
            dc.Add("RMUNIT", rmUnit)

            Stop
        End If
        d0g1f3 = rt
    End Function

    Public Function d0g1f4(
    cd As Inventor.SheetMetalComponentDefinition
) As Scripting.Dictionary
        Dim sb As Inventor.SurfaceBody
        Dim fc As Inventor.Face
        Dim PartDoc As Inventor.Point

        Dim rkMgr As Inventor.ReferenceKeyManager
        Dim kyContx As Long
        Dim kyBytes() As Byte
        Dim kyLabel As String

        Dim d0 As Scripting.Dictionary
        Dim d1 As Scripting.Dictionary
        Dim ky As Object
        Dim k1 As Object

        d0 = New Scripting.Dictionary
        d1 = New Scripting.Dictionary

        rkMgr = aiDocument(
        cd.Document
    ).ReferenceKeyManager

        kyContx = rkMgr.CreateKeyContext

        For Each sb In cd.SurfaceBodies
            For Each fc In sb.Faces
                fc.GetReferenceKey(kyBytes, kyContx)
                kyLabel = rkMgr.KeyToString(kyBytes)
                d1.Add(kyLabel, fc) '.InternalName

                With fc.Evaluator
                    Debug.Print(TypeName(fc.Geometry) &
                "(" & CStr(0 _
                    + IIf(.IsExtrudedShape, 1, 0) _
                    + IIf(.IsRevolvedShape, 2, 0)
                ) & ")")
                    With .RangeBox
                        d0.Add(d0.Count, .MaxPoint)
                        d0.Add(d0.Count, .MinPoint)
                    End With
                End With
                'If 0 Then
                'ElseIf fc.SurfaceType = kPlaneSurface Then
                'ElseIf fc.SurfaceType = kCylinderSurface Then
                'ElseIf fc.SurfaceType = kConeSurface Then
                'ElseIf fc.SurfaceType = kSphereSurface Then
                'ElseIf fc.SurfaceType = kTorusSurface Then
                'ElseIf fc.SurfaceType = kBSplineSurface Then
                'ElseIf fc.SurfaceType = kEllipticalCylinderSurface Then
                'ElseIf fc.SurfaceType = kEllipticalConeSurface Then
                'ElseIf fc.SurfaceType = kUnknownSurface Then
                'Else
                'End If
            Next

            For Each ky In d0.Keys
                PartDoc = d0.Item(ky)
                For Each k1 In d1.Keys 'fc In sb.Faces
                    fc = d1.Item(k1)
                    'If d1.Exists(k1) Then
                    If Not fc.Evaluator.RangeBox.Contains(PartDoc) Then
                        d1.Remove(k1) 'fc.InternalName
                    End If
                    'End If
                Next
                Stop
            Next
        Next
    End Function

    Public Function d0g1f5(
    sb As Inventor.SurfaceBody
) As Scripting.Dictionary
        Dim fc As Inventor.Face
        Dim PartDoc As Inventor.Point

        Dim dFc As Scripting.Dictionary
        Dim dPt As Scripting.Dictionary
        Dim kPt As Object

        dFc = New Scripting.Dictionary
        dPt = New Scripting.Dictionary

        For Each fc In sb.Faces
            dFc.Add(fc.InternalName, fc)
            With fc.Evaluator
                Debug.Print(TypeName(fc.Geometry) &
            "(" & CStr(0 _
                + IIf(.IsExtrudedShape, 1, 0) _
                + IIf(.IsRevolvedShape, 2, 0)
            ) & ")")
                With .RangeBox
                    dPt.Add(dPt.Count, .MaxPoint)
                    dPt.Add(dPt.Count, .MinPoint)
                End With
            End With
        Next

        For Each kPt In dPt.Keys
            PartDoc = dPt.Item(kPt)
            For Each fc In sb.Faces
                If dFc.Exists(fc.InternalName) Then
                    If Not fc.Evaluator.RangeBox.Contains(PartDoc) Then
                        dFc.Remove(fc.InternalName)
                        If dFc.Count = 0 Then Stop
                    End If
                End If
            Next
            Stop
        Next

        d0g1f5 = dFc
    End Function

    Public Function d0g1f6(
    fc As Inventor.Face
) As String
        Dim kyBytes() As Byte
        Dim kyContx As Long
        Dim rt As String

        With aiDocument(
        fc.SurfaceBody.ComponentDefinition.Document
    ).ReferenceKeyManager '.CreateKeyContext
            kyContx = .CreateKeyContext
            fc.GetReferenceKey(kyBytes, kyContx)
            rt = .KeyToString(kyBytes)
        End With
        d0g1f6 = rt
    End Function

    Public Function aiPoint(ob As Object) As Inventor.Point
        If TypeOf ob Is Inventor.Point Then
            aiPoint = ob
        Else
            aiPoint = Nothing
        End If
    End Function

    ' d0g2: Testing
    '
    '

    Public Function d0g2f1()
        ' Verify 3-way sorting function sort3dimsUp
        Dim ck() As Double
        ck = sort3dimsUp(2, 3, 5) : Stop
        ck = sort3dimsUp(2, 5, 3) : Stop
        ck = sort3dimsUp(3, 2, 5) : Stop
        ck = sort3dimsUp(3, 5, 2) : Stop
        ck = sort3dimsUp(5, 2, 3) : Stop
        ck = sort3dimsUp(5, 3, 2) : Stop
    End Function

    Public Function d0g2f2()
        ' Testing new spec pickup system
        Dim ky As Object

        With dcAiDocComponents(
        ThisApplication.ActiveDocument, , 0
    )
            For Each ky In .Keys
                Debug.Print(ky)
                .Item(ky) = d0g0f0(aiCompDefShtMetal(
                aiCompDefOf(aiDocPart(.Item(ky)))
            ))
                If .Item(ky) Is Nothing Then
                    .Remove(ky)
                Else
                    'Stop
                End If
            Next
        End With
    End Function

    Public Function d0g2f3()
        ' Checking some behaviors
        ' on string arrays
        ' vs Objects
        Dim ky As Object
        With New AiPropSetter
            Debug.Print(Join(.PropList(), "|"))
            For Each ky In .PropList()
                Debug.Print(ky)
            Next
        End With
    End Function

    Public Function d0g2f4(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' Return Dictionary of ALLOCATED Property
        ' Values (True/False) attached to all components
        ' and subcomponents of the active Document.
        '
        ' Where the ALLOCATED Property is not present,
        ' represent it as "<default>"
        '
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                'Debug.Print(ky
                With aiDocument(.Item(ky)).Propertys.Item(gnCustom)
                    On Error Resume Next
                    Err.Clear()
                    InvProperty = .Item("ALLOCATED")
                    If Err.Number = 0 Then
                        rt.Add(ky, CStr(InvProperty.Text) & "|" & ky)
                    Else
                        rt.Add(ky, "<default>" & "|" & ky)
                    End If
                    On Error GoTo 0
                End With
            Next
        End With
        d0g2f4 = rt
    End Function
    'Debug.Print(Join(d0g2f4(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbCrLf)

    Public Function d0g2f6(
    dc As Scripting.Dictionary, pn As String,
    Optional gn As String = gnCustom,
    Optional df As String = "<NOPROP>"
) As Scripting.Dictionary
        '
        ' Return Dictionary of named Property Values
        ' attached to all Inventor Documents
        ' in supplied Dictionary.
        '
        '
        ' Where the ALLOCATED Property is not present,
        ' represent it as "<default>"
        '
        Dim ad As Inventor.Document
        Dim rt As Scripting.Dictionary
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                ad = aiDocument(.Item(ky))
                If ad Is Nothing Then
                Else
                End If
                With ad.Propertys.Item(gn)
                    On Error Resume Next
                    Err.Clear()
                    InvProperty = .Item("ALLOCATED")
                    If Err.Number = 0 Then
                        rt.Add(ky, CStr(InvProperty.Text) & "|" & ky)
                    Else
                        rt.Add(ky, "<default>" & "|" & ky)
                    End If
                    On Error GoTo 0
                End With
            Next
        End With
        d0g2f6 = rt
    End Function
    'Debug.Print(Join(d0g2f6(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbCrLf)

    Public Function d0g2f5(dc As Scripting.Dictionary) As Scripting.Dictionary
        '
        ' Attempt to "transpose" contents of Dictionary
        ' and return a dictionary of Items mapped
        ' to sub-Dictionaries containing all keys
        ' which mapped to each value
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim kv As Object

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                kv = d0g2f5a(.Item(ky))
                With rt
                    If Not .Exists(kv) Then
                        .Add(kv, New Scripting.Dictionary)
                    End If

                    With dcOb(.Item(kv))
                        .Add(.Count, ky)
                    End With
                End With
            Next
        End With
    End Function

    Public Function d0g2f5a(vr As Object) As Object
        '
        ' Return anyObjectthat is NOT an Object
        ' Object handling MAY be addressed later.
        '
        If TypeOf vr Is Object Then
            Stop
        Else
            d0g2f5a = vr
        End If
    End Function

    Public Function dcMaterialUsage(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim PartDoc As Inventor.PartDocument
        Dim ky As Object
        Dim pn As String
        Dim mt As String

        rt = New Scripting.Dictionary
        For Each ky In dc
            PartDoc = aiDocPart(obOf(dc.Item(ky)))
            If PartDoc Is Nothing Then
                'do Nothing
            Else
                pn = PartDoc.Propertys(gnDesign).Item(pnPartNum).Text
                mt = PartDoc.Propertys(gnDesign).Item(pnMaterial).Text
                With rt
                    If .Exists(mt) Then
                        .Item(mt) = .Item(mt) & vbCrLf & vbTab & pn
                    Else
                        .Add(mt, mt & vbCrLf & vbTab & pn)
                    End If
                End With
                PartDoc = Nothing
            End If
        Next
        dcMaterialUsage = rt
    End Function
    'lsDump dcMaterialUsage(dcAiDocsOfType(kPartDocumentObject, dcAiDocComponents(aiDocActive()))).Items

    Public Function d0g3f0() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim mt As Inventor.AssemblyDocument

        rt = New Scripting.Dictionary
        With ThisApplication.ActiveMaterialLibrary
            For Each mt In .MaterialAss
                rt.Add(mt.DisplayName, mt)
            Next
        End With
        d0g3f0 = rt
    End Function
    'lsDump d0g3f0().Keys

    Public Function dcGrpByPtNum(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ''
        ''  Returns Dictionary of Dictionaries
        ''  grouping Inventor Documents in
        ''  supplied Dictionary by their Part
        ''  Numbers.
        ''
        ''  Ideally, each Document's Part Number
        ''  should be unique, and each sub Dictionary
        ''  should contain only one Document, however,
        ''  it is possible for more than one Document
        ''  to have the same Part Number.
        ''
        ''  By returning a Dictionary of Dictionaries,
        ''  this function provides a way for the client
        ''  to detect and respond to any conflicts.
        ''
        Dim rt As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String
        Dim dn As String

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))
                dn = PartDoc.FullDocumentName
                pn = CStr(aiDocPropVal(
            PartDoc, pnPartNum, gnDesign
        ))
                With rt
                    If Not .Exists(pn) Then
                        .Add(pn, New Scripting.Dictionary)
                    End If

                    With dcOb(.Item(pn))
                        If .Exists(dn) Then
                            Stop 'because something went wrong
                            'should NOT get same Document twice!
                        Else
                            .Add(dn, PartDoc)
                        End If
                    End With
                End With
                Debug.Print("")
            Next : End With

        dcGrpByPtNum = rt
        'lsDump               dcGrpByPtNum(dcAiDocComponents(aiDocActive())).Keys
        'Debug.Print(txDumpLs(dcGrpByPtNum(dcAiDocComponents(aiDocActive())).Keys)
    End Function

    Public Function dcRemapByPtNum(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ''  Returns Dictionary of Inventor
        ''  Documents keyed on Part Number
        Dim rt As Scripting.Dictionary
        Dim xt As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String

        rt = New Scripting.Dictionary
        xt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))
                pn = CStr(aiDocPropVal(
                PartDoc, pnPartNum, gnDesign
            ))
                '
                '
                '
                'If PartDoc.DocumentType =DocumentTypeEnum.kPartDocumentObject Then
                '    ' UPDATE[2021.06.22]
                '    ' moving hardware component check outside of
                '    ' and before collision check in order to skip
                '    ' known hardware items entirely.
                '    ' '
                '    ' UPDATE[2021.06.21]
                '    ' implementing a  of checks for hardware components
                '    ' probably want to move outside Dictionary check
                '    ' to catch hardware elements before they're added.
                '    ' Could lead to trouble if this prevents new hardware
                '    ' Items from being added to Genius, but don't believe
                '    ' this is a high risk.
                '    If aiDocPart(PartDoc).ComponentDefinition.IsContentMember Then
                '        'it's commodity hardware
                '        With cnGnsDoyle().Execute( _
                '            "select Family from vgMfiItems where Item = '" _
                '            & pn & "';" _
                '        )
                '            If .BOF And .EOF Then
                '                'probably not in Genius
                '                'keep it -- may need added
                '            ElseIf Split(.GetString( _
                '                adClipString, , "", vbVerticalTab _
                '            ), vbVerticalTab)(0) = "D-HDWR" Then
                '                     PartDoc = Nothing 'and move on
                '                Else
                '                    Debug.Print("") 'Breakpoint Landing
                '                    'Stop
                '                End If
                '            End If
                '        End With
                '    ElseIf PartDoc.Propertys.Item(gnDesign).Item(pnFamily).Text = "D-HDWR" Then
                '        'it's in commodity hardware family
                '         PartDoc = Nothing 'and move on
                '    ElseIf InStr(1, "|D-HDWR|D-PTS|R-PTS|", "|" & PartDoc.Propertys.Item(gnDesign).Item(pnFamily).Text & "|") > 0 Then
                '        'it's PROBABLY hardware
                '        'but keep it, just in case
                '        Debug.Print("") 'Breakpoint Landing
                '    Else 'nothing special to worry about, probably
                '        Debug.Print("") 'Breakpoint Landing
                '        'Stop
                '    End If
                'Else 'we've got an Assembly
                '    Debug.Print("") 'Breakpoint Landing
                '    'Stop
                'End If
                '
                '
                '
                If Len(pn) > 0 Then
                    With rt
                        If .Exists(pn) Then 'we have Key collsion
                            With xt 'report' it here
                                If Not .Exists(pn) Then
                                    .Add(pn, New Scripting.Dictionary)
                                End If

                                With dcOb(.Item(pn))
                                    .Add(PartDoc.FullDocumentName, PartDoc)
                                End With
                            End With
                        Else
                            .Add(pn, PartDoc)
                        End If
                    End With
                    Debug.Print("")
                Else
                    Debug.Print(InputBox("This component has no part number:" & vbCrLf & PartDoc.DisplayName & vbCrLf & CStr(aiDocPropVal(PartDoc, pnDesc, gnDesign)) & vbCrLf & vbCrLf & "Copy file path from text box for later review.", PartDoc.DisplayName, PartDoc.FullDocumentName))
                    If getFromClipBdWin10() = PartDoc.FullDocumentName Then
                        'Stop
                    Else
                        If MsgBox("Are you sure you want to continue" & vbCrLf & "without recording this file path?", vbExclamation + vbYesNo, "File Path not copied!") = vbNo Then
                            Stop
                        End If
                    End If
                End If
            Next
        End With

        If xt.Count > 0 Then
            Debug.Print(MsgBox(
            Join({
                "The following Part Numbers are",
                "assigned to more than one Model:",
                "",
                vbTab & Join(xt.Keys, vbCrLf & vbTab),
                ""
            }, vbCrLf),
            vbOKOnly Or vbInformation,
            "Duplicate Part Numbers!"
        ))
        End If

        dcRemapByPtNum = rt
        'lsDump               dcRemapByPtNum(dcAiDocComponents(aiDocActive())).Keys
        'Debug.Print(txDumpLs(dcRemapByPtNum(dcAiDocComponents(aiDocActive())).Keys)
    End Function

    Public Function dcRemapByFilePath(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ''  Returns Dictionary of Inventor
        ''  Documents re-keyed to File Path.
        ''  Typically for a Dictionary
        ''  previously remapped to another
        ''  key (most likely Part Number)
        Dim rt As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))
                pn = PartDoc.FullDocumentName
                rt.Add(pn, PartDoc)
            Next
        End With
        dcRemapByFilePath = rt
        'lsDump dcRemapByFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys
        'send2clipBd txDumpLs(dcRemapByFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys)
    End Function

    Public Function dcRemapByPtNumFilePath(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        ''  Returns Dictionary of Inventor
        ''  Documents keyed on Part Number
        ''  combined with original key
        ''  (which SHOULD be full doc path)
        Dim rt As Scripting.Dictionary
        Dim PartDoc As Inventor.Document
        Dim ky As Object
        Dim pn As String

        rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                PartDoc = aiDocument(.Item(ky))
                pn = CStr(aiDocPropVal(
                PartDoc, pnPartNum, gnDesign
            )) & vbTab & PartDoc.FullDocumentName 'ky
                rt.Add(pn, PartDoc)
            Next
        End With
        dcRemapByPtNumFilePath = rt
        'lsDump dcRemapByPtNumFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys
        'send2clipBd txDumpLs(dcRemapByPtNumFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys)
    End Function

    Public Function dcGeniusItems() As Scripting.Dictionary
        ''  Generates Dictionary of Items in Genius
        ''  Formerly d0g3f2
        Dim rt As Scripting.Dictionary
        Dim ky As ADODB.Field
        Dim vl As ADODB.Field

        rt = New Scripting.Dictionary
        With CnGnsDoyle()
            With .Execute("select Item, ItemID from vgMfiItems")
                ky = .Fields("Item")
                vl = .Fields("ItemID")
                Do Until .EOF Or .BOF
                    rt.Add(ky.Value, vl.Value)
                    .MoveNext()
                Loop
            End With
        End With
        dcGeniusItems = rt
    End Function

    Public Function d0g3f3(
    dc As Scripting.Dictionary
) As Scripting.Dictionary()
        ''  Appears intended to separate
        ''  parts in Genius from those
        ''  not there yet. I don't think
        ''  this one was working quite
        ''  right yet. New KyPick system
        ''  should handle this properly,
        ''  now, in any case.
        Dim ky As Object
        Dim rt(1) As Scripting.Dictionary

        rt(0) = New Scripting.Dictionary
        rt(1) = New Scripting.Dictionary
        With dcGeniusItems()
            For Each ky In dc
                If .Exists(ky) Then
                    rt(1).Add(ky, .Item(ky))
                Else
                    rt(0).Add(ky, "")
                End If
            Next
        End With
        d0g3f3 = rt
    End Function

    Public Function deConstrainAssyComponent(co As Inventor.ComponentOccurrence) As Long
        ''  Deletes all constraints on an occurrence
        ''  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        Dim cs As Inventor.AssemblyConstraint
        Dim ct As Long

        ct = 0
        For Each cs In co.Constraints
            cs.Delete()
            ct = ct + 1
        Next

        deConstrainAssyComponent = ct
    End Function

    Public Function deConstrainAssyDocument(ad As Inventor.AssemblyDocument) As Long
        ''  Calls deConstrainAssyComponent over all occurrences
        ''  in an assembly to remove all their constraints
        ''  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        ''  !!!That goes DOUBLE for THIS function!!!
        Dim co As Inventor.ComponentOccurrence
        Dim ct As Long

        ct = 0
        For Each co In ad.ComponentDefinition.Occurrences
            ct = ct + deConstrainAssyComponent(co)
        Next

        deConstrainAssyDocument = ct
    End Function

    Public Function dcPartsInGeniusOrNot() As Scripting.Dictionary
        Dim dcInGns As Scripting.Dictionary
        Dim dcNotIn As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        dcInGns = New Scripting.Dictionary
        dcNotIn = dcRemapByPtNum(dcAiDocComponents(aiDocActive()))
        ' dcNotIn = dcAiDocComponents(aiDocActive())
        ' dcNotIn = dcAssyDocsByPtNum(aiDocActive())
        With DcFrom2Fields(CnGnsDoyle().Execute(
        "select Item from vgMfiItems"
    ), "Item", "Item")
            For Each ky In .Keys
                With dcNotIn
                    If .Exists(ky) Then
                        dcInGns.Add(ky, .Item(ky))
                        .Remove(ky)
                    End If
                End With
            Next
        End With
        'Stop

        rt.Add("INGNS", dcInGns)
        rt.Add("NOTIN", dcNotIn)
        dcPartsInGeniusOrNot = rt
    End Function

    Public Function d0g4f1() As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim cn As ADODB.Connection
        Dim rs As ADODB.Record

        rt = New Scripting.Dictionary
        cn = CnGnsDoyle()
        rs = cn.Execute("select Item from vgMfiItems")
        d0g4f1 = rt
    End Function

    Public Function compOccFromProxy(
    oc As Inventor.ComponentOccurrence
) As Inventor.ComponentOccurrence
        Dim px As Inventor.ComponentOccurrenceProxy

        'If TypeOf oc Is Inventor.ComponentOccurrenceProxy Then
        ' px = oc
        'Stop
        ' compOccFromProxy = compOccFromProxy(px.ContainingOccurrence)
        'Else
        compOccFromProxy = oc
        'End If
    End Function

    Public Function nuPicker(
    Optional ob As KyPick = Nothing
) As KyPick
        If ob Is Nothing Then
            nuPicker = New KyPick
        Else
            nuPicker = ob
        End If
    End Function

    Public Function nuSplitter(
    Optional ob As DcSplitter = Nothing
) As DcSplitter
        If ob Is Nothing Then
            nuSplitter = New DcSplitter
        Else
            nuSplitter = ob
        End If
    End Function

    Public Function dcAiPartDocsWithRMv0(
    dcIn As Scripting.Dictionary,
    Optional WantOut As Long = 0
) As Scripting.Dictionary
        Dim ky As Object
        Dim rt(1) As Scripting.Dictionary

        If WantOut < 0 Or WantOut > 1 Then
            dcAiPartDocsWithRMv0 = dcAiPartDocsWithRMv0(dcIn, 1)
        Else
            'With nuSplitter().WithSel(New KyPickAiDocWithRM)
            With New KyPickAiDocWithRM
                rt(0) = .DcIn()
                rt(1) = .DcOut()
                For Each ky In dcIn.Keys
                    With .DcFor(dcIn.Item(ky))
                        If .Exists(ky) Then
                            Stop
                        Else
                            .Add(ky, dcIn.Item(ky))
                        End If
                    End With
                Next
            End With
            dcAiPartDocsWithRMv0 = rt(WantOut)
        End If
    End Function
    'Debug.Print(txDumpLs(dcAiPartDocsWithRMv0(dcAiDocComponents(aiDocActive()), 1).Keys)

    Public Function kyScanned(
    dcIn As Scripting.Dictionary,
    Optional pkr As KyPick = Nothing
) As KyPick
        If pkr Is Nothing Then
            kyScanned = kyScanned(dcIn, New KyPick)
        Else
            kyScanned = pkr.AfterScanning(dcIn)
        End If
    End Function
    'Debug.Print(txDumpLs(kyScanned(dcAiDocComponents(aiDocActive()), New KyPickAiPartVsAssy).dcIn().Keys)

    Public Function dcAiDocsPicked(
    dcIn As Scripting.Dictionary,
    Optional pkr As KyPick = Nothing,
    Optional WantOut As Long = 0
) As Scripting.Dictionary
        Dim ky As Object
        Dim rt(1) As Scripting.Dictionary

        If pkr Is Nothing Then
            dcAiDocsPicked = dcAiDocsPicked(dcIn, New KyPick, WantOut)
        ElseIf WantOut < 0 Or WantOut > 1 Then
            dcAiDocsPicked = dcAiDocsPicked(dcIn, pkr, 1)
        Else
            With pkr
                rt(0) = .DcIn()
                rt(1) = .DcOut()
                For Each ky In dcIn.Keys
                    With .DcFor(dcIn.Item(ky))
                        If .Exists(ky) Then
                            Stop
                        Else
                            .Add(ky, dcIn.Item(ky))
                        End If
                    End With
                Next
            End With
            dcAiDocsPicked = rt(WantOut)
        End If
    End Function
    'Debug.Print(txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), 1).Keys)
    'Debug.Print(txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), New KyPickAiDocContentCtr, 0).Keys)

    Public Function dcAiPartDocsWithRM(
    dcIn As Scripting.Dictionary,
    Optional WantOut As Long = 0
) As Scripting.Dictionary
        dcAiPartDocsWithRM = dcAiDocsPicked(dcAiDocsPicked(dcIn,
        New KyPickAiPartVsAssy, 0),
        New KyPickAiDocWithRM, WantOut)
    End Function
    'Debug.Print(txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)

    Public Function d0g5f0()
        '
    End Function

    Public Function dcAiDocGrpsByForm(
    dcIn As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcAiDocGrpsByForm -- Separate a Dictionary
        '     of Inventor Documents into
        '     categorical sub-Dictionaries
        '     according to various criteria:
        '     -   PRCH    Purchased Items
        '     -   ASSY    Assemblies
        '     -   HDWR    Hardware (Content Center)
        '     -   DBAR    Structural Parts (was BSTK)
        '                 Subtype NOT Sheet Metal
        '                 (some "Sheet Metal" Parts
        '                  might also technically
        '                  belong here. see below)
        '     -   MAYB    Likely Structural Parts
        '                 Sheet Metal subtype, but
        '                 has either no flat pattern,
        '                 or an invalid one.
        '     -   SHTM    Sheet Metal Parts
        '                 Indicated both by Subtype
        '                 and presence of a valid
        '                 flat pattern.
        '
        '     Presence in Genius, a distinction
        '     originally intended to be made here,
        '     is now planned to be made to a separate
        '     Dictionary, possibly also subcategorized,
        '     to be processed in conjunction with
        '     the results of this function.
        '
        '     The notion of passing different subgroups
        '     of this Dictionary to separate handlers
        '     for more specialized processing, while
        '     still an option, is no longer considered
        '     its primary role. Instead, the  is
        '     expected to be used in a form application
        '     which will present the various groups to
        '     the user for review, and modification as
        '     desired or necessary.
        '
        '     REV[2022.03.08.1212] All new text
        '     in function description above.
        '     see notes_2022-0308_general-01.txt
        '     for prior description
        '
        Dim rt As Scripting.Dictionary
        'Dim pkGns As KyPick
        Dim pkBuy As KyPick
        Dim pkPrt As KyPick
        Dim pkCtC As KyPick
        Dim pkSht As KyPick
        Dim pkMbe As KyPick

        rt = New Scripting.Dictionary

        ' REV[2022.03.08.1112]
        '     Disabled split on presence
        '     in Genius. Believe better
        '     addressed separately
        ''  separate items already in Genius
        ''  from those not yet in
        ' pkGns = nuPicker( _
        '    New KyPickInGenius _
        ').AfterScanning(dcIn)
        ''  NOTE: no further processing
        ''  implemented on this yet
        ''  MIGHT be better applied
        ''  at a different stage?

        ' REV[2022.03.08.1115]
        '     Add division on Purchased Parts
        '     with "out" Dictionary replacing
        '     main for Part/Assy separation.
        pkBuy = nuPicker(
        New KyPickAiDocPurchased
    ).AfterScanning(dcIn)

        With pkBuy
            rt.Add("PRCH", .DcIn)

            ''  separate parts from assemblies
            pkPrt = nuPicker(
            New KyPickAiPartVsAssy
        ).AfterScanning(.DcOut)
        End With

        With pkPrt
            rt.Add("ASSY", .DcOut)

            ''  isolate Content Center
            ''  parts from the rest
            pkCtC = nuPicker(
            New KyPickAiDocContentCtr
        ).AfterScanning(.DcIn)
        End With

        With pkCtC
            rt.Add("HDWR", .DcIn)

            ''  separate (potential) sheet
            ''  metal parts from non-sheet
            pkSht = nuPicker(
            New KyPickAiSheetMetal
        ).AfterScanning(.DcOut)
        End With

        With pkSht
            rt.Add("DBAR", .DcOut)

            pkMbe = nuPicker(
            New KyPickAiShMtl4sure
        ).AfterScanning(.DcIn)
        End With

        With pkMbe
            rt.Add("SHTM", .DcIn)
            rt.Add("MAYB", .DcOut)
        End With

        Debug.Print("")  'Breakpoint Landing

        dcAiDocGrpsByForm = rt
        'send2clipBd Join({"#If False Then", ConvertToJson(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive())), vbTab), "#End If"), vbCrLf)
        'send2clipBd ConvertToJson(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive())), vbTab)
        'Debug.Print(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))
        'Debug.Print(txDumpLs(pkCtC.dcIn.Keys)
        'Debug.Print(dcAiDocGrpsByForm(dcAssyDocsByPtNum(aiDocActive()))
        'Debug.Print(txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)
    End Function

    Public Function dcAiDocGrpsByFormAndIfac(
    dcIn As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcAiDocGrpsByFormAndIfac -- Separate a Dictionary
        '     of Inventor Documents into
        '     categorical sub-Dictionaries
        '     according to various criteria:
        '     -   PRCH    Purchased Items
        '     -   ASSY    Assemblies
        '     -   IASM    iAssembly Factories
        '     -   IPRT    iPart Factories
        '     -   HDWR    Hardware (Content Center)
        '     -   DBAR    Structural Parts (was BSTK)
        '                 Subtype NOT Sheet Metal
        '                 (some "Sheet Metal" Parts
        '                  might also technically
        '                  belong here. see below)
        '     -   MAYB    Likely Structural Parts
        '                 Sheet Metal subtype, but
        '                 has either no flat pattern,
        '                 or an invalid one.
        '     -   SHTM    Sheet Metal Parts
        '                 Indicated both by Subtype
        '                 and presence of a valid
        '                 flat pattern.
        '
        ' REV[2023.01.24.0921]
        ' copied from dcAiDocGrpsByForm to produce
        ' newObjectwith additional groupings for
        ' iPart (IPRT) and iAssembly (IASM) members.
        ' additional groups for their corresponding
        ' Factories will likely also be added.
        '
        ' text of prior REV[2022.03.08.1212] removed.
        ' see dcAiDocGrpsByForm for that.
        '
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        ' REV[2023.01.24.1156]
        ' add working Dictionary
        ' to collect iAssembly
        ' and iPart Factories
        Dim ky As Object
        Dim mb As Inventor.Document
        Dim md As Inventor.Document
        Dim fp As String

        'Dim pkGns As KyPick
        Dim pkBuy As KyPick
        Dim pkPrt As KyPick
        Dim pkCtC As KyPick
        Dim pkSht As KyPick
        Dim pkMbe As KyPick

        ' REV[2023.01.24.1009]
        ' add new pickers for
        ' iAssemblies and iParts
        Dim pkIas As KyPick
        Dim pkIpt As KyPick

        rt = New Scripting.Dictionary

        ' REV[2022.03.08.1112]
        '     Disabled split on presence
        '     in Genius. Believe better
        '     addressed separately
        ''  separate items already in Genius
        ''  from those not yet in
        ' pkGns = nuPicker( _
        '    New KyPickInGenius _
        ').AfterScanning(dcIn)
        ''  NOTE: no further processing
        ''  implemented on this yet
        ''  MIGHT be better applied
        ''  at a different stage?

        ' REV[2022.03.08.1115]
        '     Add division on Purchased Parts
        '     with "out" Dictionary replacing
        '     main for Part/Assy separation.
        pkBuy = nuPicker(
        New KyPickAiDocPurchased
    ).AfterScanning(dcIn)

        With pkBuy
            If dcIn.Count > 0 Then rt.Add("PRCH", .DcIn)

            ''  separate parts from assemblies
            pkPrt = nuPicker(
            New KyPickAiPartVsAssy
        ).AfterScanning(.DcOut)
        End With

        With pkPrt
            ''  separate iAssembly members
            ''  from stand-alone assemblies
            pkIas = nuPicker(
            New KyPickAiAssyMember
        ).AfterScanning(.DcOut)

            ''  isolate Content Center
            ''  parts from the rest
            pkCtC = nuPicker(
            New KyPickAiDocContentCtr
        ).AfterScanning(.DcIn)
        End With

        With pkIas
            If .DcOut.Count > 0 Then rt.Add("ASSY", .DcOut)

            With .DcIn
                wk = New Scripting.Dictionary

                For Each ky In .Keys
                    With aiDocAssy(.Item(ky)).ComponentDefinition
                        mb = .Parent
                        With .iAssemblyMember.ParentFactory.Parent
                            md = .Parent
                            fp = md.FullDocumentName

                            With wk
                                If Not .Exists(fp) Then
                                    .Add(fp, New Scripting.Dictionary)
                                    dcOb(.Item(fp)).Add("", md)
                                End If

                                With dcOb(.Item(fp))
                                    If .Exists(ky) Then
                                        Stop
                                    Else
                                        .Add(ky, mb)
                                    End If
                                End With

                            End With
                        End With
                    End With
                    'Debug.Print(aiDocAssy(.Item(ky)).ComponentDefinition.iAssemblyMember.ParentFactory.Parent.Parent.FullDocumentName
                    'Stop
                Next

                If wk.Count > 0 Then rt.Add("IASM", wk)
            End With
        End With

        With pkCtC
            If .DcIn.Count > 0 Then rt.Add("HDWR", .DcIn)

            ''  separate iPart members
            ''  from stand-alone parts
            pkIpt = nuPicker(
            New KyPickAiPartMember
        ).AfterScanning(.DcOut)
        End With

        With pkIpt
            With .DcIn
                wk = New Scripting.Dictionary

                For Each ky In .Keys
                    With aiDocPart(.Item(ky)).ComponentDefinition
                        mb = .Document '.Parent
                        md = aiDocPart(.iPartMember.ParentFactory.Parent) ' .Propertys.Parent ' .Parent
                        With md
                            fp = md.FullDocumentName

                            With wk
                                If Not .Exists(fp) Then
                                    .Add(fp, New Scripting.Dictionary)
                                    dcOb(.Item(fp)).Add("", md)
                                End If

                                With dcOb(.Item(fp))
                                    If .Exists(ky) Then
                                        Stop
                                    Else
                                        .Add(ky, mb)
                                    End If
                                End With
                            End With
                        End With
                    End With : Next

                If wk.Count > 0 Then rt.Add("IPRT", wk)
            End With

            ''  add iPart Factories to
            ''  Dictionary of non-Members
            With .DcOut : For Each ky In wk.Keys
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add(ky, dcOb(wk.Item(ky)).Item(""))
                    End If
                Next : End With

            ''  separate (potential) sheet
            ''  metal parts from non-sheet
            pkSht = nuPicker(
            New KyPickAiSheetMetal
        ).AfterScanning(.DcOut)
        End With

        With pkSht
            If .DcOut.Count > 0 Then rt.Add("DBAR", .DcOut)

            pkMbe = nuPicker(
            New KyPickAiShMtl4sure
        ).AfterScanning(.DcIn)
        End With

        With pkMbe
            If .DcIn.Count > 0 Then rt.Add("SHTM", .DcIn)
            If .DcOut.Count > 0 Then rt.Add("MAYB", .DcOut)
        End With

        Debug.Print("")  'Breakpoint Landing

        dcAiDocGrpsByFormAndIfac = rt
        'send2clipBd Join({"#If False Then", ConvertToJson(dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive())), vbTab), "#End If"), vbCrLf)
        'send2clipBd ConvertToJson(dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive())), vbTab)
        'Debug.Print(dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive()))
        'Debug.Print(txDumpLs(pkCtC.dcIn.Keys)
        'Debug.Print(dcAiDocGrpsByFormAndIfac(dcAssyDocsByPtNum(aiDocActive()))
        'Debug.Print(txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)
    End Function

    Public Function d0g5f2(
    dcIn As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' function d0g5f2
        '
        ' INITIATED[2021.03.23]
        ' thisObjecton dcAiDocGrpsByForm
        ' is intended to separate items
        ' in Genius from those not yet in
        ' and purchased items from those
        ' to be made, cross-referencing
        ' the two to determine individual
        ' needs for processing.
        '
        ' presently in a nonfunctional state
        ' as End of Day approaches. will hope
        ' to continue development tomorrow
        '
        Dim rt As Scripting.Dictionary
        Dim pkGns As KyPick
        Dim pkPvA As KyPick
        Dim pkCtC As KyPick
        Dim pkSht As KyPick
        Dim pkBuy As KyPick

        rt = New Scripting.Dictionary

        ''  separate items already in Genius
        ''  from those not yet in
        pkGns = nuPicker(
        New KyPickInGenius
    ).AfterScanning(dcIn)

        ''  separate purchased items
        ''  from those to be made
        pkBuy = nuPicker(
        New KyPickAiDocPurchased
    ).AfterScanning(dcIn)
        ''  NOTE: no further processing
        ''  implemented on this yet
        ''  MIGHT be better applied
        ''  at a different stage?



        ''  separate parts from assemblies
        ' pkPvA = nuPicker( _
        '    New KyPickAiPartVsAssy _
        ').AfterScanning(dcIn)
        ''rt.Add("ASSY", dck pkPvA.dcOut

        ''  isolate Content Center
        ''  parts from the rest
        ' pkCtC = nuPicker( _
        '    New KyPickAiDocContentCtr _
        ').AfterScanning(pkPvA.dcIn)
        ''rt.Add("HDWR", pkCtC.dcIn

        ''  separate (potential)
        ''  sheet metal parts
        ''  from non-sheet
        ' pkSht = nuPicker( _
        '    New KyPickAiSheetMetal _
        ').AfterScanning(pkCtC.dcOut)
        'rt.Add("SHTM", pkSht.dcIn
        'rt.Add("BSTK", pkSht.dcOut

        Debug.Print("")  'Breakpoint Landing

        d0g5f2 = rt
    End Function

    Public Function d0g5f3(
    dcIn As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' function d0g5f3 -- essentially a recreation of dcAiDocGrpsByForm
        '
        Dim rt As Scripting.Dictionary
        Dim pkBuy As KyPick
        Dim pkPrt As KyPick
        Dim pkCtC As KyPick
        Dim pkSht As KyPick
        Dim pkMbe As KyPick
        Dim pkGns As KyPick
        'Dim pk___ As KyPick

        rt = New Scripting.Dictionary
        pkGns = nuPicker(
        New KyPickInGenius
    ).AfterScanning(dcIn)

        pkBuy = nuPicker(
        New KyPickAiDocPurchased
    ).AfterScanning(dcIn)

        With pkBuy
            rt.Add("PRCH", .DcIn)
            pkPrt = nuPicker(
            New KyPickAiPartVsAssy
        ).AfterScanning(.DcOut)
        End With

        With pkPrt
            rt.Add("ASSY", .DcOut)
            pkCtC = nuPicker(
            New KyPickAiDocContentCtr
        ).AfterScanning(.DcIn)
        End With

        With pkCtC
            rt.Add("HDWR", .DcIn)
            pkSht = nuPicker(
            New KyPickAiSheetMetal
        ).AfterScanning(.DcOut)
        End With

        With pkSht
            rt.Add("DBAR", .DcOut)

            pkMbe = nuPicker(
            New KyPickAiShMtl4sure
        ).AfterScanning(.DcIn)
        End With

        With pkMbe
            rt.Add("MAYB", .DcOut)
            rt.Add("SHTM", .DcIn)
        End With

        d0g5f3 = rt
    End Function

    Public Function dcDepthAiDocGrp(
    dc As Scripting.Dictionary
) As Long
        Dim vr As Object
        Dim ob As Object
        Dim mx As Long
        Dim dx As Long
        Dim ck As Long
        Dim rt As Long

        With dc
            mx = .Count

            If mx = 0 Then
                dcDepthAiDocGrp = 0 'indeterminate
            Else
                dx = 0

                Do
                    'vr = {.Item(.Keys(dx)))
                    ob = obOf(.Item(.Keys(dx))) 'vr(0)
                    If ob Is Nothing Then 'IsObject(vr(0))
                        ck = -1 'invalid
                    Else
                        If TypeOf ob Is Scripting.Dictionary Then
                            ck = dcDepthAiDocGrp(ob)
                            If ck > 0 Then ck = 1 + ck
                        ElseIf TypeOf ob Is Inventor.Document Then
                            ck = 1
                        Else
                            ck = -1 'invalid
                        End If
                    End If

                    dx = dx + 1
                    If dx > mx Then ck = -1 'invalid

                Loop While ck = 0 'indeterminate

                dcDepthAiDocGrp = ck
            End If
        End With
    End Function

    Public Function nu_fmIfcTest05A(
    Optional dcIn As Scripting.Dictionary = Nothing
) As FmIfcTest05A
        With New FmIfcTest05A
            nu_fmIfcTest05A = .UseDict(dcIn)
        End With
        'nu_fmIfcTest05A(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))).Show(1
        'nu_fmIfcTest05A(dcAiDocGrpsByForm(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive())))).Show(1
        'nu_fmIfcTest05A(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive()))).Show(1
    End Function

    Public Function nu_fmTest05A(
    Optional dcIn As Scripting.Dictionary = Nothing
) As FmTest05
        Stop 'DO NOT USE THIS FUNCTION!
        'instead, use the Interface
        'generator nu_fmIfcTest05A

        With New FmTest05
            nu_fmTest05A = .Holding(dcIn) ' .UseDict(dcIn)
        End With
        'nu_fmTest05A(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))).Show(1
        'nu_fmTest05A(dcAiDocGrpsByForm(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive())))).Show(1
        'nu_fmTest05A(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive()))).Show(1
    End Function

    Public Function lsAssyMembers(
    aiAssy As Inventor.AssemblyDocument
) As String
        Dim dc As Scripting.Dictionary
        Dim pn As String
        Dim rt As String
        Dim ky As Object

        dc = d0g3f0.dcAiDocsByPtNum(dcAssyComponentsImmediate(aiAssy)) 'dcAiDocPartNumbers
        pn = vbCrLf & aiAssy.Propertys.Item(
        gnDesign).Item(pnPartNum).Text & vbTab
        rt = pn & Join(dc.Keys, pn)

        With nuPicker(
        New KyPickAiPartVsAssy
    ).AfterScanning(dc)
            With .DcOut
                For Each ky In .Keys
                    rt = rt & lsAssyMembers(aiDocument(obOf(.Item(ky))))
                Next
            End With
        End With

        lsAssyMembers = rt
    End Function

    Public Function d0g6f0(AiDoc As Inventor.Document) As String
        ''  Try to pick a distinct listing name
        ''  for a supplied Inventor Document
        Dim rt As String
        Dim ds As String

        With AiDoc
            With .Propertys(gnDesign)
                rt = Trim$(.Item(pnPartNum).Text)
                ds = Trim$(.Item(pnDesc).Text)
            End With

            If Len(rt) > 0 Then
                If Len(ds) > 0 Then rt = rt & ": " & ds
            ElseIf Len(ds) > 0 Then
                rt = ds
            End If

            If Len(rt) = 0 Then
                ds = .FullFileName
                If Len(ds) > 0 Then
                    With nuFso().GetFile(ds)
                        rt = .Name & " (" & .ParentFolder.Path & ")"
                    End With
                Else
                    rt = .DisplayName
                End If
            End If

            d0g6f0 = rt
        End With
    End Function

    Public Function d0g6f1()
        ''
        ''  testing form class FmTest0
        ''
        With New FmTest0
            .imTNail.Visible = False
            Debug.Print(.Controls.Count)
            Stop
        End With
    End Function

    Public Function d0g6f2(dc As Scripting.Dictionary) As Scripting.Dictionary
        ' Call this one from inside dcAiDocGrpsByForm (above)
        ' Try: debug.Print(txDumpLs(d0g6f2(pkPvA.dcIn).Keys)
        '
        Dim rt As Scripting.Dictionary
        Dim ad As Inventor.Document
        Dim InvProperty As Inventor.Property
        Dim ky As Object

        rt = New Scripting.Dictionary

        With dc
            For Each ky In .Keys
                ad = aiDocument(.Item(ky))
                If ad Is Nothing Then
                    'nothing we can do with it
                Else
                    InvProperty = ad.Propertys(gnDesign).Item(pnFamily)
                    rt.Add(ky, InvProperty)
                    With InvProperty
                        .Text = "R-PTS"
                        'Stop
                    End With
                End If
            Next
        End With

        d0g6f2 = rt
        '
        '
    End Function

    Public Function d0g6f3()
        ''
        ''  testing new empty form class FmEmpty
        ''
        With New fmEmpty
            Dim cmb As New MSForms.ComboBox
            cmb.Name = "test"
            cmb.Left = 10
            cmb.Top = 10
            .Controls.Add(cmb)
            Debug.Print(.Controls.Count)
            .Show()
            Stop
        End With
    End Function

    'Public Function d0g7f0() As String
    '    '
    '    ' This function used to transfer Property Values
    '    ' from blank model files GR12 ~ GR20
    '    ' to new versions generated from Intraflo's
    '    ' supplied STEP files. Save for reference,
    '    ' but this version should not likely be used
    '    ' as is for other tasks without review.
    '    '
    '    Dim rt As Scripting.Dictionary
    '    Dim dcPr As Scripting.Dictionary
    '    Dim sd As Inventor.Document
    '    Dim td As Inventor.Document
    '    Dim PropertySetSc As Inventor.Property
    '    Dim PropertySetTg As Inventor.Property
    '    Dim prSc As Inventor.Property
    '    Dim prTg As Inventor.Property
    '    Dim ky As Object
    '    Dim pn As Object
    '    Dim StockNumb As String

    '    rt = d0g3f0.dcAiDocsByPtNum(dcAssyComponentsImmediate(aiDocActive()))
    '    With rt
    '        For Each ky In .Keys
    '            Debug.Print(ky) 'This might want to be disabled
    '            sd = aiDocument(obOf(.Item(ky)))
    '            If UCase$(Left$(ky, 2)) = "GR" Then
    '                StockNumb = sd.Propertys(gnDesign).Item(pnStockNum).Text
    '                If .Exists(StockNumb) Then
    '                    td = aiDocument(obOf(.Item(StockNumb)))
    '                    Debug.Print("")  'Stop
    '                    PropertySetTg = td.Propertys(gnCustom)
    '                    dcPr = dcAiPropsIn(.psTg)
    '                    PropertySetSc = sd.Propertys(gnCustom)
    '                    For Each prSc In PropertySetSc
    '                        If dcPr.Exists(prSc.Name) Then
    '                            With PropertySetTg.Item(prSc.Name)
    '                                .Text = prSc.Text
    '                                Debug.Print("")  'Landing Point -- Ctrl-F8 to here
    '                            End With
    '                        Else
    '                            With prSc
    '                                PropertySetTg.Add(.text, .Name) ', .PropId
    '                                Debug.Print("")  'Landing Point -- Ctrl-F8 to here
    '                            End With
    '                        End If
    '                    Next

    '                    PropertySetSc = sd.Propertys(gnDesign)
    '                    With td.Propertys(gnDesign)
    '                        For Each pn In {
    '                        pnPartNum, pnStockNum,
    '                        pnFamily, pnDesc,
    '                        pnCatWebLink}

    '                            '.Item(pnStockNum).Text = PropertySetSc.Item(pnStockNum).Text
    '                            '.Item(pnFamily).Text = PropertySetSc.Item(pnFamily).Text
    '                            '.Item(pnCatWebLink).Text = PropertySetSc.Item(pnCatWebLink).Text
    '                            '.Item(pnDesc).Text = PropertySetSc.Item(pnDesc).Text
    '                            '.Item(pnPartNum).Text = PropertySetSc.Item(pnPartNum).Text
    '                            '.Item(pn).Text = PropertySetSc.Item(pn).Text
    '                            .Item(CStr(pn)).Text = PropertySetSc.Item(CStr(pn)).Text
    '                            Debug.Print("")  'Landing Point -- Ctrl-F8 to here
    '                        Next
    '                    End With
    '                Else
    '                End If
    '            Else
    '            End If
    '        Next
    '    End With
    'End Function

    Public Function d0g8f0() As String
        Dim dx As Long
        Dim fn As String

        For dx = 1 To 16
            fn = "Specification" & CStr(dx)
            With CnGnsDoyle().Execute(Join({
            "select distinct", fn,
            "from vgMfiItems",
            "where Family = 'D-BAR'",
            "and", fn, "is not null",
            "and", fn, "<> ''",
            "order by", fn, ";", " "}))
                If .BOF Or .EOF Then
                    'Debug.Print("<EMPTY>"
                    'Debug.Print
                Else
                    Debug.Print("[" & fn & "]")
                    Debug.Print(.GetString)
                End If
            End With
        Next
    End Function

    Public Function d0g9f0(
    Optional ad As Inventor.Document = Nothing,
    Optional pn As String = ""
) As String

        Dim vw As Inventor.View
        Dim cm As Inventor.Camera
        Dim bp As String

        If ad Is Nothing Then
            d0g9f0 = d0g9f0(ThisApplication.ActiveDocument, pn)
        ElseIf Len(pn) < 1 Then
            d0g9f0 = d0g9f0(ad, d0g9f3(ad))
        Else
            bp = "C:\Doyle_Vault\Designs\Misc\andrewT\"
            vw = ad.Views.Item(1) 'ThisApplication.ActiveView
            cm = vw.Camera

            With vw
                'Debug.Print(.Left, .Top
                'Debug.Print(.Width, .Height

                Debug.Print("")  'breakpoint anchor

                With .Camera
                    .ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation
                    .Fit()
                    .Apply()
                End With
                .Fit()
                .Update()
                Debug.Print("")  'breakpoint anchor
                '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                .SaveAsBitmap(bp & pn & "-I.png", .Width, .Height) '0, 0

                With .Camera
                    .ViewOrientationType = ViewOrientationTypeEnum.kFrontViewOrientation
                    .Fit()
                    .Apply()
                End With
                .Fit()
                .Update()
                Debug.Print("")  'breakpoint anchor
                '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                .SaveAsBitmap(bp & pn & "-F.png", 0, 0)

                With .Camera
                    .ViewOrientationType = ViewOrientationTypeEnum.kTopViewOrientation
                    .Fit()
                    .Apply()
                End With
                .Fit()
                .Update()
                Debug.Print("")  'breakpoint anchor
                '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                .SaveAsBitmap(bp & pn & "-T.png", 0, 0)

                .GoHome()
                .Update()
            End With
        End If

        Debug.Print("")  'breakpoint anchor

        d0g9f0 = ""
    End Function

    Public Function d0g9f1(ad As Inventor.AssemblyDocument) As String
        Dim BillRow As Inventor.iAssemblyTableRow

        With ad.ComponentDefinition
            If .IsiAssemblyFactory Then
                With .iAssemblyFactory
                    For Each BillRow In .TableRows
                        With BillRow
                            Debug.Print(.MemberName)
                        End With
                        ' .DefaultRow = BillRow
                        'Stop
                    Next
                End With
            Else
            End If
        End With

        d0g9f1 = ""
    End Function

    Public Function d0g9f2(oc As Inventor.ComponentOccurrence) As String
        With oc
            If .IsiAssemblyMember Then
                Stop
                '.Definition
                '.Replace

            ElseIf .IsiPartMember Then
                Stop
            Else
                Stop
            End If
        End With
    End Function

    Public Function d0g9f2as(cd As Inventor.AssemblyComponentDefinition) As String
        With cd
            '.IsiAssemblyMember
            '.iAssemblyMember
            With .iAssemblyMember
                '.ParentFactory
                '.Row
            End With
        End With
    End Function

    Public Function d0g9f3(ad As Inventor.AssemblyDocument) As String
        Dim cp As Inventor.AssemblyDocument

        With ad.ComponentDefinition.Occurrences.Item(1)
            cp = aiDocAssy(.Definition.Document)
            If cp Is Nothing Then
                d0g9f3 = "NO-NUM-ASSY"
            Else
                d0g9f3 =
                cp.Propertys(
                gnDesign).Item(
                pnPartNum).Text
            End If
        End With
    End Function

    Public Sub PlaceInAssembly()
        Dim dc As Scripting.Dictionary
        Dim cm As Inventor.CommandManager
        Dim cd As Inventor.Document
        Dim ad As Inventor.Document
        Dim rp As VbMsgBoxResult
        Dim nm As String

        With ThisApplication
            If .ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject _
        Or .ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject _
        Then
                cd = .ActiveDocument
                dc = dcAiAssyDocs(
               dcAiDocsVisible())
                dc.Remove(cd.FullDocumentName)

                With nuSelAiDoc().WithList(dc.Keys)
                    Do
                        nm = .GetReply()
                        If dc.Exists(nm) Then
                            ad = dc.Item(nm)
                            rp = vbOK
                        Else
                            ad = Nothing
                            rp = MsgBox(
                            "No Valid Assembly Selected.",
                            vbRetryCancel, "No Assembly"
                        ) ' Try Again?
                        End If
                    Loop While rp = vbRetry
                End With

                If ad Is Nothing Then
                    Debug.Print("")
                Else
                    ad.Activate()
                    cm = .CommandManager
                    With cm
                        .PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, cd.FullDocumentName)
                        .ControlDefinitions.Item("AssemblyPlaceComponentCmd").Execute
                    End With
                End If
            End If
        End With
    End Sub
End Module