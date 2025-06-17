Module modGPUpdateATrev

    Public Function dcGeniusPropsPartRev20180530(ThisApplication As Inventor.Application,
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary

        ' ... [comments unchanged] ...

        Dim rt As Scripting.Dictionary
        Dim dcIn As Scripting.Dictionary
        Dim dcFP As Scripting.Dictionary
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        Dim prPartNum As Inventor.Property
        Dim prFamily As Inventor.Property
        Dim prRawMatl As Inventor.Property
        Dim prRmUnit As Inventor.Property
        Dim prRmQty As Inventor.Property
        Dim pnModel As String
        Dim nmFamily As String
        Dim mtFamily As String
        Dim qtRawMatl As Double
        Dim pnStock As String
        Dim qtUnit As String
        Dim bomStruct As Inventor.BOMStructureEnum
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData

        If dc Is Nothing Then
            dcGeniusPropsPartRev20180530 =
                dcGeniusPropsPartRev20180530(
                    oDoc, New Scripting.Dictionary
                )
        Else
            rt = dc

            With oDoc
                If .ComponentDefinition.IsContentMember Then
                    'Stop
                End If

                With .Propertys
                    aiPropsUser = .Item(GnCustom)
                    aiPropsDesign = .Item(GnDesign)
                End With

                If .ComponentDefinition.IsContentMember Then
                    pnStock = ""
                    qtRawMatl = 0#
                    qtUnit = ""
                Else
                    prRawMatl = aiGetProp(aiPropsUser, PnRawMaterial, 1)
                    prRmUnit = aiGetProp(aiPropsUser, PnRmUnit, 1)
                    prRmQty = aiGetProp(aiPropsUser, PnRmQty, 1)

                    If prRawMatl Is Nothing Then
                        pnStock = ""
                    Else
                        pnStock = prRawMatl.Text
                    End If
                    If prRmQty Is Nothing Then
                        qtRawMatl = 0#
                    ElseIf IsNumeric(prRmQty.Text) Then
                        qtRawMatl = prRmQty.Text
                    Else
                        qtRawMatl = 0#
                    End If

                    If prRmUnit Is Nothing Then
                        qtUnit = ""
                    Else
                        qtUnit = prRmUnit.Text
                    End If
                End If

                prPartNum = aiGetProp(aiPropsDesign, PnPartNum)
                pnModel = prPartNum.Text
                prFamily = aiGetProp(aiPropsDesign, PnFamily)
                nmFamily = famVsGenius(pnModel, prFamily.Text)

                With .ComponentDefinition
                    With .MassProperties
                        On Error Resume Next
                        rt = dcWithProp(aiPropsUser, PnMass,
                            System.Math.Round(CvMassKg2LbM * .Mass, 4), rt
                        )
                        If Err.Number Then
                            Stop
                        End If
                        On Error GoTo 0
                    End With

                    ck = vbNo
                    If .IsContentMember Then
                        ck = vbYes
                    ElseIf InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" & nmFamily & "|") > 0 Then
                        ck = vbYes
                    ElseIf InStr(1, oDoc.FullFileName, "\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" & prFamily.Text & "|") > 0 Then
                        ck = newFmTest2().AskAbout(oDoc, , "Is this a Purchased Part?" & vbCrLf & "(Cancel to debug)")
                    End If

                    If ck = vbCancel Then
                        Stop
                    ElseIf ck = vbYes Then
                        If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                            On Error Resume Next
                            .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                            If Err.Number = 0 Then
                                bomStruct = .BOMStructure
                            Else
                                bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                            End If
                            On Error GoTo 0
                        Else
                            bomStruct = .BOMStructure
                        End If
                    Else
                        bomStruct = .BOMStructure
                    End If

                    If bomStruct = BOMStructureEnum.kNormalBOMStructure Then
                        pnStock = prRawMatl.Text
                        dcIn = DcFromAdoRS(CnGnsDoyle().Execute(SqlOf_GnsPartMatl(ThisApplication, pnModel)))
                        If dcIn.Count > 0 Then
                            With dcOb(dcDxFromRecSetDc(dcIn).Item(PnRawMaterial))
                                If .Count > 0 Then
                                    If Len(pnStock) > 0 Then
                                        If .Exists(pnStock) Then
                                            ' do nothing
                                        Else
                                            pnStock = ""
                                        End If
                                    End If
                                    If Len(pnStock) = 0 Then
                                        pnStock = .Keys(0)
                                    End If
                                    If .Count > 1 Then
                                        Debug.Print(prPartNum.Text & vbCrLf & vbTab & Join(.Keys, vbCrLf & vbTab))
                                        pnStock = nuSelector().GetReply(.Keys, pnStock)
                                        Debug.Print("Selected " & IIf(Len(pnStock) > 0, pnStock, "(nothing)"))
                                        Stop
                                    End If
                                End If
                                If Len(pnStock) > 0 Then
                                    If Len(CStr(prRawMatl.Text)) = 0 Then
                                        Debug.Print("")
                                    ElseIf pnStock = prRawMatl.Text Then
                                        Debug.Print("")
                                    Else
                                        Debug.Print("=== CURRENT GENIUS MATERIAL DATA ===")
                                        ck = newFmTest2().AskAbout(oDoc, "Raw Material " & prRawMatl.Text & vbCrLf & " for Item", "does not match " & pnStock & vbCrLf & "indicated in Genius." & vbCrLf & vbCrLf & "Change to match Genius?" & vbCrLf & "(Cancel to debug)")
                                        ck = vbOK
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbNo Then
                                            pnStock = prRawMatl.Text
                                        End If
                                    End If
                                    If .Exists(pnStock) Then
                                        dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                        Debug.Print("")
                                    Else
                                        Stop
                                    End If
                                Else
                                    dcIn = New Scripting.Dictionary
                                End If
                            End With
                        End If

                        With dcIn
                            If .Count = 0 Then
                                .Add("Ord", 0)
                                .Add("RM", "")
                                .Add("MtFamily", "")
                                .Add("RMQTY", 0)
                                .Add("RMUNIT", "")
                            End If
                        End With

                        If .SubType = GuidSheetMetal Then
                            With dcIn
                                If .Exists("MtFamily") Then
                                    mtFamily = .Item("MtFamily")
                                Else
                                    mtFamily = ""
                                End If
                            End With

                            If Len(mtFamily) = 0 Then
                                ck = vbRetry
                            ElseIf mtFamily = "DSHEET" Then
                                ck = vbYes
                            Else
                                ck = vbNo
                            End If

                            If ck = vbNo Then
                                dcFP = New Scripting.Dictionary
                            Else
                                dcFP = dcFlatPatVals(.ComponentDefinition)
                                If dcFP.Exists(PnThickness) Then
                                    pnStock = ptNumShtMetal(oDoc.ComponentDefinition)
                                    If Len(pnStock) = 0 Then
                                        Stop
                                    End If
                                    dcFP.Add(PnRawMaterial, pnStock)
                                Else
                                    pnStock = ""
                                End If
                            End If
                            Debug.Print("")
                            If ck = vbRetry Then
                                If dcFP.Exists("mtFamily") Then
                                    If dcFP.Item("mtFamily") = "DSHEET" Then
                                        If dcFP.Exists("OFFTHK") Then
                                            ck = newFmTest2().AskAbout(oDoc, "This Part: ", "might not be sheet metal. " & vbCrLf & vbCrLf & "Is it in fact sheet metal?")
                                            If ck = vbCancel Then
                                                ck = vbRetry
                                                Stop
                                            End If
                                        Else
                                            ck = vbYes
                                        End If
                                    ElseIf dcFP.Item("mtFamily") = "D-BAR" Then
                                        ck = vbNo
                                    Else
                                        ck = vbRetry
                                    End If
                                Else
                                    ck = vbRetry
                                End If
                            End If
                            If ck = vbRetry Then
                                Stop
                            End If

                            If ck = vbYes Then
                                rt = dcFlatPatProps(.ComponentDefinition, rt)
                            ElseIf ck = vbRetry Then
                                rt = dcFlatPatProps(.ComponentDefinition, rt)
                            ElseIf ck = vbNo Then
                                ' do nothing
                            Else
                                Stop
                            End If

                            If prRawMatl Is Nothing Then
                                If rt.Exists("OFFTHK") Then
                                    Debug.Print(aiProperty(rt.Item("OFFTHK")).Text)
                                    Stop
                                    pnStock = ""
                                    If 0 Then Stop
                                Else
                                    Stop
                                    pnStock = ptNumShtMetal(.ComponentDefinition)
                                End If
                            Else
                                If Len(prRawMatl.Text) > 0 Then
                                    If Len(pnStock) = 0 Then
                                        pnStock = ptNumShtMetal(.ComponentDefinition)
                                    End If
                                    If Len(pnStock) > 0 Then
                                        If pnStock <> prRawMatl.Text Then
                                            If UCase$(prRawMatl.Text) = pnStock Then
                                                ck = vbYes
                                            Else
                                                ck = newFmTest2().AskAbout(oDoc, "Suggest Sheet Metal change from" & vbCrLf & prRawMatl.Text & " to" & vbCrLf & pnStock & " for", "Change it?")
                                            End If
                                            If ck = vbCancel Then
                                                Stop
                                            ElseIf ck = vbYes Then
                                                prRawMatl.Text = pnStock
                                            End If
                                        End If
                                    End If
                                ElseIf Len(pnStock) > 0 Then
                                    prRawMatl.Text = pnStock
                                End If

                                If Len(prRawMatl.Text) > 0 Then
                                    If rt.Exists("OFFTHK") Then
                                        ck = newFmTest2().AskAbout(oDoc, "Assigned Raw Material " & prRawMatl.Text & vbCrLf & " might be incorrect for ", "Clear it?")
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            prRawMatl.Text = ""
                                            pnStock = prRawMatl.Text
                                        End If
                                    End If
                                    If pnStock = prRawMatl.Text Then
                                        Debug.Print("")
                                    Else
                                        pnStock = prRawMatl.Text
                                    End If
                                    With CnGnsDoyle().Execute("select Family from vgMfiItems where Item='" & Replace(pnStock, "'", "''") & "';")
                                        If .BOF Or .EOF Then
                                            If pnStock <> "0" Then
                                                If Len(pnStock) > 0 Then
                                                    Stop
                                                End If
                                            End If
                                            If rt.Exists("OFFTHK") Then
                                                pnStock = ""
                                            Else
                                                pnStock = ptNumShtMetal(oDoc.ComponentDefinition)
                                                Debug.Print("")
                                            End If
                                        End If
                                    End With
                                ElseIf rt.Exists("OFFTHK") Then
                                    pnStock = ""
                                Else
                                    pnStock = ptNumShtMetal(.ComponentDefinition)
                                End If

                                If Len(pnStock) = 0 Then
                                    With newFmTest1()
                                        If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop
                                        bd = nuAiBoxData().UsingInches.SortingDims(oDoc.ComponentDefinition.RangeBox)
                                        ck = .AskAbout(oDoc, "No Stock Found! Please Review" & vbCrLf & vbCrLf & bd.Dump(0))
                                        If ck = vbYes Then
                                            With .ItemData
                                                If .Exists(PnFamily) Then
                                                    nmFamily = .Item(PnFamily)
                                                    Debug.Print(PnFamily & "=" & nmFamily)
                                                End If
                                                If .Exists(PnRawMaterial) Then
                                                    pnStock = .Item(PnRawMaterial)
                                                    Debug.Print(PnRawMaterial & "=" & pnStock)
                                                End If
                                            End With
                                            If 0 Then Stop
                                        End If
                                    End With
                                ElseIf Left$(pnStock, 2) = "LG" Then
                                    Debug.Print(pnModel & ": PROBABLE LAGGING [" & pnStock & "]")
                                    Debug.Print("  TRY TO VERIFY. IF CHANGE REQUIRED,")
                                    Debug.Print("  FILL IN NEW VALUE FOR pnStock BELOW, ")
                                    Debug.Print("  AND PRESS ENTER ON THE LINE. WHEN ")
                                    Debug.Print("  READY, PRESS [F5] TO CONTINUE.")
                                    Debug.Print("  pnStock = """ & pnStock & """")
                                    Stop
                                End If

                                If Len(pnStock) > 0 Then
                                    Stop
                                    With CnGnsDoyle().Execute("select Family from vgMfiItems where Item='" & Replace(pnStock, "'", "''") & "';")
                                        If .BOF Or .EOF Then
                                            Stop
                                        Else
                                            With .Fields
                                                mtFamily = .Item("Family").Value
                                            End With
                                            If mtFamily Like "?-MT*" Then
                                                Debug.Print(pnModel & "[" & qtRawMatl & qtUnit & " of " & pnStock & ": " & aiPropsDesign.Item(PnDesc).Text & "]")
                                                Stop
                                            ElseIf mtFamily = "D-PTS" Then
                                                Stop
                                                mtFamily = "D-BAR"
                                            ElseIf mtFamily = "R-PTS" Then
                                                Stop
                                                mtFamily = "D-BAR"
                                            End If
                                            If mtFamily = "DSHEET" Then
                                                nmFamily = "D-RMT"
                                                qtUnit = "FT2"
                                            ElseIf mtFamily = "D-BAR" Then
                                                If Len(nmFamily) = 0 Then
                                                    nmFamily = "R-RMT"
                                                Else
                                                    Debug.Print("")
                                                End If
                                                qtUnit = prRmUnit.Text
                                                ck = vbCancel
                                                Do
                                                    If True Then
                                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                                        With oDoc.ComponentDefinition.RangeBox
                                                            Debug.Print("")
                                                            System.Math.Round((.MaxPoint.X - .MinPoint.X) / CvLenIn2cm, 4)
                                                            System.Math.Round((.MaxPoint.Y - .MinPoint.Y) / CvLenIn2cm, 4)
                                                            System.Math.Round((.MaxPoint.Z - .MinPoint.Z) / CvLenIn2cm, 4)
                                                        End With
                                                    End If
                                                    With nuAiBoxData().UsingInches().UsingBox(oDoc.ComponentDefinition.RangeBox)
                                                        Debug.Print(.Dump(0))
                                                    End With
                                                    Debug.Print("qtRawMatl = ", CStr(qtRawMatl), " 'in model. ")
                                                    If dcIn.Exists(PnRmQty) Then Debug.Print("In Genius: ", CStr(dcIn.Item(PnRmQty)))
                                                    Debug.Print("")
                                                    Debug.Print("qtUnit = """, qtUnit, """ 'in model.")
                                                    If dcIn.Exists(PnRmUnit) Then Debug.Print("In Genius: ", CStr(dcIn.Item(PnRmUnit)))
                                                    Debug.Print(" ( or try IN )")
                                                    Debug.Print("")
                                                    With nu_fmIfcMatlQty01().SeeUser(oDoc)
                                                        If .Exists(PnRmQty) Then
                                                            If CDbl("0" & CStr(qtRawMatl)) = CDbl(.Item(PnRmQty)) Then
                                                            Else
                                                                Debug.Print("qtRawMatl FROM " & qtRawMatl & " TO " & .Item(PnRmQty))
                                                                qtRawMatl = .Item(PnRmQty)
                                                            End If
                                                        Else
                                                            Stop
                                                        End If
                                                        If .Exists(PnRmUnit) Then
                                                            If qtUnit = .Item(PnRmUnit) Then
                                                            Else
                                                                Debug.Print("qtUnit FROM " & qtUnit & " TO " & .Item(PnRmUnit))
                                                                qtUnit = .Item(PnRmUnit)
                                                            End If
                                                        Else
                                                            Stop
                                                        End If
                                                    End With
                                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(qtRawMatl), qtUnit, ". IF OKAY, CONTINUE.")
                                                    ck = newFmTest2().AskAbout(oDoc, "Raw Material Quantity is now " & CStr(qtRawMatl) & qtUnit & " for", "If this is okay, click [YES]." & vbCrLf & "Otherwise, click [NO] to review." & vbCrLf & "" & vbCrLf & "( for debug, click [CANCEL] )")
                                                    If ck = vbCancel Then
                                                        Stop
                                                    End If
                                                Loop Until ck = vbYes
                                                prRmQty.Text = qtRawMatl
                                                rt = dcAddProp(prRmQty, rt)
                                                Debug.Print("")
                                            Else
                                                Stop
                                                nmFamily = ""
                                                qtUnit = ""
                                            End If
                                        End If
                                    End With
                                End If
                            End If
                        Else
                            If .DocumentInterests.HasInterest(GuidPipingSgmt) Then
                                ck = newFmTest2().AskAbout(oDoc, "", Join({"" _
                                    , "appears to be Hose or Tubing," _
                                    , "presently " & IIf(Len(pnStock) > 0, pnStock, "un") & ".", "" _
                                    , "Would you like to " & IIf(Len(pnStock) > 0, "change", "") & " it?"
                                }, vbCrLf))
                                If ck = vbCancel Then
                                    Stop
                                ElseIf ck = vbYes Then
                                    pnStock = userChoiceFromDc(DcFrom2Fields(CnGnsDoyle().Execute(SqlOf_GnsTubeHose(.ComponentDefinition.Parameters.Item("Size_Designation").Text)), "Description", "Item"), pnStock)
                                    qtUnit = Trim$(UCase$(aiPropsUser.Item("ROPL").Text))
                                    qtRawMatl = System.Math.Round(Val(Split(qtUnit & " ", " ")(0)), 4)
                                    qtUnit = Split(qtUnit & " ", " ")(1)
                                    ck = newFmTest2().AskAbout(oDoc, Join({"Stock Quantity of ", qtRawMatl & qtUnit & " of " & pnStock, "selected for Item "}, vbCrLf), Join({"If this is okay, click [YES]", "(CANCEL to debug)"}, vbCrLf))
                                    If ck = vbCancel Then
                                    ElseIf ck = vbYes Then
                                        prRawMatl.Text = pnStock
                                        prRmQty.Text = qtRawMatl
                                        prRmUnit.Text = qtUnit
                                        Debug.Print("")
                                    Else
                                        Stop
                                    End If
                                    Debug.Print("")
                                End If
                            End If
                            With newFmTest1()
                                If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop
                                bd = nuAiBoxData().UsingInches.SortingDims(oDoc.ComponentDefinition.RangeBox)
                                ck = .AskAbout(oDoc, "Please Select Stock for Machined Part" & vbCrLf & vbCrLf & bd.Dump(0))
                                If ck = vbYes Then
                                    With .ItemData
                                        If .Exists(PnFamily) Then
                                            nmFamily = .Item(PnFamily)
                                            Debug.Print(PnFamily & "=" & nmFamily)
                                        End If
                                        If .Exists(PnRawMaterial) Then
                                            pnStock = .Item(PnRawMaterial)
                                            Debug.Print(PnRawMaterial & "=" & pnStock)
                                        End If
                                    End With
                                    If 0 Then Stop
                                Else
                                    Debug.Print("")
                                End If
                            End With
                            If Len(pnStock) > 0 Then
                                With CnGnsDoyle().Execute("select Family from vgMfiItems where Item='" & Replace(pnStock, "'", "''") & "';")
                                    If .BOF Or .EOF Then
                                        Stop
                                    Else
                                        With .Fields
                                            mtFamily = .Item("Family").Value
                                        End With
                                        If mtFamily Like "?-MT*" Then
                                            Debug.Print(pnModel & "[" & qtRawMatl & qtUnit & " of " & pnStock & ": " & aiPropsDesign.item(PnDesc).Text & "]")
                                            Stop
                                        ElseIf mtFamily Like "?-PartDoc*" Then
                                            If nmFamily Like "?-RM*" Then
                                                Debug.Print("")
                                            Else
                                                ck = MsgBox(Join({
                                                    "Part " & pnModel & " uses " & pnStock,
                                                    "which is not sheet metal.",
                                                    "",
                                                    "These parts are usually assigned",
                                                    "to the Riverview family, R-RMT.",
                                                    "",
                                                    "Do you want to use this Family?",
                                                    "Click [NO] to see other options.",
                                                    "(CANCEL to debug)"
                                                }, vbCrLf), vbYesNoCancel + vbQuestion, "Select Part Family?")
                                                Debug.Print("")
                                                If ck = vbCancel Then
                                                    Stop
                                                ElseIf ck = vbYes Then
                                                    nmFamily = "R-RMT"
                                                Else
                                                    If Len(nmFamily) = 0 Then
                                                        nmFamily = "R-RMT"
                                                    End If
                                                    With NuDcPopulator().Setting("D-RMT", "Doyle (typ. sheet metal)").Setting("R-RMT", "Riverview (most others)")
                                                        If Not .Exists(nmFamily) Then
                                                            .Setting(nmFamily, "Current (" & nmFamily & ")")
                                                        End If
                                                        nmFamily = userChoiceFromDc(DcTransposed(.Dictionary()), nmFamily)
                                                    End With
                                                End If
                                            End If
                                            mtFamily = "D-BAR"
                                        ElseIf mtFamily = "D-PTS" Then
                                            mtFamily = "D-BAR"
                                            Stop
                                        ElseIf mtFamily = "R-PTS" Then
                                            mtFamily = "D-BAR"
                                            Stop
                                        End If
                                        If mtFamily = "DSHEET" Then
                                            Stop
                                            nmFamily = "D-RMT"
                                            qtUnit = "FT2"
                                        ElseIf mtFamily = "D-BAR" Then
                                            nmFamily = "R-RMT"
                                            qtUnit = prRmUnit.Text
                                            ck = vbCancel
                                            Do
                                                If True Then
                                                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                                    With oDoc.ComponentDefinition.RangeBox
                                                        Debug.Print("")
                                                        System.Math.Round((.MaxPoint.X - .MinPoint.X) / CvLenIn2cm, 4)
                                                        System.Math.Round((.MaxPoint.Y - .MinPoint.Y) / CvLenIn2cm, 4)
                                                        System.Math.Round((.MaxPoint.Z - .MinPoint.Z) / CvLenIn2cm, 4)
                                                    End With
                                                End If
                                                With nuAiBoxData().UsingInches().UsingBox(oDoc.ComponentDefinition.RangeBox)
                                                    Debug.Print(.Dump(0))
                                                End With
                                                Debug.Print("qtRawMatl = ", CStr(qtRawMatl), " 'in model. ")
                                                If dcIn.Exists(PnRmQty) Then Debug.Print("In Genius: ", CStr(dcIn.Item(PnRmQty)))
                                                Debug.Print("")
                                                Debug.Print("qtUnit = """, qtUnit, """ 'in model.")
                                                If dcIn.Exists(PnRmUnit) Then Debug.Print("In Genius: ", CStr(dcIn.Item(PnRmUnit)))
                                                Debug.Print(" ( or try IN )")
                                                Debug.Print("")
                                                With nu_fmIfcMatlQty01().SeeUser(oDoc)
                                                    If .Exists(PnRmQty) Then
                                                        If CDbl("0" & CStr(qtRawMatl)) = CDbl(.Item(PnRmQty)) Then
                                                        Else
                                                            Debug.Print("qtRawMatl FROM " & qtRawMatl & " TO " & .Item(PnRmQty))
                                                            qtRawMatl = .Item(PnRmQty)
                                                        End If
                                                    Else
                                                        Stop
                                                    End If
                                                    If .Exists(PnRmUnit) Then
                                                        If qtUnit = .Item(PnRmUnit) Then
                                                        Else
                                                            Debug.Print("qtUnit FROM " & qtUnit & " TO " & .Item(PnRmUnit))
                                                            qtUnit = .Item(PnRmUnit)
                                                        End If
                                                    Else
                                                        Stop
                                                    End If
                                                End With
                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(qtRawMatl), qtUnit, ". IF OKAY, CONTINUE.")
                                                ck = newFmTest2().AskAbout(oDoc, "Raw Material Quantity is now " & CStr(qtRawMatl) & qtUnit & " for", "If this is okay, click [YES]." & vbCrLf & "Otherwise, click [NO] to review." & vbCrLf & "" & vbCrLf & "( for debug, click [CANCEL] )")
                                                If ck = vbCancel Then
                                                    Stop
                                                End If
                                            Loop Until ck = vbYes
                                            prRmQty.Text = qtRawMatl
                                            rt = dcAddProp(prRmQty, rt)
                                            Debug.Print("")
                                        Else
                                            Stop
                                            nmFamily = ""
                                            qtUnit = ""
                                        End If
                                    End If
                                End With
                            End If
                        End If

                        If Len(pnStock) > 0 Then
                            With prRawMatl
                                If Len(Trim$(.Text)) > 0 Then
                                    If pnStock <> .Text Then
                                        ck = MsgBox(Join({
                                            "Raw Stock Change Suggested",
                                            "  for Item " & pnModel,
                                            "",
                                            "  Current : " & prRawMatl.Text,
                                            "  Proposed: " & pnStock,
                                            "", "Change It?", ""
                                        }, vbCrLf), vbYesNo, pnModel & " Stock")
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            .Text = pnStock
                                        End If
                                    End If
                                Else
                                    .Text = pnStock
                                End If
                            End With
                        End If
                        rt = dcAddProp(prRawMatl, rt)

                        With prRmUnit
                            If Len(.Text) > 0 Then
                                If Len(qtUnit) > 0 Then
                                    If .Text <> qtUnit Then
                                        ck = newFmTest2().AskAbout(oDoc, , "Raw Material " & prRawMatl.Text & vbCrLf & "Unit of Measure currently " & .Text & vbCrLf & vbCrLf & "Change to " & qtUnit & "?" & vbCrLf & " ")
                                        If ck = vbCancel Then
                                            Stop
                                        ElseIf ck = vbYes Then
                                            .Text = qtUnit
                                        End If
                                        If 0 Then Stop
                                    End If
                                End If
                            Else
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        Debug.Print("")
                    ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                        If Len(nmFamily) = 0 Then
                            If 1 Then Stop
                            nmFamily = "D-PTS"
                        End If
                    ElseIf bomStruct = BOMStructureEnum.kPhantomBOMStructure Then
                        ck = newFmTest2().AskAbout(oDoc, "For some reason, THIS Item is marked Phantom:", "Is this okay? (Click [NO] OR [CANCEL] if not)")
                        If ck = vbYes Then
                            'just let it go
                        Else
                            Stop
                        End If
                    ElseIf bomStruct = BOMStructureEnum.kInseparableBOMStructure Then
                        ck = newFmTest2().AskAbout(oDoc, "This Item is marked Inseperable:", Join({"This is likely not correct,", "and should be fixed ASAP.", "Would you like to copy the Part", "Number for later review?", "", vbCrLf & vbCrLf & "([CANCEL] to debug)", " "}))
                        If ck = vbYes Then
                            InputBox(Join({"Copy this Part Number, and paste", "it into another document or memo", "for later review."}, vbCrLf), "Copy Part Number " & pnModel, pnModel)
                        ElseIf ck = vbCancel Then
                            Stop
                        End If
                        Stop
                    Else
                        ck = newFmTest2().AskAbout(oDoc, "The following Item has an unhandled BOM Structure:", "Skip it? (Click [NO] OR [CANCEL] to review)")
                        If ck = vbYes Then
                            'just let it go
                        Else
                            Stop
                        End If
                        Stop
                    End If

                    If oDoc.ComponentDefinition.IsContentMember Then
                        ' Don't muck around with the Family!
                    Else
                        If Len(nmFamily) > 0 Then
                            If prFamily.Text <> nmFamily Then
                                On Error Resume Next
                                prFamily.Text = nmFamily
                                If Err.Number Then
                                    Debug.Print("CHGFAIL[FAMILY]{'" & prFamily.Text & "' -> '" & nmFamily & "'}: " & oDoc.DisplayName & " (" & oDoc.FullDocumentName & ")")
                                    If MsgBox("Couldn't Change Family" & vbCrLf & "for Item " & oDoc.DisplayName & vbCrLf & vbCrLf & "(" & oDoc.FullDocumentName & ")" & vbCrLf & vbCrLf & "Stop to Review?", vbYesNo Or vbDefaultButton2, oDoc.DisplayName) = vbYes Then
                                        Stop
                                    End If
                                End If
                                On Error GoTo 0
                            End If
                            rt = dcAddProp(prFamily, rt)
                        End If
                    End If
                End With

                Call iSyncPartFactory(oDoc)
                dcGeniusPropsPartRev20180530 = rt
            End With
        End If
    End Function

    Public Function dcGeniusPropsPartRev20200409(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        '
        ' dcGeniusPropsPartRev20200409
        ' [2020.04.09] begin new revision
        '
        Dim rt As Scripting.Dictionary
        ''
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        ''
        Dim prFamily As Inventor.Property
        Dim prRawMatl As Inventor.Property 'pnRawMaterial
        Dim prRmUnit As Inventor.Property 'pnRmUnit
        Dim prRmQty As Inventor.Property 'pnRmQty
        ''
        Dim nmFamily As String
        Dim mtFamily As String
        ' UPDATE[2018.05.30.01]
        Dim pnStock As String
        Dim qtUnit As String
        Dim bomStruct As Inventor.BOMStructureEnum
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData

        If dc Is Nothing Then
            dcGeniusPropsPartRev20200409 =
        dcGeniusPropsPartRev20200409(
            oDoc, New Scripting.Dictionary
        )
        Else
            rt = dc

            With oDoc
                ' Get Property s
                With .Propertys
                    aiPropsUser = .Item(gnCustom)
                    aiPropsDesign = .Item(gnDesign)
                End With

                ' Get Custom Properties
                prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)

                ' Family property is from Design, NOT Custom 
                prFamily = aiGetProp(aiPropsDesign, pnFamily)

                ' We should check HERE for possibly misidentified purchased parts
                ' UPDATE[2018.02.06.01]: Using new UserForm, see below
                With .ComponentDefinition
                    ' Request #1: Get the Mass in Pounds
                    ' and add to Custom Property GeniusMass
                    With .MassProperties
                        rt = dcWithProp(
                        aiPropsUser, pnMass,
                        System.Math.Round(cvMassKg2LbM * .Mass, 4), rt
                    )
                    End With

                    '
                    ' Get BOM Structure type, correcting if appropriate,
                    ' and prepare Family value for part, if purchased.
                    '
                    ' NOTE[2020.04.09.01]
                    ck = vbNo
                    ' UPDATE[2018.05.31.01]
                    If InStr(1, oDoc.FullFileName,
                    "\Doyle_Vault\Designs\purchased\"
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                    "|" & prFamily.Text & "|"
                ) > 0 Then
                        ' UPDATE[2020.04.09.02]
                        ck = newFmTest2().AskAbout(oDoc, ,
                        "Is this a Purchased Part?"
                    )
                    End If

                    ' Check process below replaces duplicate check/responses above.
                    ' NOTE[2020.04.09.02]
                    If ck = vbYes Then
                        If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                            On Error Resume Next
                            .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                            If Err.Number = 0 Then
                                bomStruct = .BOMStructure
                            Else
                                bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                                ' WARNING: NOT a good way to go about this
                                '     but will go with it for now
                            End If
                            On Error GoTo 0
                        Else
                            bomStruct = .BOMStructure 'to make sure this is captured
                        End If
                    Else
                        bomStruct = .BOMStructure 'to make sure this is captured
                    End If

                    'Request #2: Change Cost Center iProperty.
                    'If BOMStructure = Purchased and not content center,
                    'then Family = D-PTS, else Family = D-HDWR.
                    '
                    'UPDATE[2018.05.30.02]
                    If bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                        If .IsContentMember Then
                            nmFamily = "D-HDWR"
                        Else
                            nmFamily = "D-PTS"
                            ' NOTE[2020.04.09.03]
                        End If
                    Else
                        nmFamily = ""
                    End If
                End With
                ' At this point, nmFamily SHOULD be 
                ' to a non-blank value if Item is purchased.
                ' We should be able to check this later on,
                ' if Item BOMStructure is NOT Normal

                ' UPDATE[2020.04.09.03]
                If bomStruct = BOMStructureEnum.kNormalBOMStructure Then
                    '----------------------------------------------------'
                    If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                        '----------------------------------------------------'
                        '
                        ' NOTE[2018.05.31.01]
                        'Request #3:
                        '   Get sheet metal extent area
                        '   and add to custom property "RMQTY"
                        rt = dcFlatPatProps(.ComponentDefinition, rt)
                        ' NOTE[2018.05.30.01]

                        'Moved to start of block to check for NON sheet metal

                        'NOTE: THIS call might best be combined somehow
                        '   with the flat pattern prop pickup above.
                        '   Note especially that if dcFlatPatProps
                        '   FINDS NO .FlatPattern, then there should
                        '   BE NO sheet metal part number!
                        If prRawMatl Is Nothing Then
                            Stop ' UPDATE[2020.04.09.04]
                            If rt.Exists("OFFTHK") Then
                                ' UPDATE[2018.05.30.05]
                                Debug.Print(aiProperty(rt.Item("OFFTHK")).Text)
                                Stop 'because we're going to need to do something with this.

                                pnStock = "" 'Originally the ONLY line in this block.
                                ' A more substantial response is required here.

                                If 0 Then Stop '(just a skipover)
                            Else
                                Stop 'because we don't know IF this is sheet metal yet
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If
                        Else
                            ' NOTE[2018.09.14.01]: ACTION ADVISED
                            If Len(prRawMatl.Text) > 0 Then
                                pnStock = prRawMatl.Text
                            Else
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If

                            If Len(pnStock) = 0 Then
                                ' UPDATE[2018.05.30.03]
                                With newFmTest1()
                                    If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop

                                    bd = nuAiBoxData().UsingInches.SortingDims(
                                    oDoc.ComponentDefinition.RangeBox
                                )
                                    ck = .AskAbout(oDoc,
                                    "No Stock Found! Please Review" _
                                    & vbCrLf & vbCrLf & bd.Dump(0)
                                )

                                    If ck = vbYes Then
                                        ' UPDATE[2018.05.30.04]
                                        With .ItemData
                                            If .Exists(pnFamily) Then
                                                nmFamily = .Item(pnFamily)
                                                Debug.Print(pnFamily & "=" & nmFamily)
                                            End If

                                            If .Exists(pnRawMaterial) Then
                                                pnStock = .Item(pnRawMaterial)
                                                Debug.Print(pnRawMaterial & "=" & pnStock)
                                            End If
                                        End With
                                        If 0 Then Stop 'Use this for a debugging shim
                                    End If
                                End With
                            End If

                            If Len(pnStock) > 0 Then 'and ONLY then
                                'do we look for a Raw Material Family!

                                With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                    If .BOF Or .EOF Then
                                        Stop 'because Material value likely invalid
                                        ' NOTE[2018.09.14.02]: ACTION ADVISED
                                    Else
                                        With .Fields
                                            mtFamily = .Item("Family").Value
                                        End With

                                        If mtFamily = "DSHEET" Then
                                            'We should be okay. This is sheet metal stock
                                            nmFamily = "D-RMT"
                                            qtUnit = "FT2"
                                            ' UPDATE[2018.05.30.06]
                                        ElseIf mtFamily = "D-BAR" Then
                                            nmFamily = "R-RMT"
                                            qtUnit = prRmUnit.Text '"IN"
                                            ''may want function here
                                            ' UPDATE[2018.05.30.07]
                                            Debug.Print(aiPropsDesign.Item(pnPartNum).Text, " [", prRawMatl.Text, "]: ", aiPropsDesign.Value(pnDesc).Text)
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,")
                                            Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                                            Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                            With oDoc.ComponentDefinition.RangeBox
                                                Debug.Print("")
                                                Debug.Print(
                                                    (.MaxPoint.X - .MinPoint.X) / 2.54,
                                                    (.MaxPoint.Y - .MinPoint.Y) / 2.54,
                                                    (.MaxPoint.Z - .MinPoint.Z) / 2.54
                                                )
                                            End With
                                            ' UPDATE[2020.04.09.05]
                                            Debug.Print("")
                                            Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text))
                                            ' UPDATE[2020.04.09.05]
                                            Debug.Print("qtUnit = ""IN""")
                                            ' UPDATE[2020.04.09.05]
                                            Stop 'because we might want a D-BAR handler
                                            ' UPDATE[2020.04.09.05]
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                            Stop
                                            rt = dcAddProp(prRmQty, rt)
                                            Debug.Print("") 'Landing line for debugging. Do not disable.
                                        Else
                                            nmFamily = ""
                                            qtUnit = "" 'may want function here
                                            ' UPDATE[2018.05.30.08]
                                            Stop 'because we don't know WHAT to do with it
                                        End If
                                    End If
                                End With
                            Else
                                If 0 Then Stop 'and regroup
                                ' Things are looking a right royal mess
                                ' at the moment I'm writing this comment.
                            End If
                        End If

                        With prRawMatl
                            If Len(Trim$(.Text)) > 0 Then
                                If pnStock <> .Text Then
                                    ' UPDATE[2020.04.09.06]
                                    ck = MsgBox(
                                    Join({
                                        "Raw Stock Change Suggested",
                                        "  Current : " & prRawMatl.Text,
                                        "  Proposed: " & pnStock,
                                        "", "Change It?", ""
                                    }, vbCrLf),
                                    vbYesNo, "Change Raw Material?"
                                )
                                    '"Suggested Sheet Metal"
                                    If ck = vbYes Then .Text = pnStock
                                End If
                            Else
                                .Text = pnStock
                            End If
                        End With
                        rt = dcAddProp(prRawMatl, rt)

                        With prRmUnit
                            If Len(.Text) > 0 Then
                                If Len(qtUnit) > 0 Then
                                    If .Text <> qtUnit Then
                                        Stop 'and check both so we DON'T
                                        'automatically "fix" the RMUNIT value

                                        .Text = qtUnit

                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        ' UPDATE[2020.04.09.07]
                        Debug.Print("") 'Another landing line

                        '--------------------------------------------'
                    Else 'for standard Part (NOT Sheet Metal) ---'
                        '--------------------------------------------'
                        ' NOTE[2018.07.31.01]
                        With newFmTest1()
                            If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop

                            ' [2018.07.31.02][by AT]
                            bd = nuAiBoxData().UsingInches.SortingDims(
                                oDoc.ComponentDefinition.RangeBox
                            )

                            ck = .AskAbout(oDoc,
                                "Please Select Stock for Machined Part" _
                                & vbCrLf & vbCrLf & bd.Dump(0)
                            )

                            If ck = vbYes Then
                                ' UPDATE[2018.05.30.09]
                                With .ItemData
                                    If .Exists(pnFamily) Then
                                        nmFamily = .Item(pnFamily)
                                        Debug.Print(pnFamily & "=" & nmFamily)
                                    End If

                                    If .Exists(pnRawMaterial) Then
                                        pnStock = .Item(pnRawMaterial)
                                        Debug.Print(pnRawMaterial & "=" & pnStock)
                                    End If
                                End With
                                If 0 Then Stop 'Use this for a debugging shim
                                ' NOTE[2020.04.09.04]
                            End If
                        End With
                        '
                        '
                        '

                        ' NOTE[2020.04.09.05]
                        If Len(pnStock) > 0 Then 'and ONLY then
                            'do we look for a Raw Material Family!

                            ' NOTE[2020.04.09.06]
                            With CnGnsDoyle().Execute(
                                "select Family " &
                                "from vgMfiItems " &
                                "where Item='" & pnStock & "';"
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    ' NOTE[2020.04.09.07]
                                End If
                            End With
                            ' These closing statements moved up from below following If block
                            '

                            'mtFamily = nmFamily 'to force "correct" behavior of following section
                            If mtFamily = "DSHEET" Then
                                Stop 'because we should NOT be doing Sheet Metal in this section.
                                ' This might require further investigation and/or development, if encountered.
                                'We should be okay. This is sheet metal stock
                                nmFamily = "D-RMT"
                                qtUnit = "FT2"
                                ' UPDATE[2018.05.30.06]
                            ElseIf mtFamily = "D-BAR" Then
                                nmFamily = "R-RMT"
                                qtUnit = prRmUnit.Text '"IN"
                                ''may want function here
                                ' UPDATE[2018.05.30.07]
                                Debug.Print(aiPropsDesign.Item(pnPartNum).Text, " [", prRawMatl.Text, "]: ", aiPropsDesign.Value(pnDesc).Text)
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF CHANGE NEEDED,")
                                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,")
                                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.")
                                Debug.Print("X SPAN", "Y SPAN", "Z SPAN")
                                Debug.Print((oDoc.ComponentDefinition.RangeBox.MaxPoint.X - oDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Y - oDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (oDoc.ComponentDefinition.RangeBox.MaxPoint.Z - oDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54)
                                Debug.Print("")
                                Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.")
                                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.")
                                Debug.Print("")
                                Debug.Print("prRmQty.Text = ", CStr(prRmQty.Text))
                                Debug.Print("qtUnit = ""IN""")
                                Debug.Print("")
                                Stop 'because we might want a D-BAR handler
                                ' UPDATE[2020.04.09.05]
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ", CStr(prRmQty.Text), qtUnit, ". IF OKAY, CONTINUE.")
                                Stop
                                rt = dcAddProp(prRmQty, rt)
                                Debug.Print("") 'Landing line for debugging. Do not disable.
                            Else
                                nmFamily = ""
                                qtUnit = "" 'may want function here
                                ' UPDATE[2018.05.30.08]
                                Stop 'because we don't know WHAT to do with it
                            End If
                        Else
                            If 0 Then Stop 'and regroup
                            ' Things are looking a right royal mess
                            ' at the moment I'm writing this comment.
                        End If


                        With prRawMatl
                            If Len(Trim$(.Text)) > 0 Then
                                If pnStock <> .Text Then
                                    'Debug.Print("Raw Stock Selection"
                                    'Debug.Print("  Current : " & prRawMatl.Text
                                    'Debug.Print("  Proposed: " & pnStock
                                    'Stop 'because we might not want to change existing stock ting
                                    'if
                                    ck = MsgBox(
                                    Join({
                                        "Raw Stock Change Suggested",
                                        "  Current : " & prRawMatl.Text,
                                        "  Proposed: " & pnStock,
                                        "", "Change It?", ""
                                    }, vbCrLf),
                                    vbYesNo, "Change Raw Material?"
                                )
                                    '"Suggested Sheet Metal"
                                    If ck = vbYes Then .Text = pnStock
                                End If
                            Else
                                .Text = pnStock
                            End If
                        End With
                        rt = dcAddProp(prRawMatl, rt)

                        With prRmUnit
                            If Len(.Text) > 0 Then
                                If Len(qtUnit) > 0 Then
                                    If .Text <> qtUnit Then
                                        Stop 'and check both so we DON'T
                                        'automatically "fix" the RMUNIT value

                                        .Text = qtUnit

                                        If 0 Then Stop 'Ctrl-9 here to skip changing
                                    End If
                                End If
                            Else 'we're ting a new quantity unit
                                .Text = qtUnit
                            End If
                        End With
                        rt = dcAddProp(prRmUnit, rt)
                        '
                        '
                        '
                    End If 'Sheetmetal vs Part
                ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then
                    ' As mentioned above, nmFamily
                    ' SHOULD be  at this point
                    If Len(nmFamily) = 0 Then
                        If 1 Then Stop 'because we might
                        'need to check out the situation
                        nmFamily = "D-PTS" 'by default
                    End If
                Else
                    Stop 'because we might need
                    'to do something else
                    'based on an unexpected
                    'BOM Structure
                End If

                ' Get the design tracking property ,
                ' and update the Cost Center Property
                If oDoc.ComponentDefinition.IsContentMember Then
                    ' Don't muck around with the Family!
                Else
                    If Len(nmFamily) > 0 Then
                        prFamily.Text = nmFamily
                        rt = dcAddProp(prFamily, rt)
                        ' rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                    End If
                End If
            End With

            Call iSyncPartFactory(oDoc) 'Backport Properties to iPart Factory
            dcGeniusPropsPartRev20200409 = rt
        End If
    End Function
    ' NOTE[2018.05.30.01]:
    '     Raw Material Quantity value
    '     SHOULD be  upon return
    '     We may need to review the process
    '     to find an appropriate place
    '     to  for NON sheet metal
    ' NOTE[2018.05.31.01]:
    '     At this point, we MAY wish
    '     to check for a valid flat pattern,
    '     and otherwise attempt to verify
    '     an actual sheet metal design.
    ' NOTE[2018.07.31.01][by AT]
    '     Duped following block from above
    '     to mod for material assignment
    '     to non-sheet metal part.
    '
    '     Except, this isn't enough.
    '     Also need the code to add
    '     Stock PN to Attribute RM.
    '     That's a whole 'nother
    '     block of code, and likely
    '     best consolidated.
    ' [2018.07.31.02][by AT]
    '     Added the following to try to
    '     preselect non-sheet metal stock
    '     [and then disabled the following]
    '.dbFamily.Text = "D-BAR"
    '.LbxFamily.Text = "D-BAR"
    ' Doesn't quite do it.
    'With New aiBoxData
    ' bd = nuAiBoxData().UsingInches.UsingBox( _
    '    oDoc.ComponentDefinition.RangeBox _
    ')
    '
    'End With
    ' NOTE[2018.09.14.01]: ACTION ADVISED
    '     pnStock can probably be  to prRawMatl.Text
    '     and THEN checked for length to see if lookup needed.
    '     This might also allow us to check for machined
    '     or other non-sheet metal parts.
    ' NOTE[2018.09.14.02]: ACTION ADVISED
    '     Will need to address this situation
    '     in a more robust manner.
    '     A more thorough query above
    '     might also be called for.
    ' NOTE[2020.04.09.01]: This section should check
    '     for Purchased Part status in Genius, as well
    '     as the checks below. BOM Structure should also
    '     be checked, but TING it eventually needs
    '     to be shifted to a subsequent operation.
    ' NOTE[2020.04.09.02]:
    '     this is where Document's BOMStructure
    '     is . should be moved to a later stage
    ' NOTE[2020.04.09.03]:
    '     [original date unknown]
    '     NON Content Center members
    '     might still be D-HDWR
    '     Additional checks might
    '     be recommended
    ' NOTE[2020.04.09.04]
    '     [original date unknown]
    '     We're going to need something here
    '     to make sure raw material gets added
    '     for non sheet metal parts, as well
    '     What we're going to need to do
    '     is refactor this whole bloody thing.
    ' NOTE[2020.04.09.05]
    '     [original date unknown]
    '
    '     The following If block is copied
    '     wholesale from sheet metal section above.
    '     Some changes (to be) made to accommodate
    '     machined or other non-sheet metal stock.
    '
    '     Ultimately, whole mess to require refactor.
    '
    ' NOTE[2020.04.09.06]
    '     [original date unknown]
    '     This enclosing With block should NOT be necessary
    '     since the newFmTest1 above takes care of collecting
    '     the Stock Family along with the Stock itself
    '
    ' NOTE[2020.04.09.07]
    '     [original date unknown]
    '
    ' Content formerly here moved BELOW and OUT of this section
    ' as it should only require results of newFmTest1 exchange above
    '
    ''
    ''
    ''
    '
    ' UPDATE[2018.05.30.01]:
    '     Rename variable Family to nmFamily
    '     to minimize confusion between code
    '     and comment text in searches.
    '     Also add variable mtFamily
    '     for raw material Family name
    ' UPDATE[2018.05.30.02]:
    '     Value produced here
    '     will now be held for later processing,
    '     more toward the end of this function.
    ' UPDATE[2018.05.30.03]:
    '     Pulling ALL code/text from this section
    '     to get rid of excessive cruft.
    '
    '     In fact, reversing logic to go directly
    '     to User Prompt if no stock identified
    '
    '     IN DOUBLE FACT, hauling this WHOLE MESS
    '     RIGHT UP after initial pnStock assignment
    '     to prompt user IMMEDIATELY if no stock found
    ' UPDATE[2018.05.30.04]:
    '     Pulling some extraneous commented code
    '     from here and beginning of block
    ' UPDATE[2018.05.30.05]:
    '     Restoring original key check
    '     and adding code for debug
    '     Previously changed to "~OFFTHK"
    '     to avoid this block and its issues.
    '     (Might re-revert if not prepped to fix now)
    ' UPDATE[2018.05.30.06]: (two locations)
    '     Moving part family assignment
    '     to this section for better mapping
    '     and updating to new Family names
    '     as well as pulling up qtUnit assignment
    ' UPDATE[2018.05.30.07]: (two locations)
    '     As noted above
    '     Will keep Stop for now
    '     pending further review,
    '     hopefully soon
    ' UPDATE[2018.05.30.08]: As noted above
    '     However, might need more handling here.
    ' UPDATE[2018.05.30.09]:
    '     Pulling some extraneous commented code
    '     from here and beginning of block
    ' UPDATE[2018.05.31.01]:
    '     Combined both InStr checks
    '     by addition to generate a single test for > 0
    '     If EITHER string match succeeds, the total
    '     SHOULD exceed zero, so this SHOULD work.
    ' UPDATE[2018.02.06.01]:
    '     Using new UserForm
    ' UPDATE[2020.04.09.02]:
    '     Remove disabled/outdated code as follows
    ' UPDATE[2018.02.06]: Using same
    '     new UserForm as noted above.
    'ck = newFmTest2().AskAbout(oDoc, , _
    '        "Is this a Purchased Part?" _
    '    )
    ''ElseIf InStr(1, _
    '    "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
    '    "|" & prFamily.Text & "|" _
    ') > 0 Then
    ' UPDATE[2018.02.06]: Using same
    '     new UserForm as noted above.
    ' UPDATE[2020.04.09.03]:
    '     Removed disabled/outdated code as follows
    'Request #4: Change Cost Center iProperty.
    'If BOMStructure = Normal, then Family = D-MTO,
    'else if BOMStructure = Purchased then Family = D-PTS.
    ' UPDATE[2020.04.09.04]:
    '     Adding Stop here to see if prRawMatl
    '     ever comes up missing inside a sheet metal part
    ' UPDATE[2020.04.09.05]: (multiple points)
    '     Removing disabled/obsolete code as follows
    'Debug.Print("CURRENT RAW MATERIAL QUANTITY (";
    'Debug.Print(CStr(prRmQty.Text), ") IS SHOWN BELOW."
    'Debug.Print("IF NOT CORRECT, YOU MAY TYPE A NEW VALUE"
    'Debug.Print("IN ITS PLACE, AND PRESS ENTER TO CHANGE IT."
    'Debug.Print("SOME SUGGESTED VALUES INCLUDE X, Y, AND Z"
    'Debug.Print("EXTENTS (ABOVE) OR YOU MAY SUPPLY YOUR OWN."
    'Debug.Print(""
    'Debug.Print(""
    'Debug.Print("YOU MAY ALSO CHANGE THE UNIT OF MEASURE BELOW,"
    'Debug.Print("IF DESIRED. BE SURE TO PRESS ENTER/RETURN"
    'Debug.Print("AFTER CHANGING EITHER LINE. WHEN FINISHED, "
    'Debug.Print("PRESS [F5] TO CONTINUE."
    '
    'Debug.Print("qtUnit = """, qtUnit, """"
    '
    'Debug.Print(""
    'Debug.Print(""
    'Debug.Print(""
    '
    ' Actually, we might NOT need to stop here
    ' if bar stock is already selected,
    ' because quantities would presumably
    ' have been established already.
    ' Any D-BAR handler probably needs
    ' to be implemented in prior section(s)
    '
    ' UPDATE[2020.04.09.06]:
    '     Removing disabled/obsolete code as follows
    'Debug.Print("Raw Stock Selection"
    'Debug.Print("  Current : " & prRawMatl.Text
    'Debug.Print("  Proposed: " & pnStock
    'Stop 'because we might not want to change existing stock ting
    'if
    ' UPDATE[2020.04.09.07]:
    '     Removing disabled/obsolete code as follows
    ' rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
    ' Plan to remove commented line above,
    ' superceded by the one above that

    Public Function dcGnsPartProps(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        ' NOTES[2021.03.12]
        ' Don't recall when this function block was created.
        ' Probably around 2020.04.09, with the generation
        ' of function dcGeniusPropsPartRev20200409, above.
        '
        ' As of this writing, no code present, so will
        ' use this to rebuild the Part Properties retrieval
        ' function more or less from the ground up.
        '
        ' Primary goal: reconstruct the basic process
        ' as faithfully as possible, but in a NONdestructive
        ' manner. That is, avoid changing the Part Document
        ' in any way, but simply collect as much information
        ' as is available, and generate whatever else is needed,
        ' and possible WITHOUT altering the Document.
        '
        ' NOTES[2021.03.22]
        ' Following review of process in functions
        ' dcGeniusPropsPartRev20180530 and dcFlatPatProps,
        ' added calls to aiGetProp to retrieve all Property
        ' items checked and/or  by those functions.
        '
        ' Again, this function should NOT attempt to create
        ' any missing/nonexistent Property items, in order
        ' to avoid altering the source Document at this stage.
        '

        '
        '
        Dim rt As Scripting.Dictionary
        ''
        ''  Property s
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        ''
        ''
        ''  Properties
        'Dim prPartNum   As Inventor.Property 'pnPartNum
        'Dim prFamily    As Inventor.Property 'pnFamily
        'Dim prRawMatl   As Inventor.Property 'pnRawMaterial
        'Dim prRmUnit    As Inventor.Property 'pnRmUnit
        'Dim prRmQty     As Inventor.Property 'pnRmQty
        ''
        ''
        ''  Property Values
        Dim pnModel As String
        Dim nmFamily As String
        Dim pnStock As String
        Dim mtFamily As String
        Dim qtUnit As String
        ''
        ''
        ''
        Dim bomStruct As Inventor.BOMStructureEnum
        Dim ck As VbMsgBoxResult
        Dim bd As AiBoxData

        rt = New Scripting.Dictionary
        '

        With oDoc
            ' Get Property s
            With .Propertys
                aiPropsUser = .Item(gnCustom)
                aiPropsDesign = .Item(gnDesign)
            End With
        End With

        With rt
            ' Get Part Number and Family
            ' Properties from Design 
            .Add(pnPartNum, aiGetProp(aiPropsDesign, pnPartNum)) 'prPartNum
            .Add(pnFamily, aiGetProp(aiPropsDesign, pnFamily)) 'prFamily

            ' Get Custom Properties
            .Add(pnRawMaterial, aiGetProp(aiPropsUser, pnRawMaterial)) 'prRawMatl
            .Add(pnRmUnit, aiGetProp(aiPropsUser, pnRmUnit)) 'prRmUnit
            .Add(pnRmQty, aiGetProp(aiPropsUser, pnRmQty)) 'prRmQty
            'NOTE[2021.03.12]: Removed 'create' flag
            'from these function calls to prevent
            'creation of nonexistent Properties,
            'which would alter the source Document.
            'NOTE ALSO: should try to obtain all other
            'custom Properties intended to generate,
            'in case they're already present.

            ' Get Custom Mass/Dimensional Properties
            .Add(pnMass, aiGetProp(aiPropsUser, pnMass)) '<prMass>
            '.Add(pnRmQty, aiGetProp(aiPropsUser, pnRmQty) 'prRmQty
            '   this one already called above
            .Add(pnWidth, aiGetProp(aiPropsUser, pnWidth)) '<prWidth>
            .Add(pnLength, aiGetProp(aiPropsUser, pnLength)) '<prLength>
            .Add(pnArea, aiGetProp(aiPropsUser, pnArea)) '<prArea>
            '.Add("OFFTHK", aiGetProp(aiPropsUser, "OFFTHK") '<prOffThk>
            '   disabled -- not sure if needed any longer
            '   and results in many fewer Prop Dicts
            '   with 'NoVal' Properties

            ' prPartNum = .Item(pnPartNum)
            'pnModel = prPartNum.Text
            ' prFamily = .Item(pnFamily)
            ' prRawMatl = .Item(pnRawMaterial)
            ' prRmUnit = .Item(pnRmUnit)
            ' prRmQty = .Item(pnRmQty)

            Debug.Print("") 'Breakpoint Landing
            'Debug.Print(dumpLsKeyVal(mGr1g0f1(rt), ":", ",")
            'Debug.Print(dumpLsKeyVal(mGr1g0f1(rt))
            'Stop 'Hard
        End With

        '
        dcGnsPartProps = rt
        '
        '
        '
    End Function

    Public Function dcGnsPartsWithProps(
    oDoc As Inventor.Document
) As Scripting.Dictionary
        '
        ' function dcGnsPartsWithProps
        '
        ' returns Dictionary of Dictionaries
        ' containing Genius-related Properties
        ' for each Component of supplied
        ' Inventor Document, be it Part
        ' or Assembly.
        '
        ' NOTE: actual Dictionary processing
        ' removed to separate function
        ' dcGnsPartsWithPropsFromDc
        ' in order to support invocation
        ' from other functions w/o need
        ' for actual source Document
        '
        'Dim dc As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        Dim it As Inventor.PartDocument

        rt = dcGnsPartsWithPropsFromDc(
        dcAiDocComponents(oDoc, , 0)
    )

        Debug.Print("") 'Breakpoint Landing
        dcGnsPartsWithProps = rt
        'Debug.Print(dumpLsKeyVal(mGr1g0f2(dcGnsPartsWithProps(ThisApplication.ActiveDocument)), vbCrLf & vbTab, vbCrLf & vbCrLf)
        'send2clipBd "{" & dumpLsKeyVal(mGr1g0f2(dcGnsPartsWithProps(ThisApplication.ActiveDocument), ": ", "," & vbCrLf & vbTab), "," & vbCrLf & vbTab, vbCrLf & "}," & vbCrLf & "{") & vbCrLf & "}" & vbCrLf
        'send2clipBd ConvertToJson(dcGnsPartsWithProps(ThisApplication.ActiveDocument), " ")
    End Function

    Public Function dcGnsPartsWithPropsFromDc(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' function dcGnsPartsWithPropsFromDc
        '
        ' returns Dictionary of Dictionaries
        ' containing Genius-related Properties
        ' for each Inventor Document in supplied
        ' Dictionary. Intended for invocation
        ' against a Dictionary of Inventor
        ' Documents generated by and/or within
        ' a separate function or procedure.
        '
        ' Initial creation intended to support
        ' companion function dcGnsPartsWithProps
        ' along with any others which might
        ' require it
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        Dim it As Inventor.PartDocument

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                it = aiDocPart(.Item(ky))
                If it Is Nothing Then
                    'Stop
                Else
                    rt.Add(ky, dcGnsPartProps(.Item(ky)))
                End If
            Next : End With

        Debug.Print("") 'Breakpoint Landing
        dcGnsPartsWithPropsFromDc = rt
    End Function

    Public Function dcOfDcAiPropVals(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        If dc Is Nothing Then
        Else
            With dc : For Each ky In .Keys
                    rt.Add(ky, dcAiPropValsFromDc(
                dcOb(.Item(ky))
            ))
                Next : End With
        End If

        Debug.Print("") 'Breakpoint Landing
        dcOfDcAiPropVals = rt
    End Function

    Public Function dcSansNoVals(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim ob As Object

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                it = .Item(ky) : With rt
                    If TypeOf it Is Object Then
                        ' ob = obOf(it)
                        If obOf(it) Is Nothing Then 'don't keep
                            Debug.Print("") 'Breakpoint Landing
                        Else
                            .Add(ky, it)
                            Debug.Print("") 'Breakpoint Landing
                        End If
                    ElseIf it Is Nothing Then 'don't keep
                        Debug.Print("") 'Breakpoint Landing
                    ElseIf it Is Nothing Then 'don't keep
                        Debug.Print("") 'Breakpoint Landing
                    Else
                        .Add(ky, it)
                        Debug.Print("") 'Breakpoint Landing
                    End If
                End With
            Next : End With
        dcSansNoVals = rt
    End Function

    Public Function dcOfOnlyNoVals(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim ob As Object

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                it = .Item(ky) : With rt
                    If TypeOf it Is Object Then
                        ' ob = obOf(it)
                        If obOf(it) Is Nothing Then 'don't keep
                            .Add(ky, it)
                            Debug.Print("") 'Breakpoint Landing
                        Else
                            Debug.Print("") 'Breakpoint Landing
                        End If
                    ElseIf it Is Nothing Then 'don't keep
                        .Add(ky, it)
                        Debug.Print("") 'Breakpoint Landing
                    ElseIf it Is Nothing Then 'don't keep
                        .Add(ky, it)
                        Debug.Print("") 'Breakpoint Landing
                    Else
                        Debug.Print("") 'Breakpoint Landing
                    End If
                End With
            Next : End With
        dcOfOnlyNoVals = rt
    End Function

    Public Function dc4noValStatus(it As Object,
    hasVal As Scripting.Dictionary,
    noVal As Scripting.Dictionary
) As Scripting.Dictionary
        If TypeOf it Is Object Then
            If obOf(it) Is Nothing Then
                dc4noValStatus = noVal
                Debug.Print("") 'Breakpoint Landing
            Else
                dc4noValStatus = hasVal
                Debug.Print("") 'Breakpoint Landing
            End If
        ElseIf it Is Nothing Then
            dc4noValStatus = noVal
            Debug.Print("") 'Breakpoint Landing
        ElseIf it Is Nothing Then
            dc4noValStatus = noVal
            Debug.Print("") 'Breakpoint Landing
        Else
            dc4noValStatus = hasVal
            Debug.Print("") 'Breakpoint Landing
        End If
    End Function

    Public Function dcOfNoValStatus(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim nv As Scripting.Dictionary
        Dim hv As Scripting.Dictionary
        Dim ky As Object
        Dim it As Object
        Dim ob As Object
        Dim ck As Long

        rt = New Scripting.Dictionary
        hv = New Scripting.Dictionary
        nv = New Scripting.Dictionary
        rt.Add("HASVAL", hv)
        rt.Add("NOVAL", nv)

        With dc : For Each ky In .Keys
                dc4noValStatus(
            .Item(ky), hv, nv
        ).Add(ky, .Item(ky))
            Next : End With
        dcOfNoValStatus = rt
    End Function

    Public Function dcOfDcNoValStatus(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcOfDcNoValStatus
        '
        ' Given a Dictionary of Dictionaries,
        ' return a Dictionary of "No Value Status"
        ' Dictionaries for each Item in the
        ' source Dictionary
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                rt.Add(ky, dcOfNoValStatus(
            dcOb(.Item(ky))
        ))
            Next : End With
        dcOfDcNoValStatus = rt
    End Function

    Public Function dcOfDcWithNoVals(
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' dcOfDcWithNoVals
        '
        ' Given a Dictionary of Dictionaries,
        ' return a sub Dictionary of those
        ' with at least one "No Value" Item
        '
        Dim rt As Scripting.Dictionary
        Dim wk As Scripting.Dictionary
        Dim ky As Object

        rt = New Scripting.Dictionary
        With dcOfDcNoValStatus(dc) : For Each ky In .Keys
                wk = dcOb(dcOb(.Item(ky)).Item("NOVAL"))
                If wk.Count > 0 Then rt.Add(ky, wk)
            Next : End With
        dcOfDcWithNoVals = rt
        'Debug.Print(txDumpLs(dcOfDcWithNoVals(dcGnsPartsWithProps(ThisApplication.ActiveDocument)).Keys)
    End Function

    Public Function mGr1g0f1(
    ob As Inventor.PartDocument,
    dcIfIs As Scripting.Dictionary,
    dcIfNot As Scripting.Dictionary
) As Scripting.Dictionary 'Object
        '
        '
        '
        If ob Is Nothing Then
            Stop
        Else
            If ob.ComponentDefinition.IsContentMember Then
                mGr1g0f1 = dcIfIs
            Else
                mGr1g0f1 = dcIfNot
            End If
        End If
    End Function

    Public Function mGr1g0f2(
    dc As Scripting.Dictionary
    ) As Scripting.Dictionary
        '
        '
        '
        Dim rt As Scripting.Dictionary
        Dim ky As Object
        Dim InvProperty As Scripting.Dictionary

        rt = New Scripting.Dictionary
        With dc : For Each ky In .Keys
                InvProperty = dcOb(.Item(ky))
                If InvProperty Is Nothing Then
                    Stop
                Else
                    rt.Add(ky, DumpLsKeyVal(
                        mGr1g0f1(InvProperty, New Scripting.Dictionary, New Scripting.Dictionary)
                    ))
                End If
            Next : End With
        mGr1g0f2 = rt
    End Function

    Public Function dcGeniusPropsPartPre20180530(
    oDoc As Inventor.PartDocument,
    Optional dc As Scripting.Dictionary = Nothing
    ) As Scripting.Dictionary
        '
        ' REV[2022.08.26.1204]
        ' moved to module modGPUpdateATrev
        ' from modGPUpdateAT to get it out
        ' of the way, while keeping it on
        ' hand for reference, just in case.
        '
        Dim rt As Scripting.Dictionary
        ''
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        ''
        Dim prFamily As Inventor.Property
        Dim prRawMatl As Inventor.Property 'pnRawMaterial
        Dim prRmUnit As Inventor.Property 'pnRmUnit
        Dim prRmQty As Inventor.Property 'pnRmQty
        ''
        Dim Family As String
        Dim pnStock As String
        Dim qtUnit As String
        Dim bomStruct As Inventor.BOMStructureEnum
        Dim ck As VbMsgBoxResult

        If dc Is Nothing Then
            dcGeniusPropsPartPre20180530 =
                dcGeniusPropsPartPre20180530(
                    oDoc, New Scripting.Dictionary)
        Else
            rt = dc

            With oDoc
                ' Get the custom property .
                aiPropsUser = .Propertys.Item(gnCustom)
                aiPropsDesign = .Propertys.Item(gnDesign)
                prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                ''[2018-03-13:Add 1 to create RM property if not found]
                ''[2018-05-15:Add following to get props for RM Unit & Qty]
                prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                ''
                prFamily = aiGetProp(aiPropsDesign, pnFamily)

                ' We should check HERE for possibly misidentified purchased parts
                ' UPDATE[2018.02.06]: Using new UserForm, see below
                With .ComponentDefinition
                    ck = vbNo
                    If InStr(1, oDoc.FullFileName, "\Doyle_Vault\Designs\purchased\") > 0 Then
                        ' UPDATE[2018.02.06]: Using new UserForm
                        '     to show image and details
                        '     of part to be verified.
                        ck = newFmTest2().AskAbout(oDoc, , "Is this a Purchased Part?")
                    ElseIf InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" & prFamily.Text & "|") > 0 Then
                        ' This ElseIf condition should be combinable
                        ' with the initial If condition above
                        ' to simplify this check process.
                        ' All/most text in this clause should be
                        ' redundant and removable, once updated
                        ' check process has been validated.
                        ' UPDATE[2018.02.06]: Using same
                        '     new UserForm as noted above.
                        ck = newFmTest2().AskAbout(oDoc, , "Is this a Purchased Part?")
                    End If

                    ' Check process below replaces duplicate check/responses above.
                    ' Should be able to merge back into main branch
                    ' once code above is validated and refactored.
                    If ck = vbYes Then
                        If .BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                            .BOMStructure = BOMStructureEnum.kPurchasedBOMStructure
                        End If
                    End If

                    bomStruct = .BOMStructure
                End With

                'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
                With .ComponentDefinition.MassProperties
                    rt = dcWithProp(aiPropsUser, pnMass, System.Math.Round(.Mass * cvMassKg2LbM, 4), rt)
                End With

                '----------------------------------------------------'
                If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                    '----------------------------------------------------'
                    'Request #4: Change Cost Center iProperty.
                    'If BOMStructure = Normal, then Family = D-MTO,
                    'else if BOMStructure = Purchased then Family = D-PTS.
                    'With .ComponentDefinition

                    If bomStruct = BOMStructureEnum.kNormalBOMStructure Then '.BOMStructure
                        'If prRawMatl.Text = "" Or cnGnsDoyle().Execute("select I.Family from vgMfiItems As I where I.Item='" & prRawMatl.Text & "';").Fields("Family").Text = "DSHEET" Then
                        'Request #3:
                        '   Get sheet metal extent area
                        '   and add to custom property "RMQTY"
                        rt = dcFlatPatProps(.ComponentDefinition, rt)

                        'Moved to start of block to check for NON sheet metal

                        'NOTE: THIS call might best be combined somehow
                        '   with the flat pattern prop pickup above.
                        '   Note especially that if dcFlatPatProps
                        '   FINDS NO .FlatPattern, then there should
                        '   BE NO sheet metal part number!
                        If prRawMatl Is Nothing Then
                            If rt.Exists("~OFFTHK") Then
                                pnStock = ""
                            Else
                                Stop 'because we don't know IF this is sheet metal yet
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If
                        Else
                            If prRawMatl.Text = "" Then
                                'Stop 'because we're not sure what we have.
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            Else
                                Stop
                                With CnGnsDoyle().Execute(
                                    "select i.Family from vgMfiItems i " &
                                    "where i.Item='" & prRawMatl.Text & "';"
                                )
                                    With .Fields("Family")
                                        If .Value = "DSHEET" Then
                                            'We should be okay. This is sheet metal stock
                                        ElseIf .Value = "D-BAR" Then
                                            Stop 'because we might want a D-BAR handler
                                        Else
                                            Stop 'because we don't know WHAT do with it
                                        End If
                                    End With
                                End With
                                pnStock = prRawMatl.Text
                            End If
                        End If

                        If Len(pnStock) > 0 Then
                            'Stop
                            Family = "D-MTO"
                            ' This needs to change.
                            ' Got code above that might better determind
                            ' whether this should be D-MTO or R-MTO
                            ' and don't forget about D/R-PTS/O options
                            '
                            ' NEW IDEA[2018-05-04]
                            ' We SHOULD be able to guess between D-MTO and R-MTO
                            ' based on the family of the raw material,
                            ' as determined in the prior section.
                            ' Plan to adjust the code here and above
                            ' to allow for that opportunity.
                            '
                        Else
                            '
                            ' We MIGHT have an incorrectly marked PURCHASED part
                            'Stop
                            ' We'll want to see about fixing that here, maybe?
                            '

                            With newFmTest1()
                                'aiSMdef.Document
                                If Not oDoc.ComponentDefinition.Document Is oDoc Then Stop
                                ' Define local variables for thickness and material
                                Dim invSheetMetalThickness As Double
                                Dim invSheetMetalMaterial As String
                                Dim docName As String

                                ' Attempt to get thickness and material from the ComponentDefinition
                                On Error Resume Next
                                invSheetMetalThickness = 0#
                                invSheetMetalMaterial = ""
                                docName = oDoc.DisplayName
                                ' Try to get thickness property (if it exists)
                                If oDoc.ComponentDefinition.Type = ObjectTypeEnum.kSheetMetalComponentDefinitionObject Then
                                    invSheetMetalThickness = oDoc.ComponentDefinition.Thickness
                                    invSheetMetalMaterial = oDoc.ComponentDefinition.Material.Name
                                End If
                                On Error GoTo 0

                                If .AskAbout(oDoc, "No Stock Found! Please Review") = vbYes Then
                                    Join({
                                        "NO STOCK# for",
                                        Format(invSheetMetalThickness, "0.000") & "in",
                                        invSheetMetalMaterial,
                                        "in " & docName,
                                        "Stop/pause here?"
                                    }, vbCrLf)
                                End If
                            End With
                            With .ItemData
                                If .Exists(pnFamily) Then
                                    Family = .Item(pnFamily)
                                    Debug.Print(pnFamily & "=" & Family)
                                End If

                                If .Exists(pnRawMaterial) Then
                                    pnStock = .Item(pnRawMaterial)
                                    Debug.Print(pnRawMaterial & "=" & pnStock)
                                End If
                            End With
                        End If

                        ' prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                        ' rt = dcAddProp(prRmUnit, rt)
                        qtUnit = "FT2"
                        prRawMatl.Text = pnStock
                        rt = dcAddProp(prRawMatl, rt)
                        ' rt = dcWithProp(aiPropsUser, pnRawMaterial, pnStock, rt)
                        '
                        If aiGetProp(aiPropsUser, pnRmUnit) Is Nothing Then
                            'Stop
                            'Else
                            'If aiGetProp(aiPropsUser, pnRmUnit).Text <> "FT2" Then
                            '         ' prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                            If Len(prRmUnit.Text) > 0 Then
                                If prRmUnit.Text <> qtUnit Then
                                    Stop 'so we DON'T automatically "fix" the RMUNIT value
                                    '"FT2"
                                End If
                            Else
                                'we're ting a new quantity unit
                            End If
                            'End If
                            ' When this Stop activates, skip the next line
                            rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                            ' Want to change this part to allow for alternate RMUNIT values
                            ' When prior Stop is activated, use Ctrl-F9
                            ' to continue at the Stop line below.
                            If 0 Then Stop 'to give us a skipover point

                            'Moved Flat Pattern data collection to beginning of block
                            'to facilitate detection of possible NON sheet metal part
                            'Else
                            'Debug.Print(prRawMatl.Text
                            'Stop
                            'End If
                        End If
                    ElseIf bomStruct = BOMStructureEnum.kPurchasedBOMStructure Then '.BOMStructure
                        Family = "D-PTS"
                    Else
                        Stop 'because we might need to do something else
                    End If
                Else
                    '--------------------------'
                    '  Else 'for standard Part ---'
                    '--------------------------'
                    'Request #2: Change Cost Center iProperty.
                    'If BOMStructure = Purchased and not content center,
                    'then Family = D-PTS, else Family = D-HDWR.
                    With .ComponentDefinition
                        If bomStruct = BOMStructureEnum.kPurchasedBOMStructure _
                            And .IsContentMember = False Then
                            Family = "D-PTS"
                        Else
                            'Family = "D-HDWR"
                            Family = ""
                        End If
                    End With
                End If

                ' Get the design tracking property ,
                ' and update the Cost Center Property
                If Len(Family) > 0 Then
                    rt = dcWithProp(aiPropsDesign, pnFamily, Family, rt)
                End If

                Call iSyncPartFactory(oDoc) 'Backport Properties to iPart Factory
                dcGeniusPropsPartPre20180530 = rt
            End With
        End If
    End Function

End Module