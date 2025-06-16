Module src0
    Dim inventorApp As Inventor.Application
    Public Sub GeniusPropertiesUpdater()

        Dim answer As Integer

        answer = MsgBox("Are you sure you want to process this document? The process may require a few minutes depending on assembly size. Suppressed and excluded parts will not be processed", vbYesNo + vbQuestion, "Process Document Custom iProperties")

        If answer = vbYes Then


            Dim ActiveDoc As Document
            ActiveDoc = InventorApp.ActiveDocument

            If ActiveDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

                ' Get the active assembly.
                Dim invAsmDoc As AssemblyDocument
                invAsmDoc = InventorApp.ActiveDocument
                'MsgBox(("Assembly Name: " & invAsmDoc.DisplayName)

                'Call IterateAssy(invAsmDoc.ComponentDefinition.Occurrences, 1)

                MsgBox("Process completed")

            Else

                Dim invPartDoc As PartDocument
                invPartDoc = InventorApp.ActiveDocument

                'Call IteratePart(invPartDoc)

                MsgBox("Process completed: Part " & invPartDoc.DisplayName & " processed")


            End If

        Else

            'do nothing

        End If

    End Sub
    '
    '

    Public Function IterateAssyRevA0(Occurences As ComponentOccurrences, Level As Integer) As Long
        'Iterate through the assembly
        Dim invOcc As ComponentOccurrence
        Dim TotalStepCount As Long
        Dim CurrentStepCount As Long
        Dim invProgressBar As ProgressBar
        Dim invDoc As PartDocument
        Dim invSheetMetalComp As SheetMetalComponentDefinition
        Dim invCustomProperty As Inventor.Property
        Dim invSheetMetalMass As Double
        Dim invGeniusMassProperty As Inventor.Property
        Dim invGeniusMaterial As String
        Dim invSheetMetalName As String
        Dim invSheetMetalMaterial As String
        Dim invRMProperty As Inventor.Property
        Dim invRMUOMProperty As Inventor.Property
        Dim invFlatPattern As FlatPattern
        Dim oExtent As Box
        Dim dLength As Double
        Dim dWidth As Double
        Dim dArea As Double
        Dim oUOM As UnitsOfMeasure
        Dim strWidth As String
        Dim strLength As String
        Dim strArea As String
        Dim invRMQTYProperty As Inventor.Property
        Dim invWidthProperty As Inventor.Property
        Dim invLengthProperty As Inventor.Property
        Dim invAreaProperty As Inventor.Property
        Dim Family As String
        Dim invDesignInfo As Inventor.Property
        Dim invCostCenterProperty As Inventor.Property
        Dim invPartDocComp As PartComponentDefinition
        Dim invPartMass As Double
        Dim invCustomPartProperty As Inventor.Property
        Dim invGeniusPartMassProperty As Inventor.Property
        'Dim Family as String
        'Dim invDesignInfo As Property
        'Dim invCostCenterProperty As Property

        For Each invOcc In Occurences
            'MsgBox(("TYPE: " & invOcc.Definition.Type & vbCrLf & "VISIBLE: " & invOcc.Visible & vbCrLf & "NAME: " & invOcc.Name & vbCrLf & "Suboccurence: " & invOcc.SubOccurrences.Count & vbCrLf & "Occurence Type: " & invOcc.Definition.Occurrences.Type & vbCrLf & "BOMStructure: " & invOcc.BOMStructure)
            'Remove suppressed and excluded parts from the process
            If invOcc.Definition.Type <> ObjectTypeEnum.kAssemblyComponentDefinitionObject And invOcc.Definition.Type <> ObjectTypeEnum.kWeldmentComponentDefinitionObject Then
                If False And invOcc.Visible And Not invOcc.Suppressed And Not invOcc.Excluded And invOcc.Definition.Type <> ObjectTypeEnum.kWeldsComponentDefinitionObject Then
                    '-------------------------------'
                    'Create Progress Bar Information
                    '-------------------------------'

                    'Define Total Steps
                    TotalStepCount = Occurences.Count

                    'Define Current Step
                    CurrentStepCount = CurrentStepCount + 1

                    'Create a new ProgressBar object.
                    invProgressBar = InventorApp.CreateProgressBar(True, TotalStepCount, "Progressing: ")

                    '  the message for the progress bar
                    invProgressBar.Message = "Processing - " & invOcc.Name & " - " & CurrentStepCount & "/" & TotalStepCount
                    invProgressBar.UpdateProgress()

                    ' Get the active part document.
                    invDoc = invOcc.Definition.Document

                    '-------------------'
                    'Check if SheetMetal'
                    '-------------------'
                    If False And invDoc.SubType = guidSheetMetal Then
                        invSheetMetalComp = invDoc.ComponentDefinition

                        ' Get the custom property .
                        invCustomProperty =
                        invDoc.Propertys.Item("Inventor User Defined Properties")

                        'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
                        invSheetMetalMass = System.Math.Round(invSheetMetalComp.MassProperties.Mass * cvMassKg2LbM, 4)

                        ' Attempt to get an existing custom property named "GeniusMass".
                        On Error Resume Next
                        invGeniusMassProperty = invCustomProperty.Item(pnMass)
                        If Err.Number <> 0 Then
                            ' Failed to get the property, which means it doesn't exist so we'll create it.
                            Call invCustomProperty.Add(invSheetMetalMass, pnMass)
                        Else
                            ' Got the property so update the value.
                            invGeniusMassProperty.Text = invSheetMetalMass
                        End If
                        'Request #2: Get Genius SheetMetal by matching Style Name and Material. Add to Custom Property RM


                        invSheetMetalName = invSheetMetalComp.ActiveSheetMetalStyle.Name

                        invSheetMetalMaterial = invSheetMetalComp.ActiveSheetMetalStyle.Material.Name

                        'Map combination to corresponding Genius Part Number
                        If invSheetMetalMaterial = "Stainless Steel" Then
                            If invSheetMetalName = "18 GA" Then
                                invGeniusMaterial = "FS-48x96x0.048"
                            ElseIf invSheetMetalName = "14 GA" Then
                                invGeniusMaterial = "FS-60x120x0.075"
                            ElseIf invSheetMetalName = "13 GA" Then
                                invGeniusMaterial = "FS-60x97x0.09"
                            ElseIf invSheetMetalName = "12 GA" Then
                                invGeniusMaterial = "FS-60x120x0.105"
                            ElseIf invSheetMetalName = "10 GA" Then
                                invGeniusMaterial = "FS-60x144x0.135"
                            ElseIf invSheetMetalName = "3/16""" Then
                                invGeniusMaterial = "FS-60x144x0.188"
                            ElseIf invSheetMetalName = "1/4""" Then
                                invGeniusMaterial = "FS-60x144x0.25"
                            ElseIf invSheetMetalName = "5/16""" Then
                                invGeniusMaterial = "FS-60x144x0.313"
                            ElseIf invSheetMetalName = "3/8""" Then
                                invGeniusMaterial = "FS-60x144x0.375"
                            ElseIf invSheetMetalName = "1/2""" Then
                                invGeniusMaterial = "FS-60x144x0.5"
                            Else
                                invGeniusMaterial = ""
                            End If
                        ElseIf invSheetMetalMaterial = "Steel, Mild" Then
                            If invSheetMetalName = "14 GA" Then
                                invGeniusMaterial = "FM-60x144x0.075"
                            ElseIf invSheetMetalName = "12 GA" Then
                                invGeniusMaterial = "FM-60x144x0.105"
                            ElseIf invSheetMetalName = "10 GA" Then
                                invGeniusMaterial = "FM-60x144x0.135"
                            ElseIf invSheetMetalName = "3/16""" Then
                                invGeniusMaterial = "FM-60x144x0.188"
                            ElseIf invSheetMetalName = "1/4""" Then
                                invGeniusMaterial = "FM-60x144x0.25"
                            ElseIf invSheetMetalName = "5/16""" Then
                                invGeniusMaterial = "FM-60x144x0.313"
                            ElseIf invSheetMetalName = "3/8""" Then
                                invGeniusMaterial = "FM-60x144x0.375"
                            ElseIf invSheetMetalName = "1/2""" Then
                                invGeniusMaterial = "FM-60x144x0.5"
                            ElseIf invSheetMetalName = "5/8""" Then
                                invGeniusMaterial = "FM-60x144x0.625"
                            ElseIf invSheetMetalName = "3/4""" Then
                                invGeniusMaterial = "FM-60x120x0.75"
                            ElseIf invSheetMetalName = "1""" Then
                                invGeniusMaterial = "FM-48x120x1"
                            Else
                                invGeniusMaterial = ""
                            End If
                        Else
                            invGeniusMaterial = ""
                        End If 'Mapping of material

                        ' Attempt to get an existing custom property named "RM".
                        On Error Resume Next
                        invRMProperty = invCustomProperty.Item(pnRawMaterial)
                        If Err.Number <> 0 Then
                            ' Failed to get the property, which means it doesn't exist so we'll create it.
                            Call invCustomProperty.Add(invGeniusMaterial, pnRawMaterial)
                        Else
                            ' Got the property so update the value.
                            invRMProperty.Text = invGeniusMaterial
                        End If

                        ' Attempt to get an existing custom property named "RMUOM".
                        On Error Resume Next
                        invRMUOMProperty = invCustomProperty.Item(pnRmUnit)
                        If Err.Number <> 0 Then
                            ' Failed to get the property, which means it doesn't exist so we'll create it.
                            Call invCustomProperty.Add("FT2", pnRmUnit)
                        Else
                            ' Got the property so update the value.
                            invRMUOMProperty.Text = "FT2"
                        End If

                        'Request #3: Get sheet metal extent area and add to custom property "RMQTY"
                        invFlatPattern = invSheetMetalComp.FlatPattern

                        'Check to see if flat exists
                        If Not invFlatPattern Is Nothing Then

                            ' Get the extent of the face.
                            oExtent = invFlatPattern.Body.RangeBox

                            ' Extract the width, length and area from the range.
                            dLength = oExtent.MaxPoint.X - oExtent.MinPoint.X
                            dWidth = oExtent.MaxPoint.Y - oExtent.MinPoint.Y
                            dArea = dLength * dWidth

                            ' Convert these values into the document units.
                            ' This will result in strings that are identical
                            ' to the strings shown in the Extent dialog.
                            oUOM = invDoc.UnitsOfMeasure
                            strWidth = oUOM.GetStringFromValue(dWidth, oUOM.GetStringFromType(oUOM.LengthUnits))
                            strLength = oUOM.GetStringFromValue(dLength, oUOM.GetStringFromType(oUOM.LengthUnits))
                            strArea = oUOM.GetStringFromValue(dArea, oUOM.GetStringFromType(oUOM.LengthUnits) & "^2")

                            ' Add area to custom property 
                            ' Attempt to get an existing custom property named "RMQTY".
                            On Error Resume Next
                            invRMQTYProperty = invCustomProperty.Item(pnRmQty)
                            If Err.Number <> 0 Then
                                ' Failed to get the property, which means it doesn't exist so we'll create it.
                                Call invCustomProperty.Add(dArea * cvArSqCm2SqFt, pnRmQty)
                            Else
                                ' Got the property so update the value.
                                invRMQTYProperty.Text = dArea * cvArSqCm2SqFt
                            End If

                            ' Add Width to custom property 
                            ' Attempt to get an existing custom property named "Extent_Width".
                            On Error Resume Next
                            invWidthProperty = invCustomProperty.Item(pnWidth)
                            If Err.Number <> 0 Then
                                ' Failed to get the property, which means it doesn't exist so we'll create it.
                                Call invCustomProperty.Add(strWidth, pnWidth)
                            Else
                                ' Got the property so update the value.
                                invWidthProperty.Text = strWidth
                            End If

                            ' Add Length to custom property 
                            ' Attempt to get an existing custom property named "Extent_Length".
                            On Error Resume Next
                            invLengthProperty = invCustomProperty.Item(pnLength)
                            If Err.Number <> 0 Then
                                ' Failed to get the property, which means it doesn't exist so we'll create it.
                                Call invCustomProperty.Add(strLength, pnLength)
                            Else
                                ' Got the property so update the value.
                                invLengthProperty.Text = strLength
                            End If

                            ' Add AreaDescription to custom property 
                            ' Attempt to get an existing custom property named "Extent_Area".
                            On Error Resume Next
                            invAreaProperty = invCustomProperty.Item(pnArea)
                            If Err.Number <> 0 Then
                                ' Failed to get the property, which means it doesn't exist so we'll create it.
                                Call invCustomProperty.Add(strArea, pnArea)
                            Else
                                ' Got the property so update the value.
                                invAreaProperty.Text = strArea
                            End If
                        End If

                        'Request #4: Change Cost Center iProperty. If BOMStructure = Normal, then Family = D-MTO, else if BOMStructure = Purchased then Family = D-PTS.

                        If invSheetMetalComp.BOMStructure = BOMStructureEnum.kNormalBOMStructure Then
                            Family = "D-MTO"
                        ElseIf invSheetMetalComp.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                            Family = "D-PTS"
                        End If

                        ' Get the design tracking property .
                        invDesignInfo =
                    invDoc.Propertys.Item("Design Tracking Properties")

                        ' Update the Cost Center Property
                        invCostCenterProperty = invDesignInfo.Item(pnFamily)
                        invCostCenterProperty.Text = Family
                        Family = Family 'Just put this in for a next line to run to (Ctrl-F8).
                        'Otherwise, stepping in or through previous line would run to end with no break

                        '----------------------'
                        'Else, if standard Part'
                        '----------------------'

                    Else
                        'Get the Parts Component Definition
                        invPartDocComp = invDoc.ComponentDefinition

                        'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
                        invPartMass = System.Math.Round(invPartDocComp.MassProperties.Mass * cvMassKg2LbM, 4)

                        ' Get the custom property .
                        invCustomPartProperty =
                        invDoc.Propertys.Item("Inventor User Defined Properties")

                        ' Attempt to get an existing custom property named "GeniusMass".
                        On Error Resume Next
                        invGeniusPartMassProperty = invCustomPartProperty.Item(pnMass)
                        If Err.Number <> 0 Then
                            ' Failed to get the property, which means it doesn't exist so we'll create it.
                            Call invCustomPartProperty.Add(invPartMass, pnMass)
                        Else
                            ' Got the property so update the value.
                            invGeniusPartMassProperty.Text = invPartMass
                        End If

                        'Request #2: Change Cost Center iProperty. If BOMStructure = Purchased and not content center, then Family = D-PTS, else Family = D-HDWR.
                        'Dim Family as String

                        If invPartDocComp.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure And invPartDocComp.IsContentMember = False Then
                            Family = "D-PTS"
                        Else
                            Family = "D-HDWR"
                        End If

                        ' Get the design tracking property .
                        'Dim invDesignInfo As Property
                        invDesignInfo =
                        invDoc.Propertys.Item("Design Tracking Properties")

                        ' Update the Cost Center Property
                        'Dim invCostCenterProperty As Property
                        invCostCenterProperty = invDesignInfo.Item(pnFamily)
                        invCostCenterProperty.Text = Family
                    End If 'Sheetmetal vs Part

                    ' Terminate the progress bar.
                    invProgressBar.Close()

                Else
                End If 'Visible, suppressed, excluded or Welds
            Else 'assembly, iterate through next level
                Debug.Print(IterateAssyRevA0(invOcc.SubOccurrences, Level + 1))
            End If 'part or assembly
        Next


    End Function
    'Debug.Print(IterateAssyRevA0(aiDocAssy(InventorApp.ActiveDocument).ComponentDefinition.Occurrences, 1)

End Module