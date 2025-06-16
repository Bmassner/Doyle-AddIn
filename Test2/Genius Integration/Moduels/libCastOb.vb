Module libCastOb
    Dim inventorApp As Inventor.Application
    Public Function obOf(vr As Object) As Object
        If TypeOf vr Is Object Then
            obOf = vr
        Else
            obOf = Nothing
        End If
    End Function

    Public Function dcOb(vr As Object) As Scripting.Dictionary
        If TypeOf vr Is Object Then
            If vr Is Nothing Then
                dcOb = Nothing
            ElseIf TypeOf vr Is Scripting.Dictionary Then
                dcOb = vr
            Else
                dcOb = Nothing
            End If
        Else
            dcOb = Nothing
        End If
    End Function

    Public Function fdOb(vr As Object) As ADODB.Field
        If TypeOf vr Is Object Then
            If TypeOf vr Is ADODB.Field Then
                fdOb = vr
            Else
                fdOb = Nothing
            End If
        Else
            fdOb = Nothing
        End If
    End Function

    Public Function aiDocument(doc As Object) As Inventor.Document
        If doc Is Nothing Then
            aiDocument = doc
        ElseIf TypeOf doc Is Inventor.Document Then
            aiDocument = doc
        Else
            aiDocument = Nothing
        End If
    End Function
    'For Each itm In ActiveDocsComponents(InventorApp): Debug.Print(aiDocument(obOf(itm)).FullFileName: Next

    Public Function aiDocActive() As Inventor.Document
        aiDocActive = InventorApp.ActiveDocument
    End Function

    Public Function aiDocPart(
    doc As Inventor.Document
) As Inventor.PartDocument
        If doc Is Nothing Then
            aiDocPart = doc
        ElseIf TypeOf doc Is Inventor.PartDocument Then
            aiDocPart = doc
        Else
            aiDocPart = Nothing
        End If
    End Function

    Public Function aiDocPartFromCCtr(
    doc As Inventor.Document
) As Inventor.PartDocument
        Dim rt As Inventor.PartDocument

        If doc Is Nothing Then
            aiDocPartFromCCtr = doc
        Else
            rt = aiDocPart(doc)
            If rt Is Nothing Then
                aiDocPartFromCCtr = rt
            ElseIf rt.ComponentDefinition.IsContentMember Then
                aiDocPartFromCCtr = rt
            Else
                aiDocPartFromCCtr = Nothing
            End If
        End If
    End Function

    Public Function aiDocAssy(
    doc As Inventor.Document
) As Inventor.AssemblyDocument
        If doc Is Nothing Then
            aiDocAssy = doc
        ElseIf TypeOf doc Is Inventor.AssemblyDocument Then
            aiDocAssy = doc
        Else
            aiDocAssy = Nothing
        End If
    End Function

    Public Function aiDocDwg(
    doc As Inventor.Document
) As Inventor.DrawingDocument
        If doc Is Nothing Then
            aiDocDwg = doc
        ElseIf TypeOf doc Is Inventor.DrawingDocument Then
            aiDocDwg = doc
        Else
            aiDocDwg = Nothing
        End If
    End Function

    Private Function aiCompDefinition(
    doc As Object
) As Inventor.ComponentDefinition
        '
        ' REV[2022.08.31.1313] OBSOLETED
        ' -   no calls found to this function
        ' -   aiCompDefOf serves same purpose
        '     in (slightly?) more robust manner
        ' -   changed scope to Private
        '     to prevent future usage
        '     outside local scope
        '
        If TypeOf doc Is Inventor.ComponentDefinition Then
            aiCompDefinition = doc
        Else
            aiCompDefinition = Nothing
        End If
    End Function

    Public Function aiCompDefOf(doc As Object
) As Inventor.ComponentDefinition 'Inventor.Document
        '
        ' aiCompDefOf -- Return the ComponentDefinition
        '     of ANY Inventor Document which has one.
        ' NOTE: currently returns ComponentDefinition objects
        '     only from Part and Assembly Documents.
        ' NOTE[2022.08.31.1202]: copied comments from redundant
        '     function obAiCompDefAny prior to its deprecation
        '
        If doc Is Nothing Then
            aiCompDefOf = Nothing
        ElseIf TypeOf doc Is Inventor.ComponentDefinition Then
            aiCompDefOf = doc
        ElseIf TypeOf doc Is Inventor.Document Then
            With aiDocument(doc)
                If .DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    aiCompDefOf = aiDocAssy(doc).ComponentDefinition
                ElseIf .DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    aiCompDefOf = aiDocPart(doc).ComponentDefinition
                Else
                    aiCompDefOf = Nothing
                End If
            End With
        Else
            aiCompDefOf = Nothing
        End If
    End Function

    Public Function obAiCompDefAny(
    AiDoc As Inventor.Document
) As Inventor.ComponentDefinition
        '
        ' obAiCompDefAny -- Return the ComponentDefinition
        '     of ANY Inventor Document which has one.
        ' NOTE: currently returns ComponentDefinition objects
        '     only from Part and Assembly Documents.
        ' NOTE[2022.08.31.1203]: rediscovered original
        '     implementation aiCompDefOf, copied comments
        '     there prior to deprecation of this implementation
        '
        If AiDoc Is Nothing Then
            obAiCompDefAny = Nothing
        ElseIf AiDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            obAiCompDefAny = aiDocAssy(AiDoc).ComponentDefinition
        ElseIf AiDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            obAiCompDefAny = aiDocPart(AiDoc).ComponentDefinition
        Else
            obAiCompDefAny = Nothing
        End If
        'Debug.Print(TypeName(obAiCompDefAny(aiDocument(userChoiceFromDc())))
    End Function

    Public Function aiCompDefPart(doc As Object
) As Inventor.PartComponentDefinition
        '
        ' REV[2022.08.31.1247]
        ' added ElseIf check for PartDocument
        ' to accept Inventor Document as well
        ' as ComponentDefinition
        ' applied same to functions {
        '     aiCompDefPart
        ' }
        '
        '
        If doc Is Nothing Then
            aiCompDefPart = Nothing
        ElseIf TypeOf doc Is Inventor.PartComponentDefinition Then
            aiCompDefPart = doc
        ElseIf TypeOf doc Is Inventor.PartDocument Then
            aiCompDefPart = aiDocPart(doc).ComponentDefinition
        Else
            aiCompDefPart = Nothing
        End If
    End Function

    Public Function aiCompDefShtMetal(ob As Object
) As Inventor.SheetMetalComponentDefinition
        If ob Is Nothing Then
            aiCompDefShtMetal = Nothing
        ElseIf TypeOf ob Is Inventor.SheetMetalComponentDefinition Then
            aiCompDefShtMetal = ob
        ElseIf TypeOf ob Is Inventor.PartDocument Then
            aiCompDefShtMetal = aiCompDefShtMetal(
            aiDocPart(ob).ComponentDefinition
        )
        Else
            aiCompDefShtMetal = Nothing
        End If
    End Function

    Public Function aiCompDefAssy(ob As Object
) As Inventor.AssemblyComponentDefinition
        If ob Is Nothing Then
            aiCompDefAssy = Nothing
        ElseIf TypeOf ob Is Inventor.AssemblyComponentDefinition Then
            aiCompDefAssy = ob
        ElseIf TypeOf ob Is Inventor.AssemblyDocument Then
            aiCompDefAssy = aiDocAssy(ob).ComponentDefinition
        Else
            aiCompDefAssy = Nothing
        End If
    End Function

    Public Function aiProperty(ob As Object) As Inventor.Property
        If ob Is Nothing Then
            'Stop
            aiProperty = Nothing
        ElseIf TypeOf ob Is Inventor.Property Then
            aiProperty = ob
        Else
            'Stop 'because this is NOT a Property!
            aiProperty = Nothing
        End If
    End Function

    Public Function aiPlane(ob As Object) As Inventor.Plane
        If TypeOf ob Is Inventor.Plane Then
            aiPlane = ob
        Else
            aiPlane = Nothing
        End If
    End Function

    Public Function aiCompOcc(
    ob As Object
) As Inventor.ComponentOccurrence
        If TypeOf ob Is Inventor.ComponentOccurrence Then
            aiCompOcc = ob
        Else
            aiCompOcc = Nothing
        End If
    End Function

    Public Function obAiProp(ob As Object) As Inventor.Property
        If TypeOf ob Is Inventor.Property Then
            obAiProp = ob
        Else
            obAiProp = Nothing
        End If
    End Function

    Public Function obAiParam(ob As Object) As Inventor.Parameter
        If TypeOf ob Is Inventor.Parameter Then
            obAiParam = ob
        Else
            obAiParam = Nothing
        End If
    End Function

    Public Function obVbProject(ob As Object) As VBIDE.VBProject
        If TypeOf ob Is VBIDE.VBProject Then
            obVbProject = ob
        Else
            obVbProject = Nothing
        End If
    End Function

    Public Function obVbCodeMod(ob As Object) As VBIDE.CodeModule
        If TypeOf ob Is VBIDE.CodeModule Then
            obVbCodeMod = ob
        Else
            obVbCodeMod = Nothing
        End If
    End Function

End Module