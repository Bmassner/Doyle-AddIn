Module GeniusIntegration

    Public Sub Update_Genius_Properties(ThisApplication As Inventor.Application)
        Dim fc As GnsIfcAiDoc = New GnsIfcAiDoc()
        Dim dc As Scripting.Dictionary
        Dim rt As Scripting.Dictionary = New Scripting.Dictionary
        Dim ActiveDoc As Inventor.Documents
        Dim txOut As String

        ' Confirm user wants to process the active document
        If MsgBox("Process this document? This may take a few minutes.", vbYesNo + vbQuestion, "Process Document") <> vbYes Then Exit Sub


        If ActiveDoc Is Nothing Then
            MsgBox("No active document found.", vbExclamation)
            Exit Sub
        End If

        ' Collect components (parts or assemblies)
        Dim ct As Long = 1
        If ActiveDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            If MsgBox("Include main assembly in processing?", vbYesNo + vbQuestion, "Include Main Assembly?") <> vbYes Then ct = 0
        End If

        ' Get component dictionary by part number
        dc = dcOb(dcAiDocCompSetsByPtNum(ActiveDoc, ct).Item(1))

        ' Group components by form
        Dim groups As Scripting.Dictionary = dcAiDocGrpsByForm(dc)

        ' Process each group
        For Each groupKey In {"ASSY", "SHTM", "MAYB", "DBAR", "PRCH"}
            If groups.Exists(groupKey) Then
                Dim groupDict As Scripting.Dictionary = dcOb(groups.Item(groupKey))
                For Each partKey In groupDict.Keys
                    On Error GoTo HandleError
                    rt.Add(partKey, fc.Props(aiDocument(groupDict.Item(partKey))))
                    dcOb(rt.Item(partKey)).Add("FORM", groupKey)
HandleError:
                    If Err.Number <> 0 Then
                        Debug.Print("Error processing " & partKey & ": " & Err.Description)
                        Err.Clear()
                    End If
                Next
            End If
        Next

        ' Convert results to property values
        For Each key In rt.Keys
            rt.Item(key) = dcPropVals(dcOb(rt.Item(key)))
        Next

        ' Find any items not processed
        dc = DcKeysMissing(dc, rt)
        rt = dcRecSetDcDx4json(dcDxFromRecSetDc(rt))

        If dc.Count > 0 Then
            rt.Add("NOTPROCESSED", dc.Keys)
        End If

    End Sub

    Public Sub Update_iPtAssy_Genius_Props()
        Dim md As Inventor.Document = aiDocActive()
        Dim dc As Scripting.Dictionary = gnsUpdtAll_iFact(compDefOf(md))
        If dc.Count = 0 Then
            MsgBox("No iPart/iAssembly properties updated.", vbOKOnly)
        End If
    End Sub
End Module