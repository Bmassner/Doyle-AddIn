Class GnsIfcAiDoc
    Private cn As ADODB.Connection
    Private rs As ADODB.Record
    Private rsFamBuy As ADODB.Record
    Private fdItem As ADODB.Field
    Private fdFamily As ADODB.Field
    Private fd As ADODB.Field

    '
    ' NOTE: the following SQL text constants
    ' are left over from the initial design
    ' of this class (under a different name).
    '
    ' Their code will remain in place until
    ' such time as their value may be better
    ' ascertained. Assuming they remain useful,
    ' they should be exported to a separate
    ' library module for storage of SQL source.
    '
    Private Const sql01 As String = "" _
    & "select Item, Family " _
    & "from vgMfiItems " _
    & "" _
    & ""

    Private Const sql02 As String = "" _
    & "Select F.Family, F.Description1, " _
    & "F.DefaultPlanningId As pln, " _
    & "F.ProductCategory As cat " _
    & "From vgMfiFamilies F " _
    & "Where F.Type = 'R' " _
    & "And F.FamilyGroup = 'PARTS' " _
    & ""

    Public Function Props(AiDoc As Inventor.Document,
    Optional dc As Scripting.Dictionary = Nothing
) As Scripting.Dictionary
        If dc Is Nothing Then
            Props = Props(AiDoc,
        New Scripting.Dictionary)
            'ElseIf AiDoc Is Nothing Then
            ' REV[2022.03.18.1111] <-(seriously!)
            '     disable branch for void AiDoc
            '     should not be necessary since
            '     fall-through branch should
            '     return dc regardless
            '     '
            '     Props = dc
        Else
            Props =
        assyProps(aiDocAssy(AiDoc),
        partProps(aiDocPart(AiDoc),
        gnrlProps(AiDoc, dc)))
        End If
    End Function

    Private Function gnrlProps(
    AiDoc As Inventor.Document,
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        '
        ' gnrlProps -- derived 2023.01.13 from partProps
        '     to collect (and ) Properties applicable
        '     to both Parts and Assemblies
        '     '
        '     original intent to remove Family assignment
        '     from Part Property section to a more general
        '     applicable to both Parts and Assemblies,
        '     however, that's proving a more challenging
        '     task than anticipated
        '     '
        '     presently just a stub, pending further review
        '
        If AiDoc Is Nothing Then
            gnrlProps = dc
        Else
            'With AiDoc
            '    With .Propertys
            '        With .Item(gnDesign)
            '        End With
            '    End With
            'End With
            gnrlProps = dc
        End If
    End Function

    Private Function partProps(
    AiDoc As Inventor.PartDocument,
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        If AiDoc Is Nothing Then
            partProps = dc
        Else
            Debug.Print("") 'Breakpoint Landing
            partProps =
        dcGeniusPropsPart(
        AiDoc, dc)
            Debug.Print("") 'Breakpoint Landing
        End If
    End Function

    Private Function assyProps(
    AiDoc As Inventor.AssemblyDocument,
    dc As Scripting.Dictionary
) As Scripting.Dictionary
        If AiDoc Is Nothing Then
            assyProps = dc
        Else
            Debug.Print("") 'Breakpoint Landing
            assyProps =
        dcGeniusPropsAssy(
        AiDoc, dc)
            Debug.Print("") 'Breakpoint Landing
        End If
    End Function

    Private Sub Class_Initialize()
        ' cn = cnGnsDoyle()
        ' rs = cn.Execute(sql01)
        'With rs
        '    With .Fields
        '         fdItem = .Item("Item")
        '         fdFamily = .Item("Family")
        '    End With
        'End With

        ' rsFamBuy = cn.Execute(sql02)
        'With rsFamBuy
        '    With .Fields
        '    End With
        'End With
    End Sub

    '
    ' OBSOLETE SECTION
    '
    ' All functions below this comment
    ' are left over from the initial
    ' effort to create some form of
    ' Genius-oriented interface to
    ' Autodesk Inventor Documents,
    ' and particularly Part Documents.
    '
    ' The TestXX functions, originally Public,
    ' have been rendered Private to hide them
    ' from any client processes. Their code
    ' remains in place pending possible use
    ' in some future process(es).
    '
    ' They should at some point be removed,
    ' once their value is better established,
    ' and any useful portions incorporated
    ' into appropriate procedures.
    '

    Private Function Test01(invDoc As Inventor.PartDocument) As Inventor.BOMStructureEnum
        ' Present Role: Categorize Part Document
        '
        '
        Dim nmFamily As String
        Dim bomStruct As Inventor.BOMStructureEnum
        '
        '

        With invDoc
            nmFamily = .Propertys(gnDesign).Item(pnFamily).Text

            If .ComponentDefinition.IsContentMember Then
                If .ComponentDefinition.BOMStructure _
            = BOMStructureEnum.kPurchasedBOMStructure Then
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                    nmFamily = "D-HDWR"
                Else
                    Stop
                End If
            Else 'ry to identify other purchased part
                If InStr(1, invDoc.FullFileName,
                "\Doyle_Vault\Designs\purchased\"
            ) > 0 Then 'this is PROBABLY a purchased part
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                ElseIf g0f0(nmFamily) = BOMStructureEnum.kPurchasedBOMStructure Then
                    'this is almost certainly a purchased part
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                Else 'we'll assume it's NOT purchased.
                    bomStruct = BOMStructureEnum.kDefaultBOMStructure
                    'Use Default to indicate
                    'NON purchased parts.
                    'We can determine ACTUAL
                    'BOM structure elsewhere.
                End If
            End If
            ''
            '''

        End With
        Test01 = bomStruct
        '
        '
        Debug.Print("") 'Landing Line for Debug use. Do not disable.
        '
        '
        '
    End Function

    Private Function Test02(invDoc As Inventor.PartDocument) As Inventor.BOMStructureEnum
        ' Present Role: Categorize Part Document
        '
        '
        Dim rt As Scripting.Dictionary
        '    ''
        Dim aiPropsUser As Inventor.Property
        Dim aiPropsDesign As Inventor.Property
        '    ''
        Dim prFamily As Inventor.Property
        'Dim prPartNum   As Inventor.Property
        '    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
        '    Dim prRmUnit    As Inventor.Property 'pnRmUnit
        '    Dim prRmQty     As Inventor.Property 'pnRmQty
        '    ''
        Dim nmFamily As String
        '    Dim mtFamily As String
        '    ' UPDATE[2018.05.30]:
        '    '     Rename variable Family to nmFamily
        '    '     to minimize confusion between code
        '    '     and comment text in searches.
        '    '     Also add variable mtFamily
        '    '     for raw material Family name
        '    Dim pnStock As String
        '    Dim qtUnit As String
        Dim bomStruct As Inventor.BOMStructureEnum
        '    Dim ck As VbMsgBoxResult
        '    '
        '    '
        '
        With invDoc
            ' Get Property s
            'With .Propertys
            '             aiPropsUser = .Item(gnCustom)
            ' aiPropsDesign = .Item(gnDesign)
            'End With

            ' Get Custom Properties
            '         prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
            '         prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
            '         prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)

            ' Family property is from Design, NOT Custom 
            ' prFamily = aiGetProp(aiPropsDesign, pnFamily)
            ' prPartNum = aiGetProp(aiPropsDesign, pnPartNum)
            'nmFamily = prFamily.Text
            nmFamily = .Propertys(gnDesign).Item(pnFamily).Text

            If .ComponentDefinition.IsContentMember Then
                If .ComponentDefinition.BOMStructure _
            = BOMStructureEnum.kPurchasedBOMStructure Then
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                    nmFamily = "D-HDWR"
                Else
                    Stop
                End If
            Else 'Try to identify other purchased part
                If InStr(1, invDoc.FullFileName,
                "\Doyle_Vault\Designs\purchased\"
            ) > 0 Then
                    'this is PROBABLY a purchased part
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                ElseIf g0f0(nmFamily) = BOMStructureEnum.kPurchasedBOMStructure Then
                    'this is almost certainly a purchased part
                    bomStruct = BOMStructureEnum.kPurchasedBOMStructure
                Else 'we'll assume it's NOT purchased.
                    bomStruct = BOMStructureEnum.kNormalBOMStructure
                    'If .SubType = guidSheetMetal Then
                    'It's PROBABLY sheet metal, BUT
                    'it might be something else
                    'Else
                    'End If
                End If
            End If
            ''
            '''

            '        With .ComponentDefinition
            '            '''
            '            'Get GeniusMass
            '            With .MassProperties
            '                 rt = dcWithProp( _
            '                    aiPropsUser, pnMass, _
            '                    cvMassKg2LbM * .Mass, rt _
            '                )
            '            End With
            '            'Will want to move this elsewhere
            '            'Should focus here on categorization
            'End With
        End With
        Test02 = bomStruct
        '
        '
        Debug.Print("") 'Landing Line for Debug use. Do not disable.
        '
        '
        '
    End Function

    Private Function g0f0(f As String) As Inventor.BOMStructureEnum
        With rsFamBuy
            .Filter = "Family = '" & f & "'"
            If .BOF Or .EOF Then
                g0f0 = BOMStructureEnum.kDefaultBOMStructure
            Else
                g0f0 = BOMStructureEnum.kPurchasedBOMStructure
            End If
        End With
    End Function

End Class