Class gnsIfcAiPart_V0_2

    Inherits GnsIfcAiDoc

    Private Function gnsIfcAiDoc_Props(
    AiDoc As Inventor.Document,
    Optional dc As Scripting.IDictionary = Nothing
) As Scripting.IDictionary
        gnsIfcAiDoc_Props =
    gnsIfcAiDoc_Props.dcGeniusPropsPartRev20180530(
    aiDocPart(AiDoc), dc)
    End Function

End Class