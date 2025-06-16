Class IfcDatum

    Private obNow As Object
    Private valNow As Object
    Private valWas As Object
    '
    Public Function Connect(This As Object) As IfcDatum
        If This Is Nothing Then
            Connect = Me
        ElseIf TypeOf This Is IfcDatum Then
            Connect = This
        Else
            Connect = obIfcDatum(This)
        End If
        Return Connect ' Add this line
    End Function

    Public Function MakeValue(This As Object) As IfcDatum
        MakeValue = Me
    End Function

    Public Function Commit() As IfcDatum
        Commit = Me
    End Function

    Public Function Itself() As IfcDatum
        Itself = Me
    End Function

    Public Function Connected(
    Optional ToThis As Object = Nothing
) As Boolean
        If ToThis Is Nothing Then
            Connected = Not obNow Is Nothing
        Else
            Connected = obNow Is ToThis
        End If
    End Function

    Public Function Value() As Object
        Value = valNow
    End Function

    Private Sub Class_Initialize()
        valNow = ""
    End Sub
End Class