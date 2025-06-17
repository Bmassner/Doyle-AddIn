Class ifcAiProperty

    Inherits IfcDatum

    'Private PropertySet As Inventor.Property
    Private InvProperty As Inventor.Property
    'Private nm As String
    Private vlWas As Object
    Private vlNow As Object

    Public Overloads Function Connect(
    ToProp As Inventor.Property
    ) As IfcDatum
        InvProperty = ToProp
        If ToProp Is Nothing Then
            ' PropertySet = Nothing
            'nm = ""
            vlWas = vbEmpty
            ''  not sure what's best here
            ''  be prepared to change
            ''  as necessary
        Else
            With InvProperty
                ' PropertySet = .Parent
                'nm = .Name
                vlWas = .Text
            End With
        End If
        vlNow = vlWas

        Connect = Me
    End Function
    ' replaces disabled function below
    '
    'Public Function AttachedTo(Name As String, _
    '    Optional InProp As Inventor.Property = Nothing _
    ') As ifcAiProperty
    '    '
    '    '
    '    '
    '    nm = Name
    '
    '    If Not InProp Is Nothing Then
    '         PropertySet = InProp
    '    End If
    '
    '    If Not PropertySet Is Nothing Then
    '        On Error Resume Next
    '         Property = PropertySet.Item(nm)
    '        If Err.Number = 0 Then
    '            vlWas = Property.Text
    '        Else
    '             Property = Nothing
    '            vlWas = Empty
    '        End If
    '        On Error GoTo 0
    '    End If
    '
    '     AttachedTo = Me
    'End Function

    Public Overloads Function MakeValue(
    This As Object
    ) As IfcDatum
        If TypeOf This Is Object Then
            ' really should NOT support this
            ' but will let stand for now
            ' vlNow = This
            ' no. will opt for
            ' 'do nothing' instead
        Else
            vlNow = This
        End If

        MakeValue = Me
    End Function
    ' replaces disabled function below
    '
    'Public Function WithValue( _
    '    NewVal AsObject_
    ') As ifcAiProperty
    '    Me.Text = NewVal
    '
    '     WithValue = Me
    'End Function

    Public Overloads Function Commit() As ifcAiProperty
        'Dim PropertySet As Inventor.Property
        'Dim ck As Object

        If vlWas Is Nothing Then
            Stop
            'don't know about clearing Property
        Else
            If InvProperty Is Nothing Then
                'If PropertySet Is Nothing Then
                'can't do anything
                'Else
                ' Property = PropertySet.Add(vlWas, nm)
                'If Property Is Nothing Then
                '     Commit = Me
                'Else
                '     Commit = Commit()
                'End If
                'End If
            Else
                vlWas = InvProperty.Text 'ck

                If vlNow = vlWas Then 'ck
                    'shouldn't need updated
                ElseIf CStr(vlNow) = CStr(vlWas) Then 'ck
                    'PROBABLY shouldn't need updated
                Else
                    On Error Resume Next

                    InvProperty.Text = vlNow 'vlWas
                    If Err.Number = 0 Then 'should be okay
                        'don't worry about it
                    Else 'something went wrong
                        Stop 'and see what we can do
                    End If

                    On Error GoTo 0
                End If
            End If
        End If

        Commit = Me
    End Function

    Private Overloads Function Itself() As IfcDatum
        Itself = Me
    End Function
    ' replaces disabled function below
    '
    'Public Function Obj() As ifcAiProperty
    '     Obj = Me
    'End Function

    Private Overloads Function Connected(
    Optional ToThis As Inventor.Property = Nothing
    ) As Boolean
        If ToThis Is Nothing Then
            Connected = Not InvProperty Is Nothing
        Else
            Connected = InvProperty Is ToThis
        End If
    End Function

    Private Overloads Function Value() As Object
        If TypeOf vlNow Is Object Then
            ' this should NOT ever happen
            ' but just to be robust...
            Value = vlNow
        Else
            Value = vlNow
        End If
    End Function

    Public Function Status() As Long
        Status = -1
    End Function

    Public Function Name() As Object
        If InvProperty Is Nothing Then
            Name = InvProperty.Name
        Else
            Name = ""
        End If
    End Function

    'Public Property Get Value() As Object
    '    Value = vlWas
    'End Property
    '
    'Public Property Let Value(NewVal As Object)
    '    If IsEmpty(NewVal) Then
    '        Stop
    '    'ElseIf IsNull(NewVal) Then
    '    'ElseIf IsMissing(NewVal) Then
    '    ElseIf IsObject(NewVal) Then
    '        Stop
    '    Else
    '        vlWas = NewVal
    '    End If
    'End Property

    Private Sub Class_Initialize()
        'nm = ""
        vlWas = vbEmpty
        ' PropertySet = Nothing
        InvProperty = Nothing
    End Sub

    Private Sub Class_Terminate()
        'If PropertySet Is Nothing Then 'nowhere to save
        'so nothing to do but drop it
        'Else
        If InvProperty Is Nothing Then 'we need
            'to create InvProperty, if desired
            '(and possible)
            Stop
        ElseIf vlWas = InvProperty.Text Then 'no change
            'so nothing needs doing
        ElseIf CStr(vlWas) = CStr(InvProperty.Text) Then
            'likely no REAL change
        Else 'value has changed
            'and MIGHT need to be committed
            Stop
        End If
    End Sub

    Private Function ifcDatum_Commit() As IfcDatum
        '
    End Function

    Private Function ifcDatum_Connect(
    ToThis As Object
) As IfcDatum
        ifcDatum_Connect =
    Connect(obAiProp(ToThis))
    End Function

    Private Function ifcDatum_Connected(
    Optional ToThis As Object = Nothing
) As Boolean
        ifcDatum_Connected =
    Connected(obAiProp(ToThis))
    End Function

    Private Function ifcDatum_Itself() As IfcDatum
        ifcDatum_Itself = Me
    End Function

    Private Function ifcDatum_MakeValue(
    This As Object
) As IfcDatum
        ifcDatum_MakeValue _
    = MakeValue(This)
    End Function

    Private Function ifcDatum_Value() As Object
        If TypeOf vlNow Is Object Then
            ' this should NOT ever happen
            ' but just to be robust...
            ifcDatum_Value = vlNow
        Else
            ifcDatum_Value = vlNow
        End If
    End Function
End Class