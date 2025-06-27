Public Class UIElement
    Public ReadOnly Property ElementName As String = ""
    Public Property Children As New List(Of UIElement)
    Public Property UIElementMetadata As UIElementMetadata = Nothing

    Public Sub New(strName As String)
        ElementName = strName
    End Sub

    ''' <summary>
    ''' Returns a string signature of this node and all its descendants, with each strElementName separated by ', '.
    ''' </summary>
    Public ReadOnly Property Signature As String
        Get
            Dim names As New List(Of String) From {ElementName}
            AddChildNames(Children, names)
            Return String.Join(", ", names)
        End Get
    End Property

    Private Sub AddChildNames(children As List(Of UIElement), names As List(Of String))
        For Each child In children
            names.Add(child.ElementName)
            If child.Children IsNot Nothing AndAlso child.Children.Count > 0 Then
                AddChildNames(child.Children, names)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Returns a deep copy of this UIElement and all its descendants.
    ''' </summary>
    Public Function Clone() As UIElement
        Dim clonedElement As New UIElement(ElementName)
        For Each child In Children
            clonedElement.Children.Add(child.Clone())
        Next
        clonedElement.UIElementMetadata = If(UIElementMetadata?.Clone(), Nothing)
        Return clonedElement
    End Function

End Class
