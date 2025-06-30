''' <summary>
''' Represents a UI element in a tree structure.
''' Each UIElement has a name and can have multiple children, forming a hierarchy.
''' The ElementName is used to create a unique signature for the element and its descendants.
''' </summary>
Public Class UIElement

    ''' <summary>
    ''' The name of the UIElement, which is used to identify it in the tree structure.
    ''' It is used to create a signature for the element and its descendants.
    ''' </summary>
    Public ReadOnly Property ElementName As String

    ''' <summary>
    ''' Represents a tree structure where each UIElement can have multiple children.
    ''' </summary>
    Public Property Children As List(Of UIElement) = New List(Of UIElement)

    Public Sub New(strName As String)
        ElementName = strName
    End Sub

    ''' <summary>
    ''' Returns a string signature of this node and all its descendants, with each strElementName 
    ''' separated by ', '.
    ''' </summary>
    Public ReadOnly Property Signature As String
        Get
            Dim names As New List(Of String) From {ElementName}
            names.AddRange(GetChildNames(Children))
            Return String.Join(", ", names)
        End Get
    End Property

    ''' <summary>
    ''' Recursively collects the ElementName of all descendants into a flat list.
    ''' </summary>
    Private Function GetChildNames(children As List(Of UIElement)) As List(Of String)
        Dim names As New List(Of String)
        If children Is Nothing OrElse children.Count = 0 Then Return names

        For Each child In children
            names.Add(child.ElementName)
            If child.Children IsNot Nothing AndAlso child.Children.Count > 0 Then
                names.AddRange(GetChildNames(child.Children))
            End If
        Next
        Return names
    End Function

    ''' <summary>
    ''' Returns a deep copy of this UIElement and all its descendants.
    ''' </summary>
    Public Function Clone() As UIElement
        Dim clonedElement As New UIElement(ElementName)

        If Children Is Nothing Then Return clonedElement

        For Each child In Children
            clonedElement.Children.Add(child.Clone())
        Next
        Return clonedElement
    End Function

End Class
