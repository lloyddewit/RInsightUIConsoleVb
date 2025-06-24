Public Class UIElement
    Public ReadOnly Property strElementName As String
    Public ReadOnly Property strLabel As String
    Public Property lstChildren As New List(Of UIElement)

    Private ReadOnly strResetDefault As String

    Public Sub New(strName As String, Optional defaultValue As String = "", Optional labelValue As String = "")
        strElementName = strName
        strResetDefault = defaultValue
        strLabel = labelValue
    End Sub

    ''' <summary>
    ''' Returns a string signature of this node and all its descendants, with each strElementName separated by ', '.
    ''' </summary>
    Public ReadOnly Property strSignature As String
        Get
            Dim names As New List(Of String) From {strElementName}
            AddChildNames(lstChildren, names)
            Return String.Join(", ", names)
        End Get
    End Property

    ''' <summary>
    ''' Returns the default value as a string. Can be overridden by child classes.
    ''' </summary>
    Public Overridable ReadOnly Property defaultAsString As String
        Get
            Return strResetDefault
        End Get
    End Property

    Private Sub AddChildNames(children As List(Of UIElement), names As List(Of String))
        For Each child In children
            names.Add(child.strElementName)
            If child.lstChildren IsNot Nothing AndAlso child.lstChildren.Count > 0 Then
                AddChildNames(child.lstChildren, names)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Returns a deep copy of this UIElement and all its descendants.
    ''' </summary>
    Public Function Clone() As UIElement
        Dim clonedElement As New UIElement(Me.strElementName)
        For Each child In Me.lstChildren
            clonedElement.lstChildren.Add(child.Clone())
        Next
        Return clonedElement
    End Function

End Class

''' <summary>
''' A UIElement that has a boolean default value.
''' </summary>
Public Class UIElementBoolean
    Inherits UIElement

    ''' <summary>
    ''' The default boolean value for this element.
    ''' </summary>
    Private ReadOnly bResetDefault As Boolean

    ''' <summary>
    ''' Returns "True" if default is True, otherwise "False".
    ''' </summary>
    Public Overrides ReadOnly Property DefaultAsString As String
        Get
            Return If(bResetDefault, "True", "False")
        End Get
    End Property

    Public Sub New(strName As String, Optional defaultValue As Boolean = False)
        MyBase.New(strName)
        bResetDefault = defaultValue
    End Sub
End Class

''' <summary>
''' Abstract base class for numeric UI elements. Requires child classes to implement min, max, and increment properties.
''' </summary>
Public MustInherit Class UIElementNumber
    Inherits UIElement

    Protected Sub New(strName As String, Optional defaultValue As String = "")
        MyBase.New(strName, defaultValue)
    End Sub

    ''' <summary>
    ''' The minimum value allowed for this element.
    ''' </summary>
    Public MustOverride ReadOnly Property min As Double

    ''' <summary>
    ''' The maximum value allowed for this element.
    ''' </summary>
    Public MustOverride ReadOnly Property max As Double

    ''' <summary>
    ''' The increment step for this element.
    ''' </summary>
    Public MustOverride ReadOnly Property increment As Double
End Class

''' <summary>
''' Represents a UI element for integer values, with required min, max, and increment properties.
''' </summary>
Public Class UIElementInteger
    Inherits UIElementNumber

    Private ReadOnly _min As Integer
    Private ReadOnly _max As Integer
    Private ReadOnly _increment As Integer

    Public Sub New(strName As String, minValue As Integer, maxValue As Integer, incrementValue As Integer, Optional defaultValue As String = "")
        MyBase.New(strName, defaultValue)
        _min = minValue
        _max = maxValue
        _increment = incrementValue
    End Sub

    ''' <summary>
    ''' The minimum integer value allowed for this element.
    ''' </summary>
    Public Overrides ReadOnly Property min As Double
        Get
            Return _min
        End Get
    End Property

    ''' <summary>
    ''' The maximum integer value allowed for this element.
    ''' </summary>
    Public Overrides ReadOnly Property max As Double
        Get
            Return _max
        End Get
    End Property

    ''' <summary>
    ''' The increment step for this element (integer).
    ''' </summary>
    Public Overrides ReadOnly Property increment As Double
        Get
            Return _increment
        End Get
    End Property
End Class

''' <summary>
''' Represents a UI element for real (double) values, with required min, max, and increment properties.
''' </summary>
Public Class UIElementReal
    Inherits UIElementNumber

    Private ReadOnly _min As Double
    Private ReadOnly _max As Double
    Private ReadOnly _increment As Double

    Public Sub New(strName As String, minValue As Double, maxValue As Double, incrementValue As Double, Optional defaultValue As String = "")
        MyBase.New(strName, defaultValue)
        _min = minValue
        _max = maxValue
        _increment = incrementValue
    End Sub

    ''' <summary>
    ''' The minimum double value allowed for this element.
    ''' </summary>
    Public Overrides ReadOnly Property min As Double
        Get
            Return _min
        End Get
    End Property

    ''' <summary>
    ''' The maximum double value allowed for this element.
    ''' </summary>
    Public Overrides ReadOnly Property max As Double
        Get
            Return _max
        End Get
    End Property

    ''' <summary>
    ''' The increment step for this element (double).
    ''' </summary>
    Public Overrides ReadOnly Property increment As Double
        Get
            Return _increment
        End Get
    End Property
End Class
