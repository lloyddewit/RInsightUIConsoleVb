''' <summary>
''' Represents metadata for a UI element (e.g. a text box)
''' This class serves as a base class for different types of UI element metadata,
''' such as boolean, number, and enumeration metadata.
''' </summary>
Public Class UIElementMetadata
    Public Property Label As String = ""
    Public Property ResetDefault As String = ""

    ''' <summary>
    ''' Checks if the ResetDefault value is valid for this metadata.
    ''' This method should be overridden in derived classes to provide specific validation logic.
    ''' </summary>
    ''' <returns>True if the data members are not nothing</returns>
    Public Overridable Function IsResetDefaultValid() As Boolean
        Return ResetDefault IsNot Nothing AndAlso Label IsNot Nothing
    End Function
End Class

''' <summary>
''' Represents metadata for a boolean UI element (e.g. a check box)
''' </summary>
Public Class UIElementMetadataBoolean
    Inherits UIElementMetadata

    ''' <summary>
    ''' Checks if the ResetDefault value is valid for a boolean metadata.
    ''' </summary>
    ''' <returns>true if ResetDefault is either "True" or "False"</returns>
    Public Overrides Function IsResetDefaultValid() As Boolean
        Return ResetDefault.Equals("True") OrElse
               ResetDefault.Equals("False")
    End Function
End Class

''' <summary>
''' Represents metadata for a numeric UI element (e.g. a NUD)
''' </summary>
Public Class UIElementMetadataNumber
    Inherits UIElementMetadata

    Public Property Min As Double = 0
    Public Property Max As Double = 0
    Public Property Increment As Double = 0
    Public Property IsInteger As Boolean = False

    ''' <summary>
    ''' Checks if the ResetDefault value is valid for a number metadata.
    ''' </summary>
    ''' <returns>True if ResetDefault is a number within the specified Min and Max range,
    ''' and if IsInteger is True, it should be an integer.</returns>
    Public Overrides Function IsResetDefaultValid() As Boolean
        Dim value As Double
        If Double.TryParse(ResetDefault, value) Then
            Return value >= Min AndAlso value <= Max AndAlso
                   (Not IsInteger OrElse value = Math.Floor(value))
        End If
        Return False
    End Function
End Class

''' <summary>
''' Shows if the enumeration should represent constants (e.g. days of the week), a dataframe, 
''' a column, or something else.
''' </summary>
Public Enum OptionType
    Constant
    Dataframe
    Column
    Other
End Enum

''' <summary>
''' Represents metadata for an enumeration UI element (e.g. a dataframe selector)
''' </summary>
Public Class UIElementMetadataEnumeration
    Inherits UIElementMetadata

    Public Property Options As New List(Of String)
    Public Property OptionType As OptionType

    ''' <summary>
    ''' Checks if the ResetDefault value is valid for an enumeration metadata.
    ''' </summary>
    ''' <returns>True if ResetDefault is one of the options. Also returns true if the options 
    '''     list is empty and ResetDefault is null or empty.</returns>
    Public Overrides Function IsResetDefaultValid() As Boolean
        If Options Is Nothing OrElse Options.Count = 0 Then
            Return String.IsNullOrEmpty(ResetDefault)
        End If

        Return Options.Contains(ResetDefault)
    End Function
End Class