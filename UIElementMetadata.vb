Public Class UIElementMetadata
    Public Property Label As String = ""
    Public Property ResetDefault As String = ""
End Class


Public Class UIElementMetadataBoolean
    Inherits UIElementMetadata
    'todo: ensure ResetDeefault is a boolean string ("true" or "false")
End Class

Public Class UIElementMetadataNumber
    Inherits UIElementMetadata

    Public Property Min As Double = 0
    Public Property Max As Double = 0
    Public Property Increment As Double = 0
    Public Property IsInteger As Boolean = False
End Class

Public Enum OptionType
    Constant
    Dataframe
    Column
    Other
End Enum

Public Class UIElementMetadataEnumeration
    Inherits UIElementMetadata

    Public Property Options As New List(Of String)
    Public Property OptionType As OptionType

    'todo ensure ResetDefault is one of the Options
End Class
