Public Class UIElementMetadata
    Public Property Label As String = ""
    Public Property ResetDefault As String = ""
    Public Function Clone() As UIElementMetadata
        Dim clonedMetadata As New UIElementMetadata With {
            .Label = Label,
            .ResetDefault = ResetDefault
        }
        Return clonedMetadata
    End Function
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

    Public Overloads Function Clone() As UIElementMetadataNumber
        Dim clonedMetadata As New UIElementMetadataNumber With {
            .Label = Label,
            .ResetDefault = ResetDefault,
            .Min = Min,
            .Max = Max,
            .Increment = Increment,
            .IsInteger = IsInteger
        }
        Return clonedMetadata
    End Function
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

    Public Overloads Function Clone() As UIElementMetadataEnumeration
        Dim clonedMetadata As New UIElementMetadataEnumeration With {
            .Label = Label,
            .ResetDefault = ResetDefault,
            .Options = New List(Of String)(Options),
            .OptionType = OptionType
        }
        Return clonedMetadata
    End Function
End Class
