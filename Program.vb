Imports System.IO
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Module Program
    Sub Main(args As String())
        Dim jsonSerializerSettings As New JsonSerializerSettings With {
            .TypeNameHandling = TypeNameHandling.Auto,
            .Formatting = Formatting.Indented
        }
        jsonSerializerSettings.Converters.Add(New StringEnumConverter())
        Dim json As String = JsonConvert.SerializeObject(New Dictionary(Of String, UIElementMetadata) From {
            {"Base", New UIElementMetadata With {.Label = "Base Label", .ResetDefault = "default"}},
            {"Boolean", New UIElementMetadataBoolean With {.Label = "Boolean Label", .ResetDefault = "true"}},
            {"Integer", New UIElementMetadataNumber With {.Label = "Integer Label", .ResetDefault = "10", .Min = 0, .Max = 100, .Increment = 1, .IsInteger = True}},
            {"Enum", New UIElementMetadataEnumeration With {.Label = "Enum Label", .ResetDefault = "OptionA", .Options = New List(Of String) From {"OptionA", "OptionB", "OptionC"}, .OptionType = OptionType.Dataframe}}
        }, jsonSerializerSettings)

        ' Write to Desktop\tmp\elementTree.json
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim tmpFolder As String = Path.Combine(desktopPath, "tmp")
        Dim jsonFilePath As String = Path.Combine(tmpFolder, "testMetadata.json")
        File.WriteAllText(jsonFilePath, json)

        'deserialize the JSON back to a dictionary
        Dim deserializedDict As Dictionary(Of String, UIElementMetadata) = JsonConvert.DeserializeObject(Of Dictionary(Of String, UIElementMetadata))(File.ReadAllText(jsonFilePath), jsonSerializerSettings)

        'UISpecBuilder.WriteUISpec()
    End Sub
End Module
