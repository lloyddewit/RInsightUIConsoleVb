Imports System.IO
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports Newtonsoft.Json

Module Program
    Sub Main(args As String())

        Dim elements As New List(Of UIElementMetadata)
        Dim UIElementMetadata As New UIElementMetadata With {
            .Label = "Base Label",
            .ResetDefault = "default"
        }
        Dim UIElementMetadataBoolean As New UIElementMetadataBoolean With {
            .Label = "Boolean Label",
            .ResetDefault = "true"
        }
        Dim UIElementMetadataNumber As New UIElementMetadataNumber With {
            .Label = "Integer Label",
            .ResetDefault = "10",
            .Min = 0,
            .Max = 100,
            .Increment = 1,
            .IsInteger = True
        }
        Dim UIElementMetadataEnumeration As New UIElementMetadataEnumeration With {
            .Label = "Enum Label",
            .ResetDefault = "OptionA",
            .Options = New List(Of String) From {"OptionA", "OptionB", "OptionC"},
            .OptionType = OptionType.Dataframe
        }

        'Base: Base Label, Default: default
        'Boolean: Boolean Label, Default: true
        'Integer: Integer Label, Default: 10, Min: 0, Max: 100, Increment: 1, IsInteger: True
        'Enum: Enum Label, Default: OptionA, Options: OptionA, OptionB, OptionC, OptionType: Dataframe

        elements.Add(UIElementMetadata)
        elements.Add(UIElementMetadataBoolean)
        elements.Add(UIElementMetadataNumber)
        elements.Add(UIElementMetadataEnumeration)

        ' Serialize to JSON
        Dim settings As New JsonSerializerSettings With {.TypeNameHandling = TypeNameHandling.Auto}
        Dim json As String = JsonConvert.SerializeObject(elements, Formatting.Indented, settings)

        ' Write to Desktop\tmp\elementTree.json
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim tmpFolder As String = Path.Combine(desktopPath, "tmp")
        Dim jsonFilePath As String = Path.Combine(tmpFolder, "elementTree.json")
        File.WriteAllText(jsonFilePath, json)

        'deserialize the JSON back into a list of UIElement objects
        Dim lstDialogTypes As List(Of UIElementMetadata) = JsonConvert.DeserializeObject(Of List(Of UIElementMetadata))(File.ReadAllText(jsonFilePath), settings)
        For Each el In lstDialogTypes
            If TypeOf el Is UIElementMetadataBoolean Then
                Dim boolEl = DirectCast(el, UIElementMetadataBoolean)
                Console.WriteLine($"Boolean: {boolEl.Label}, Default: {boolEl.ResetDefault}")
            ElseIf TypeOf el Is UIElementMetadataNumber Then
                Dim intEl = DirectCast(el, UIElementMetadataNumber)
                Console.WriteLine($"Integer: {intEl.Label}, Default: {intEl.ResetDefault}, Min: {intEl.Min}, Max: {intEl.Max}, Increment: {intEl.Increment}, IsInteger: {intEl.IsInteger}")
            ElseIf TypeOf el Is UIElementMetadataEnumeration Then
                Dim enumEl = DirectCast(el, UIElementMetadataEnumeration)
                Console.WriteLine($"Enum: {enumEl.Label}, Default: {enumEl.ResetDefault}, Options: {String.Join(", ", enumEl.Options)}, OptionType: {enumEl.OptionType}")
            Else
                Console.WriteLine($"Base: {el.Label}, Default: {el.ResetDefault}")
            End If
        Next

        'UISpecBuilder.WriteUISpec()
    End Sub
End Module
