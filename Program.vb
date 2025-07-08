Module Program
    Sub Main(args As String())
        Dim traitCorrelations As UIElement = UISpecBuilder.GetUISpec("TraitCorrelations")
        UISpecBuilder.WriteJsonFile(traitCorrelations, "TraitCorrelations")
    End Sub
End Module
