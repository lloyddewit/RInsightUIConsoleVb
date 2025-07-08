Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

''' <summary>
''' This class provides methods to build a UI specification for a dialog by reading 
''' transformations, building a UIElement tree, and recursively and iteratively removing duplicates 
''' until no more duplicates can be removed. The result is a minimal tree of UI elements that can 
''' be used to build a UI dialog.
''' </summary>
Public Class UISpecBuilder
    ' Prevent instantiation because this class is designed to only contain shared methods
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Writes a UI specification for a given dialog by reading transformations, building a 
    ''' UIElement tree, and then recursively and iteratively removing duplicates until no more 
    ''' duplicates can be removed. The result is a minimal tree of UI elements that can be used to 
    ''' build a UI dialog. Also writes debug information to a file.
    ''' </summary>
    ''' <param name="dialogName"> Creates the UI specification for this dialog</param>
    ''' <returns> A minimal tree of UI elements that can be used to build a UI dialog</returns>
    Public Shared Function GetUISpec(dialogName As String) As UIElement

        ' Build the initial UIElement tree
        Dim root As New UIElement("Root")
        Dim transformations As List(Of clsTransformationRModel) = GetTransformations(dialogName)
        For Each model In transformations
            Dim element = GetUIElementTree(model)
            If element IsNot Nothing Then
                root.Children.Add(element)
            End If
        Next

        ' Remove duplicate siblings
        Dim rootNoDuplicateSiblings = GetUIElementNoDuplicateSiblings(root)

        ' Remove nodes that are duplicates of their ancestors
        Dim rootNoDuplicateAncestors = GetUIElementNoDuplicateAncestors(
            rootNoDuplicateSiblings, Nothing, Nothing)

        ' Iteratively move duplicates to the LCA (Lowest Common Ancestor)
        Dim rootDuplicatesInLca As UIElement = Nothing
        Dim rootNoDuplicateAncestorsFinal As UIElement
        Do
            ' If there are duplicates in different branches, then find the duplicate tree with
            '    the longest signature, and add it to the LCA
            rootDuplicatesInLca = GetUIElementAddDuplicatesToLca(rootNoDuplicateAncestors)
            rootNoDuplicateAncestorsFinal = rootNoDuplicateAncestors.Clone()

            ' Remove nodes that are duplicates of their ancestors
            rootNoDuplicateAncestors = GetUIElementNoDuplicateAncestors(
                rootDuplicatesInLca, Nothing, Nothing)

            ' Keep repeating until no more duplicates are found
        Loop Until rootDuplicatesInLca Is Nothing

        ' Remove duplicate true/false siblings
        Dim rootNoDuplicateSiblingsTrueFalse = GetUIElementNoDuplicateSiblings(
            rootNoDuplicateAncestorsFinal, True)

        ' Remove true/false nodes that are duplicates of their ancestors
        Dim rootNoDuplicateAncestorsTrueFalse = GetUIElementNoDuplicateAncestors(
            rootNoDuplicateSiblingsTrueFalse, Nothing, Nothing, True)

        ' Read and verify metadata
        Dim metadata As Dictionary(Of String, UIElementMetadata) =
            GetMetadata(dialogName)

        ' For debugging, write the element tree to a file on the desktop
        Dim debugRoots As Dictionary(Of String, UIElement) =
            New Dictionary(Of String, UIElement) From {
            {"Before duplicate removal:", root},
            {"After duplicate sibling removal:", rootNoDuplicateSiblings},
            {"After duplicate ancestor removal:", rootNoDuplicateAncestors},
            {"After creating and cleaning LCA tree:", rootNoDuplicateAncestorsFinal},
            {"After duplicate true/false sibling removal:", rootNoDuplicateSiblingsTrueFalse},
            {"After true/false duplicate ancestor removal:", rootNoDuplicateAncestorsTrueFalse}
        }
        WriteDebugFile(debugRoots, metadata)

        Return rootNoDuplicateAncestorsTrueFalse
    End Function

    ''' <summary>
    ''' Serializes the given UIElement tree to JSON and writes it to a file in the tmp folder on 
    ''' the desktop.
    ''' </summary>
    ''' <param name="root">The root of the UIElement tree to serialize.</param>
    ''' <param name="dialogName">The name of the dialog, used to create the filename.</param>
    ''' <exception cref="ArgumentNullException">Thrown if the root is null.</exception>
    Public Shared Sub WriteJsonFile(root As UIElement, dialogName As String)
        If root Is Nothing Then Throw New ArgumentNullException(NameOf(root))

        Dim settings As New JsonSerializerSettings With {
            .TypeNameHandling = TypeNameHandling.Auto,
            .Formatting = Formatting.Indented}
        settings.Converters.Add(New StringEnumConverter())

        Dim json As String = JsonConvert.SerializeObject(root, settings)

        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim tmpFolder As String = Path.Combine(desktopPath, "tmp")
        Dim jsonPath As String = Path.Combine(tmpFolder, "dlg" & dialogName & "UIElements.json")

        If Not Directory.Exists(tmpFolder) Then
            Directory.CreateDirectory(tmpFolder)
        End If

        File.WriteAllText(jsonPath, json)
    End Sub

    ''' <summary>
    ''' Returns a string representation of the UIElement tree structure, with indentation for 
    ''' each level.
    ''' </summary>
    ''' <param name="element">The UIElement to start from.</param>
    ''' <param name="level">The current level in the tree (used for indentation).</param>
    ''' <param name="title">Title for the output.</param>
    ''' <param name="metadataDict">Optional metadata dictionary for additional information.</param>
    ''' <returns>A string representation of the UIElement tree.</returns>
    Private Shared Function GetDebugInfo(element As UIElement, level As Integer, title As String,
            Optional metadataDict As Dictionary(Of String, UIElementMetadata) = Nothing) As String
        If element Is Nothing Then Return String.Empty

        Dim sb As New Text.StringBuilder()
        If Not String.IsNullOrEmpty(title) Then
            sb.AppendLine(vbLf & vbLf & title & vbLf)
        End If

        Dim elementName As String = element.ElementName
        If Not String.IsNullOrEmpty(elementName) Then

            If elementName.EndsWith(" F") Then
                elementName = elementName.Substring(0, elementName.Length - 2)
            End If

            Dim metadataString As String = ""
            If metadataDict IsNot Nothing AndAlso metadataDict.ContainsKey(elementName) Then

                Dim metadata As UIElementMetadata = metadataDict(elementName)

                metadataString = $"{metadata.Label}, {metadata.ResetDefault}"
                If TypeOf metadata Is UIElementMetadataNumber Then
                    Dim numMetadata As UIElementMetadataNumber =
                        CType(metadata, UIElementMetadataNumber)
                    metadataString &= $", Min: {numMetadata.Min}, Max: {numMetadata.Max}, " &
                        $"Increment: {numMetadata.Increment}, IsInteger: {numMetadata.IsInteger}"
                ElseIf TypeOf metadata Is UIElementMetadataEnumeration Then
                    Dim enumMetadata As UIElementMetadataEnumeration =
                        CType(metadata, UIElementMetadataEnumeration)
                    metadataString &= $", Options: {String.Join(", ", enumMetadata.Options)}, " &
                        $"OptionType: {enumMetadata.OptionType}"
                End If
            End If

            sb.AppendLine($"{New String(" "c, level * 2)}Level {level}: {element.ElementName}, " &
                $"{metadataString}")
        End If

        For Each child In element.Children
            sb.Append(GetDebugInfo(child, level + 1, Nothing, metadataDict))
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Finds the LCA (Lowest Common Ancestor) element from the <paramref name="root"/> element of 
    ''' a cloned UIElement tree. It finds the LCA element using <paramref name="path"/> and 
    ''' <paramref name="lcaElement"/>.
    ''' </summary>
    ''' <remarks>
    ''' Both <paramref name="path"/> and <paramref name="lcaElement"/> should come from the 
    ''' original tree (i.e. before the tree was cloned). The method traverses the cloned tree 
    ''' according to the path and returns the equivalent LCA element in the cloned tree. This 
    ''' method assumes that the path is a valid path in the cloned tree and that the LCA element
    ''' exists in the cloned tree. If the path does not lead to the specified LCA element in the 
    ''' cloned tree, an exception is thrown.
    ''' </remarks>
    ''' <param name="root">The root of the cloned UIElement tree.</param>
    ''' <param name="path">The path to the LCA element in the original tree.</param>
    ''' <param name="lcaElement">The original LCA element.</param>
    ''' <returns>The LCA element in the cloned tree.</returns>
    ''' <exception cref="ArgumentNullException">Thrown if any of the parameters are null or 
    '''     invalid.</exception>
    ''' <exception cref="InvalidOperationException">Thrown if the path does not lead to the 
    '''     specified LCA element.</exception>
    Private Shared Function GetLcaElementFromClonedTree(root As UIElement, path As List(Of UIElement),
                                                        lcaElement As UIElement) As UIElement
        If path Is Nothing OrElse path.Count = 0 _
            OrElse root Is Nothing _
            OrElse lcaElement Is Nothing Then Throw New ArgumentNullException("Invalid parameter.")

        ' Follow the path through the cloned tree until we find the LCA element
        Dim current As UIElement = root
        For i As Integer = 1 To path.Count - 1
            ' Note: if the path elements and the LCA element both come from the same tree, then we
            '       could also do `If path(i - 1) Is lcaElement Then` but I found comparing the
            '       signatures was safer
            If path(i - 1).Signature = lcaElement.Signature Then
                Return current
            End If

            ' Get the next element in the path
            ' Note: We determine the next element in the path based on the node's signature.
            '       There should only be one node with this signature, because any duplicate
            '       siblings should have been removed in earlier steps.
            Dim signature = path(i).Signature
            current = current.Children.Single(Function(child) child.Signature = signature)
        Next

        Throw New InvalidOperationException("Path does not lead to the specified LCA element.")
    End Function

    ''' <summary>
    ''' Finds the lowest common ancestor (LCA) node given two paths in the UIElement tree.
    ''' </summary>
    ''' <remarks>
    ''' The paths are lists of UIElement nodes representing the path from the root to the target nodes.
    ''' The LCA is the deepest node that is an ancestor of both target nodes. 
    ''' </remarks>
    ''' <param name="path1">The first path in the UIElement tree.</param>
    ''' <param name="path2">The second path in the UIElement tree.</param>
    ''' <returns>The LCA node, or Nothing if no common ancestor is found.</returns>
    Private Shared Function GetLcaElementFromTwoPaths(path1 As List(Of UIElement),
                                   path2 As List(Of UIElement)) As UIElement
        If path1 Is Nothing OrElse path2 Is Nothing _
            OrElse path1.Count = 0 OrElse path2.Count = 0 Then Return Nothing

        Dim minLen = Math.Min(path1.Count, path2.Count)
        Dim lca As UIElement = Nothing
        For i = 0 To minLen - 1
            If path1(i) Is path2(i) Then
                lca = path1(i)
            Else
                Exit For
            End If
        Next
        Return lca
    End Function

    ''' <summary>
    ''' Reads the metadata for a dialog from a JSON file and deserializes it into a dictionary of 
    ''' UIElementMetadata objects.
    ''' </summary>
    ''' <remarks>
    ''' The metadata file is expected to be in a specific directory structure based on the dialog name.
    ''' </remarks>"
    ''' <param name="dialogName">The name of the dialog for which metadata is being read.</param>
    ''' <returns>A dictionary where keys are element names and values are their metadata.</returns>
    ''' <exception cref="FileNotFoundException">Thrown if the metadata file does not exist.</exception>
    ''' <exception cref="JsonException">Thrown if the JSON is invalid or cannot be deserialized.</exception>
    ''' <exception cref="ArgumentException">Thrown if the ResetDefault value is invalid for any 
    '''     UIElementMetadata.</exception>
    Private Shared Function GetMetadata(
                               dialogName As String) As Dictionary(Of String, UIElementMetadata)
        Dim dialogDefinitionsPath As String = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "assets", "DialogDefinitions")
        Dim metadataPath As String = Path.Combine(
            dialogDefinitionsPath, "Dlg" & dialogName, "dlg" & dialogName & "Metadata.json")

        Dim settings As New JsonSerializerSettings With {
            .TypeNameHandling = TypeNameHandling.Auto,
            .Formatting = Formatting.Indented}
        settings.Converters.Add(New StringEnumConverter())

        Dim metadata As Dictionary(Of String, UIElementMetadata) = JsonConvert.DeserializeObject(
            Of Dictionary(Of String, UIElementMetadata))(File.ReadAllText(metadataPath), settings)

        For Each kvp In metadata
            Dim value As UIElementMetadata = kvp.Value
            If Not value.IsResetDefaultValid() Then
                Throw New ArgumentException(
                    $"ResetDefault value '{value.ResetDefault}' is invalid for '{kvp.Key}' " &
                    $"({value.GetType().Name})")
            End If
        Next

        Return metadata
    End Function

    ''' <summary>
    ''' Some signatures contain " F" to indicate a false condition. This function removes any " F" 
    ''' suffixes from the signature. This function can be used to ignore variations of true/false 
    ''' when comparing signatures.
    ''' </summary>
    ''' <param name="strSignature">The signature to process.</param>
    ''' <returns>The processed signature without any " F".</returns>
    Private Shared Function GetSignatureIgnoreTrueFalse(strSignature As String) As String
        If strSignature.EndsWith(" F", StringComparison.Ordinal) Then
            strSignature = strSignature.Substring(0, strSignature.Length - 2)
        End If
        strSignature = strSignature.Replace(" F,", ",")
        Return strSignature
    End Function

    ''' <summary>
    ''' Reads the transformations for <paramref name="dialogName"/> from a JSON file and 
    ''' deserializes them into a list of clsTransformationRModel objects.
    ''' </summary>
    ''' <param name="dialogName">The name of the dialog for which transformations are being read.</param>
    ''' <returns>A list of clsTransformationRModel objects representing the transformations.</returns>
    Private Shared Function GetTransformations(dialogName As String) As List(Of clsTransformationRModel)

        Dim dialogDefinitionsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets",
                                                 "DialogDefinitions")
        Dim dialogPath As String = Path.Combine(
            dialogDefinitionsPath, "Dlg" & dialogName, "dlg" & dialogName & ".json")

        Dim transformationsJson As String = File.ReadAllText(dialogPath)
        Return JsonConvert.DeserializeObject(Of List(Of clsTransformationRModel))(transformationsJson)
    End Function

    ''' <summary>
    ''' Returns a new UIElement tree with one extra node added: if there are duplicate signatures 
    ''' in the tree, finds the duplicated node with the longest signature, finds any two nodes with 
    ''' this signature, finds their LCA, and adds a duplicate of the node as a child of the LCA. 
    ''' If no duplicates exist, returns Nothing.
    ''' </summary>
    ''' <param name="root"> The root UIElement of the tree to process.</param>
    ''' <returns> A new UIElement tree with the duplicate node added, or Nothing if no duplicates 
    '''           were found.</returns>
    Private Shared Function GetUIElementAddDuplicatesToLca(root As UIElement) As UIElement
        If root Is Nothing Then Return Nothing

        ' Create a dictionary of all paths to each signature
        Dim signaturePaths As New Dictionary(Of String, List(Of List(Of UIElement)))
        UpdateSignaturePaths(root, New List(Of UIElement), signaturePaths)

        ' Convert the dictionary to a list of key-value pairs, filter out signatures with only
        ' one path (i.e. only include duplicate signatures)
        Dim duplicatedSignatures As List(Of KeyValuePair(Of String, List(Of List(Of UIElement)))) =
            signaturePaths.Where(Function(kvp) kvp.Value.Count > 1).ToList()
        If duplicatedSignatures.Count = 0 Then
            Return Nothing
        End If

        ' Find the longest duplicated signature.
        ' Note: We want to process the largest/most complex duplicate subtree first, and the
        '       longest signature usually means the largest subtree. In theory, this is not
        '       always the case, but in practice it works well because our main concern is to
        '       prevent smaller duplicate subtrees within other duplicate subtrees being processed
        '       first (results in a sub-optimal tree).
        Dim longestDup = duplicatedSignatures.OrderByDescending(
            Function(kvp) kvp.Key.Length).First()

        ' Get the paths of the the first two nodes with this signature (the path is the list of
        ' the node's direct ancestors)
        Dim path1 As List(Of UIElement) = longestDup.Value(0)
        Dim path2 As List(Of UIElement) = longestDup.Value(1)

        ' Clone the node we want to add later to the LCA (node could come from either path because
        ' both nodes should be identical)
        Dim elementToAdd = path1.Last().Clone()

        ' Get the element that is the LCA element of the two paths
        Dim lcaElement As UIElement = GetLcaElementFromTwoPaths(path1, path2)
        If lcaElement Is Nothing Then Return Nothing

        ' Clone the tree so we can add nodes without mutating the original
        Dim newRootElement As UIElement = root.Clone()

        ' Find the LCA node in the cloned tree
        Dim lcaClone = GetLcaElementFromClonedTree(newRootElement, path1, lcaElement)
        If lcaClone Is Nothing Then Return newRootElement

        ' Add a clone of the duplicate node to the LCA's children (if not already present)
        If Not lcaClone.Children.Any(Function(child) child.Signature = elementToAdd.Signature) Then
            lcaClone.Children.Add(elementToAdd)
        End If

        Return newRootElement
    End Function

    ''' <summary>
    ''' Recursively removes nodes from the UIElement tree that are duplicates of their ancestors, 
    ''' or duplicates of their parent's siblings. Returns the root node of the new tree.
    ''' </summary>
    ''' <param name="element"> The tree's root element</param>
    ''' <param name="ancestorsSignatures"> A set of signatures of the ancestors of the current 
    '''                                    node</param>
    ''' <param name="parentSiblingsSignatures"> A set of signatures of the siblings of the parent 
    '''                                         node</param>
    ''' <param name="ignoreTrueFalse"> If true, ignore variations of true/false signatures</param>
    ''' <returns>The root element of a new UIElement tree without duplicate ancestors and 
    '''          duplicates of parent siblings.</returns>
    Private Shared Function GetUIElementNoDuplicateAncestors(element As UIElement,
                                        ancestorsSignatures As HashSet(Of String),
                                        parentSiblingsSignatures As HashSet(Of String),
                                        Optional ignoreTrueFalse As Boolean = False) As UIElement
        If element Is Nothing Then Return Nothing

        If ancestorsSignatures Is Nothing Then
            ancestorsSignatures = New HashSet(Of String)
        End If

        ' If this strSignature matches any ancestor, remove this node
        If IsSignatureInSet(element.Signature, ancestorsSignatures, ignoreTrueFalse) Then
            Return Nothing
        End If

        ' Make a list of ancestors for this node's children, include the siblings of this parent
        Dim ancestorsSignaturesNew As New HashSet(Of String)(ancestorsSignatures)
        If parentSiblingsSignatures IsNot Nothing Then
            For Each siblingSignature In parentSiblingsSignatures
                ancestorsSignaturesNew.Add(siblingSignature)
            Next
        End If

        ' Make a list of parent siblings for this nodes children
        Dim siblingSignatures As New HashSet(Of String)
        For Each child In element.Children
            siblingSignatures.Add(child.Signature)
        Next

        ' Recursively remove duplicate ancestors from the node's children
        Dim newChildren As New List(Of UIElement)
        For Each child In element.Children
            Dim cleanedChild = GetUIElementNoDuplicateAncestors(child, ancestorsSignaturesNew,
                                                                siblingSignatures, ignoreTrueFalse)
            If cleanedChild IsNot Nothing Then
                newChildren.Add(cleanedChild)
            End If
        Next

        ' Return a new node to avoid mutating the original
        Dim newElement As New UIElement(element.ElementName)
        newElement.Children = newChildren
        Return newElement
    End Function

    ''' <summary>
    ''' Recursively removes duplicate siblings from the UIElement tree.
    ''' This method ensures that no two siblings have the same signature, optionally ignoring
    ''' true/false variations of signatures.
    ''' </summary>
    ''' <param name="element"> The tree's root element</param>
    ''' <param name="ignoreTrueFalse"> If true, ignore variations of true/false signatures</param>
    ''' <returns>The root element of a UIElement tree without duplicate siblings</returns>
    Private Shared Function GetUIElementNoDuplicateSiblings(
            element As UIElement, Optional ignoreTrueFalse As Boolean = False) As UIElement

        If element Is Nothing Then Return Nothing

        ' Process children and remove duplicates among siblings
        Dim seenSignatures As New HashSet(Of String)
        Dim childrenParse1 As New List(Of UIElement)
        For Each child In element.Children
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, ignoreTrueFalse)
            If cleanedChild IsNot Nothing _
                    AndAlso Not IsSignatureInSet(
                        cleanedChild.Signature, seenSignatures, ignoreTrueFalse) Then
                seenSignatures.Add(cleanedChild.Signature)
                childrenParse1.Add(cleanedChild)
            End If
        Next

        ' Loop backwards through new children to see if we can remove any more siblings
        seenSignatures = New HashSet(Of String)
        Dim childrenParse2 As New List(Of UIElement)
        For iChildIndex As Integer = childrenParse1.Count - 1 To 0 Step -1
            Dim child = childrenParse1(iChildIndex)
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, ignoreTrueFalse)
            If cleanedChild IsNot Nothing _
                    AndAlso Not IsSignatureInSet(
                        cleanedChild.Signature, seenSignatures, ignoreTrueFalse) Then
                seenSignatures.Add(cleanedChild.Signature)
                childrenParse2.Add(cleanedChild)
            End If
        Next

        ' Create a new node to avoid mutating the original
        Dim newElement As New UIElement(element.ElementName)
        newElement.Children = childrenParse2
        Return newElement
    End Function

    ''' <summary>
    ''' Recursively builds a UIElement tree from an R transformation tree.
    ''' </summary>
    ''' <param name="transformation"> the tree of R transformations</param>
    ''' <returns>A tree of UI elements</returns>
    Private Shared Function GetUIElementTree(transformation As clsTransformationRModel) As UIElement
        If transformation Is Nothing Then Return Nothing

        If String.IsNullOrEmpty(transformation.strValueKey) AndAlso
            (transformation.lstTransformations Is Nothing OrElse
             transformation.lstTransformations.Count = 0) Then
            Return Nothing
        End If

        Dim strElementName As String = transformation.strValueKey
        strElementName += If(transformation.enumTransformationType =
            clsTransformationRModel.TransformationType.ifFalseExecuteChildTransformations, " F", "")

        Dim element As New UIElement(strElementName)
        If transformation.lstTransformations Is Nothing Then Return element

        For Each child In transformation.lstTransformations
            Dim childElement = GetUIElementTree(child)
            If childElement IsNot Nothing Then
                element.Children.Add(childElement)
            End If
        Next

        Return element
    End Function

    ''' <summary>
    ''' Returns True if <paramref name="signature"/> is in <paramref name="signaturesToCompare"/>,
    ''' or if <paramref name="signature"/> is a prefix of a signature in 
    ''' <paramref name="signaturesToCompare"/>.
    ''' </summary>
    ''' <param name="ignoreTrueFalse"> If true, ignore variations of true/false signatures</param>
    ''' <param name="signature"> The signature to look for</param>
    ''' <param name="signaturesToCompare">The set of signatures to compare against</param>
    ''' <returns>True if <paramref name="signature"/> is in
    '''     <paramref name="signaturesToCompare"/></returns>
    Private Shared Function IsSignatureInSet(signature As String,
                                      signaturesToCompare As HashSet(Of String),
                                      ignoreTrueFalse As Boolean) As Boolean

        ' Create a cleaned set of signatures to compare against (" F" removed if required)
        Dim signatureCleaned As String = If(ignoreTrueFalse,
                                            GetSignatureIgnoreTrueFalse(signature), signature)
        Dim signaturesToCompareCleaned = New HashSet(Of String)
        For Each sig In signaturesToCompare
            Dim sigCleaned As String = If(ignoreTrueFalse, GetSignatureIgnoreTrueFalse(sig), sig)
            signaturesToCompareCleaned.Add(sigCleaned)
        Next

        ' If the signature is identical to a signature in the set of signatures, return true
        If signaturesToCompareCleaned.Contains(signatureCleaned) Then
            Return True
        End If

        ' If the signature is a prefix of a signature in the set of signatures, return true
        Dim signatureExtended As String = signatureCleaned & ", "
        Dim signatureLen As Integer = signatureExtended.Length
        For Each sig In signaturesToCompareCleaned
            If sig.Length >= signatureLen AndAlso String.Compare(sig.Substring(0, signatureLen),
                                            signatureExtended, StringComparison.Ordinal) = 0 Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Recursively traverses the tree where <paramref name="node"/> is the root, and updates the 
    ''' <paramref name="signaturePaths"/> dictionary of signature paths for each UIElement node. 
    ''' The resulting dictionary has all the tree's unique signatures as keys, and for each key a 
    ''' list of paths as values. Each path is a list of UIElement nodes representing the path from 
    ''' the root to the node with that signature.
    ''' </summary>
    ''' <remarks>
    ''' Normally I would implement this method as a Function rather than a Sub. 
    ''' But <paramref name="signaturePaths"/> would still be needed as a parameter, so I found 
    ''' using a Sub was more readable.
    ''' </remarks>
    ''' <param name="node"> The root of the (sub)tree to recursively traverse</param>
    ''' <param name="parentNodes"> The list of parent nodes leading to the current node</param>
    ''' <param name="signaturePaths"> The dictionary of signature paths to update</param>
    Private Shared Sub UpdateSignaturePaths(node As UIElement, parentNodes As List(Of UIElement),
                            signaturePaths As Dictionary(Of String, List(Of List(Of UIElement))))

        If node Is Nothing Then Return

        ' Add the current node to the list of parent nodes
        Dim newPath = New List(Of UIElement)(parentNodes) From {node}

        ' Add the list of parent nodes to the dictionary (where the key is this node's signature)
        If Not String.IsNullOrEmpty(node.Signature) Then
            If Not signaturePaths.ContainsKey(node.Signature) Then
                signaturePaths(node.Signature) = New List(Of List(Of UIElement))()
            End If
            signaturePaths(node.Signature).Add(newPath)
        End If

        If node.Children Is Nothing OrElse node.Children.Count = 0 Then
            Return
        End If

        ' Recursively update the dictionary of signature paths for each child
        For Each child In node.Children
            UpdateSignaturePaths(child, newPath, signaturePaths)
        Next
    End Sub

    ''' <summary>
    ''' Writes debug information about the trees in <paramref name="uiElementRoots"/> to a file on 
    ''' the desktop. The debug information includes the structure of each UIElement tree, with 
    ''' indentation for each level, and metadata for each element if available.
    ''' </summary>
    ''' <param name="uiElementRoots"> Dictionary of trees to write debug information for. The key 
    ''' is the title and the value is the root of the tree.</param>
    ''' <param name="metadata"> The metadata associated with the UI elements</param>
    Private Shared Sub WriteDebugFile(uiElementRoots As Dictionary(Of String, UIElement),
                                      metadata As Dictionary(Of String, UIElementMetadata))

        Dim debugInfo As String = ""
        For Each kvp In uiElementRoots
            Dim title = kvp.Key
            Dim element = kvp.Value
            debugInfo += GetDebugInfo(element, 0, title)
        Next

        If uiElementRoots.Count > 0 Then
            Dim lastKey = uiElementRoots.Keys.Last()
            debugInfo += GetDebugInfo(uiElementRoots(lastKey), 0, "After adding metadata:",
                                      metadata)
        End If

        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim debugPath As String = Path.Combine(desktopPath, "tmp", "elementTree.txt")
        If File.Exists(debugPath) Then
            File.AppendAllText(debugPath, debugInfo)
        Else
            Directory.CreateDirectory(Path.GetDirectoryName(debugPath))
            File.WriteAllText(debugPath, debugInfo)
        End If
    End Sub

End Class
