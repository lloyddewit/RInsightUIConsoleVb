Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

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
    ''' <param name="dialogName"></param>
    Public Shared Sub WriteUISpec(dialogName As String)

        ' Build the initial UIElement tree
        Dim root As New UIElement("Root")
        Dim transformations As List(Of clsTransformationRModel) = ReadTransformations(dialogName)
        For Each model In transformations
            Dim element = BuildUIElementTree(model)
            If element IsNot Nothing Then
                root.Children.Add(element)
            End If
        Next

        ' Remove duplicate siblings
        Dim rootNoDuplicateSiblings = GetUIElementNoDuplicateSiblings(root)

        ' Remove nodes that are duplicates of their ancestors of their ancestors
        Dim rootNoDuplicateAncestors = GetUIElementNoDuplicateAncestors(
            rootNoDuplicateSiblings, Nothing, Nothing)

        Dim rootDuplicatesInLca As UIElement = Nothing
        Dim rootNoDuplicateAncestorsFinal As UIElement
        Do
            ' If there are duplicates in different branches, then find the largest duplicate tree
            '   and add it to the LCA (Lowest Common Ancestor)
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
            ReadAndValidateMetadata(dialogName)

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
    End Sub

    ''' <summary>
    ''' Recursively builds a UIElement tree from a clsTransformationRModel tree.
    ''' </summary>
    Private Shared Function BuildUIElementTree(model As clsTransformationRModel) As UIElement
        If model Is Nothing Then Return Nothing

        ' Only create a UIElement if strValueKey is not null or empty
        If String.IsNullOrEmpty(model.strValueKey) AndAlso (model.lstTransformations Is Nothing OrElse model.lstTransformations.Count = 0) Then
            Return Nothing
        End If

        Dim strElementName As String = model.strValueKey
        strElementName += If(model.enumTransformationType = clsTransformationRModel.TransformationType.ifFalseExecuteChildTransformations, " F", "")
        Dim element As New UIElement(strElementName)
        If model.lstTransformations IsNot Nothing Then
            For Each child In model.lstTransformations
                Dim childElement = BuildUIElementTree(child)
                If childElement IsNot Nothing Then
                    element.Children.Add(childElement)
                End If
            Next
        End If
        Return element
    End Function

    ''' <summary>
    ''' Clones a subtree (deep copy) of a UIElement node.
    ''' </summary>
    Private Shared Function CloneSubtree(node As UIElement) As UIElement
        If node Is Nothing Then Return Nothing
        Dim newNode As New UIElement(node.ElementName)
        For Each child In node.Children
            Dim newChild = CloneSubtree(child)
            If newChild IsNot Nothing Then
                newNode.Children.Add(newChild)
            End If
        Next
        Return newNode
    End Function

    ' Helper: Clone the tree and build a map from old node to new node
    Private Shared Function CloneTree(node As UIElement, cloneMap As Dictionary(Of UIElement, UIElement)) As UIElement
        If node Is Nothing Then Return Nothing
        Dim newNode As New UIElement(node.ElementName)
        cloneMap(node) = newNode
        For Each child In node.Children
            Dim newChild = CloneTree(child, cloneMap)
            If newChild IsNot Nothing Then
                newNode.Children.Add(newChild)
            End If
        Next
        Return newNode
    End Function

    ' Helper: Find the lowest common ancestor from a list of paths
    Private Shared Function FindLCAFromPaths(paths As List(Of List(Of UIElement))) As UIElement
        If paths Is Nothing OrElse paths.Count = 0 Then Return Nothing
        Dim minLen = paths.Min(Function(p) p.Count)
        Dim lca As UIElement = Nothing
        For i = 0 To minLen - 1
            Dim thisNode = paths(0)(i)
            Dim iIndex = i 'needed to prevent warning in line below
            If paths.All(Function(p) p(iIndex) Is thisNode) Then
                lca = thisNode
            Else
                Exit For
            End If
        Next
        Return lca
    End Function

    ''' <summary>
    ''' Finds the LCA node in the clone tree, given a path and the original LCA node.
    ''' </summary>
    Private Shared Function FindNodeByPath(root As UIElement, path As List(Of UIElement), cloneMap As Dictionary(Of UIElement, UIElement), Optional upToLca As UIElement = Nothing) As UIElement
        If path Is Nothing OrElse path.Count = 0 Then Return Nothing
        Dim current As UIElement = root
        For i As Integer = 1 To path.Count - 1 ' skip root (already at root)
            If upToLca IsNot Nothing AndAlso path(i - 1) Is upToLca Then
                Exit For
            End If
            Dim nextName = path(i).ElementName
            current = current.Children.FirstOrDefault(Function(child) child.ElementName = nextName)
            If current Is Nothing Then Exit For
        Next
        Return current
    End Function

    ''' <summary>
    ''' Returns a string representation of the UIElement tree structure, with indentation for each level.
    ''' </summary>
    Private Shared Function GetDebugInfo(element As UIElement, level As Integer, title As String, Optional metadataDict As Dictionary(Of String, UIElementMetadata) = Nothing) As String
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
                    Dim numMetadata As UIElementMetadataNumber = CType(metadata, UIElementMetadataNumber)
                    metadataString &= $", Min: {numMetadata.Min}, Max: {numMetadata.Max}, Increment: {numMetadata.Increment}, IsInteger: {numMetadata.IsInteger}"
                ElseIf TypeOf metadata Is UIElementMetadataEnumeration Then
                    Dim enumMetadata As UIElementMetadataEnumeration = CType(metadata, UIElementMetadataEnumeration)
                    metadataString &= $", Options: {String.Join(", ", enumMetadata.Options)}, OptionType: {enumMetadata.OptionType}"
                End If
            End If

            sb.AppendLine($"{New String(" "c, level * 2)}Level {level}: {element.ElementName}, {metadataString}")
        End If

        For Each child In element.Children
            sb.Append(GetDebugInfo(child, level + 1, Nothing, metadataDict))
        Next
        Return sb.ToString()
    End Function

    Private Shared Function GetSignatureIgnoreTrueFalse(strSignature As String) As String
        If strSignature.EndsWith(" F", StringComparison.Ordinal) Then
            strSignature = strSignature.Substring(0, strSignature.Length - 2)
        End If
        strSignature = strSignature.Replace(" F,", ",")
        Return strSignature
    End Function

    ''' <summary>
    ''' Returns a new UIElement tree with one extra node added: if there are duplicate signatures in the tree,
    ''' finds the duplicated node with the longest signature, finds any two nodes with this signature, finds their LCA,
    ''' and adds a duplicate of the node as a child of the LCA. If no duplicates exist, returns Nothing.
    ''' </summary>
    Private Shared Function GetUIElementAddDuplicatesToLca(element As UIElement) As UIElement
        If element Is Nothing Then Return Nothing

        ' Step 1: Traverse and record all paths to each strSignature
        Dim signatureToPaths As New Dictionary(Of String, List(Of List(Of UIElement)))(StringComparer.OrdinalIgnoreCase)
        RecordSignaturePaths(element, New List(Of UIElement), signatureToPaths)

        ' Step 2: Find the duplicated signature with the longest signature string
        Dim duplicatedSignatures = signatureToPaths.Where(Function(kvp) kvp.Value.Count > 1).ToList()
        If duplicatedSignatures.Count = 0 Then
            Return Nothing
        End If

        Dim longestDup = duplicatedSignatures.OrderByDescending(Function(kvp) kvp.Key.Length).First()
        Dim signature As String = longestDup.Key
        Dim paths As List(Of List(Of UIElement)) = longestDup.Value

        ' Step 3: Pick any two nodes with this signature and find their LCA
        Dim path1 = paths(0)
        Dim path2 = paths(1)
        Dim nodeToDuplicate = path2.Last() ' Use the second node as the one to duplicate
        Dim lca = FindLCAFromPaths(New List(Of List(Of UIElement)) From {path1, path2})
        If lca Is Nothing Then Return Nothing

        ' Step 4: Clone the tree so we can add nodes without mutating the original
        Dim cloneMap As New Dictionary(Of UIElement, UIElement)
        Dim newRoot As UIElement = CloneTree(element, cloneMap)

        ' Step 5: Find the LCA node in the cloned tree
        Dim lcaClone = FindNodeByPath(newRoot, path1, cloneMap, upToLca:=lca)
        If lcaClone Is Nothing Then Return newRoot

        ' Step 6: Add a clone of the duplicate node to the LCA's children (if not already present)
        Dim nodeToAdd = CloneSubtree(cloneMap(nodeToDuplicate))
        If Not lcaClone.Children.Any(Function(child) child.Signature = nodeToAdd.Signature) Then
            lcaClone.Children.Add(nodeToAdd)
        End If

        Return newRoot
    End Function

    Private Shared Function GetUIElementNoDuplicateAncestors(element As UIElement,
                                                      ancestorsSignatures As HashSet(Of String),
                                                      parentSiblings As HashSet(Of String),
                                                      Optional bIgnoreTrueFalse As Boolean = False) As UIElement
        If element Is Nothing Then Return Nothing

        If ancestorsSignatures Is Nothing Then
            ancestorsSignatures = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        End If

        ' If this strSignature matches any ancestor, remove this node
        If IsSignatureInSet(element.Signature, ancestorsSignatures, bIgnoreTrueFalse) Then
            Return Nothing
        End If

        ' Make a list of ancestors for this nodes children, include the siblings of this parent
        Dim ancestorsSignaturesNew As New HashSet(Of String)(ancestorsSignatures, StringComparer.OrdinalIgnoreCase)
        If parentSiblings IsNot Nothing Then
            For Each siblingSignature In parentSiblings
                ancestorsSignaturesNew.Add(siblingSignature)
            Next
        End If

        ' Make a list of parent siblings for this nodes children
        Dim siblingSignatures As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each child In element.Children
            siblingSignatures.Add(child.Signature)
        Next

        Dim newChildren As New List(Of UIElement)
        For i As Integer = 0 To element.Children.Count - 1
            Dim child = element.Children(i)
            Dim cleanedChild = GetUIElementNoDuplicateAncestors(child, ancestorsSignaturesNew, siblingSignatures, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing Then
                newChildren.Add(cleanedChild)
            End If
        Next

        ' Create a new node to avoid mutating the original
        Dim newElement As New UIElement(element.ElementName)
        newElement.Children = newChildren
        Return newElement
    End Function

    Private Shared Function GetUIElementNoDuplicateSiblings(element As UIElement, Optional bIgnoreTrueFalse As Boolean = False) As UIElement
        If element Is Nothing Then Return Nothing

        ' Process children and remove duplicates among siblings
        Dim seenSignatures As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim childrenParse1 As New List(Of UIElement)
        For Each child In element.Children
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing AndAlso Not IsSignatureInSet(cleanedChild.Signature, seenSignatures, bIgnoreTrueFalse) Then
                seenSignatures.Add(cleanedChild.Signature)
                childrenParse1.Add(cleanedChild)
            End If
        Next

        'loop backwards through new children to see if we can remove any more siblings
        seenSignatures = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim childrenParse2 As New List(Of UIElement)
        For iChildIndex As Integer = childrenParse1.Count - 1 To 0 Step -1
            Dim child = childrenParse1(iChildIndex)
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing AndAlso Not IsSignatureInSet(cleanedChild.Signature, seenSignatures, bIgnoreTrueFalse) Then
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
    ''' Checks if strSignature is in signaturesToCompare, or if any string in signaturesToCompare starts with strSignature plus ", ".
    ''' </summary>
    Private Shared Function IsSignatureInSet(strSignature As String,
                                      signaturesToCompare As HashSet(Of String),
                                      bIgnoreTrueFalse As Boolean) As Boolean

        Dim strSignatureCleaned As String = If(bIgnoreTrueFalse, GetSignatureIgnoreTrueFalse(strSignature), strSignature)

        Dim signaturesToCompareCleaned = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each sig In signaturesToCompare
            Dim sigCleaned As String = If(bIgnoreTrueFalse, GetSignatureIgnoreTrueFalse(sig), sig)
            signaturesToCompareCleaned.Add(sigCleaned)
        Next

        If signaturesToCompareCleaned.Contains(strSignatureCleaned) Then
            Return True
        End If

        Dim strSignatureExtended As String = strSignatureCleaned & ", "
        Dim iSignatureLen As Integer = strSignatureExtended.Length
        For Each sig In signaturesToCompareCleaned
            If sig.Length >= iSignatureLen AndAlso String.Compare(sig.Substring(0, iSignatureLen), strSignatureExtended, StringComparison.Ordinal) = 0 Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function ReadAndValidateMetadata(dialogName As String) As Dictionary(Of String, UIElementMetadata)
        Dim dialogDefinitionsPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "DialogDefinitions")
        Dim metadataPath As String = Path.Combine(dialogDefinitionsPath, "Dlg" & dialogName, "dlg" & dialogName & "Metadata.json")

        Dim settings As New JsonSerializerSettings With {
            .TypeNameHandling = TypeNameHandling.Auto,
            .Formatting = Formatting.Indented}
        settings.Converters.Add(New StringEnumConverter())

        Dim metadata As Dictionary(Of String, UIElementMetadata) = JsonConvert.DeserializeObject(Of Dictionary(Of String, UIElementMetadata))(File.ReadAllText(metadataPath), settings)

        For Each kvp In metadata
            Dim key = kvp.Key
            Dim meta = kvp.Value
            If Not meta.IsResetDefaultValid() Then
                Console.WriteLine($"Warning: ResetDefault value '{meta.ResetDefault}' is invalid for '{key}' ({meta.GetType().Name})")
            End If
        Next

        Return metadata
    End Function

    Private Shared Function ReadTransformations(dialogName As String) As List(Of clsTransformationRModel)

        Dim dialogDefinitionsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "DialogDefinitions")
        Dim dialogPath As String = Path.Combine(
            dialogDefinitionsPath, "Dlg" & dialogName, "dlg" & dialogName & ".json")

        Dim transformationsJson As String = File.ReadAllText(dialogPath)
        Return JsonConvert.DeserializeObject(Of List(Of clsTransformationRModel))(transformationsJson)

    End Function

    ''' <summary>
    ''' Recursively records all paths to each node by strSignature.
    ''' </summary>
    Private Shared Sub RecordSignaturePaths(node As UIElement, path As List(Of UIElement), signatureToPaths As Dictionary(Of String, List(Of List(Of UIElement))))
        If node Is Nothing Then Return
        Dim newPath = New List(Of UIElement)(path) From {node}
        If Not String.IsNullOrEmpty(node.Signature) Then
            If Not signatureToPaths.ContainsKey(node.Signature) Then
                signatureToPaths(node.Signature) = New List(Of List(Of UIElement))()
            End If
            signatureToPaths(node.Signature).Add(newPath)
        End If
        For Each child In node.Children
            RecordSignaturePaths(child, newPath, signatureToPaths)
        Next
    End Sub

    Private Shared Sub WriteDebugFile(uiElements As Dictionary(Of String, UIElement), metadata As Dictionary(Of String, UIElementMetadata))

        Dim debugInfo As String = ""
        For Each kvp In uiElements
            Dim title = kvp.Key
            Dim element = kvp.Value
            debugInfo += GetDebugInfo(element, 0, title)
        Next

        If uiElements.Count > 0 Then
            Dim lastKey = uiElements.Keys.Last()
            debugInfo += GetDebugInfo(uiElements(lastKey), 0, "After adding metadata:", metadata)
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
