Imports System.IO
Imports Newtonsoft.Json

Public Class UISpecBuilder
    ' Prevent instantiation
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Write a UI spec based on the transformations defined in the R model.
    ''' This method reads the transformation definitions from a JSON file, builds a UIElement tree,
    ''' removes duplicates, and writes the final tree structure to a text file on the desktop.
    ''' </summary>
    Public Shared Sub WriteUISpec()

        Dim dlgDefinitionsPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "DialogDefinitions")
        Dim strDialogName As String = "TraitCorrelations"

        ' Load and deserialize the transformation JSON
        Dim strDialogPath As String = Path.Combine(dlgDefinitionsPath, "Dlg" & strDialogName, "dlg" & strDialogName & ".json")
        Dim strTransformationsRJson As String = File.ReadAllText(strDialogPath)

        Dim lstTransformToScript As List(Of clsTransformationRModel) = JsonConvert.DeserializeObject(Of List(Of clsTransformationRModel))(strTransformationsRJson)

        ' Build the initial UIElement tree
        Dim rootElement As New UIElement("Root")
        For Each model In lstTransformToScript
            Dim element = BuildUIElementTree(model)
            If element IsNot Nothing Then
                rootElement.lstChildren.Add(element)
            End If
        Next

        Dim strElementTree As String = vbLf & vbLf & "Before duplicate removal:" & vbLf & vbLf
        strElementTree += OutputUIElementTree(rootElement, 0)

        ' Remove duplicate siblings
        Dim rootElementNoDuplicateSiblings = GetUIElementNoDuplicateSiblings(rootElement)
        strElementTree += vbLf & vbLf & "After duplicate sibling removal:" & vbLf & vbLf
        strElementTree += OutputUIElementTree(rootElementNoDuplicateSiblings, 0)

        ' Remove nodes that are duplicates of their ancestors, or siblings of their ancestors
        Dim rootElementNoDuplicateAncestors = GetUIElementNoDuplicateAncestors(rootElementNoDuplicateSiblings, Nothing, Nothing)
        strElementTree += vbLf & vbLf & "After duplicate ancestor removal:" & vbLf & vbLf
        strElementTree += OutputUIElementTree(rootElementNoDuplicateAncestors, 0)

        Dim rootElementDuplicatesInLca As UIElement = Nothing
        Dim rootElementNoDuplicateAncestorsFinal As UIElement
        Do
            ' If there are duplicates in different branches, then find the largest duplicate tree and
            '   add add it to the LCA (Lowest Common Ancestor)
            rootElementDuplicatesInLca = GetUIElementAddDuplicatesToLca(rootElementNoDuplicateAncestors) ' updated variable name here
            strElementTree += vbLf & vbLf & "After adding longest duplicate to Lowest Common Ancestor (LCA):" & vbLf & vbLf
            strElementTree += OutputUIElementTree(rootElementDuplicatesInLca, 0)

            rootElementNoDuplicateAncestorsFinal = rootElementNoDuplicateAncestors.Clone()

            ' Remove nodes that are duplicates of their ancestors, or siblings of their ancestors
            rootElementNoDuplicateAncestors = GetUIElementNoDuplicateAncestors(rootElementDuplicatesInLca, Nothing, Nothing)
            strElementTree += vbLf & vbLf & "After cleaning LCA tree:" & vbLf & vbLf
            strElementTree += OutputUIElementTree(rootElementNoDuplicateAncestors, 0)

        Loop Until rootElementDuplicatesInLca Is Nothing

        ' Remove duplicate true/false siblings
        Dim rootElementNoDuplicateSiblingsTrueFalse = GetUIElementNoDuplicateSiblings(rootElementNoDuplicateAncestorsFinal, True)
        strElementTree += vbLf & vbLf & "After duplicate true/false sibling removal:" & vbLf & vbLf
        strElementTree += OutputUIElementTree(rootElementNoDuplicateSiblingsTrueFalse, 0)

        ' Remove true/false nodes that are duplicates of their ancestors, or siblings of their ancestors
        Dim rootElementNoDuplicateAncestorsTrueFalse = GetUIElementNoDuplicateAncestors(rootElementNoDuplicateSiblingsTrueFalse, Nothing, Nothing, True)
        strElementTree += vbLf & vbLf & "After true/false duplicate ancestor removal:" & vbLf & vbLf
        strElementTree += OutputUIElementTree(rootElementNoDuplicateAncestorsTrueFalse, 0)

        ' Write the element tree to a file on the desktop
        Dim strDesktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim strFilePath As String = Path.Combine(strDesktopPath, "tmp", "elementTree.txt")
        If File.Exists(strFilePath) Then
            File.AppendAllText(strFilePath, strElementTree)
        Else
            Directory.CreateDirectory(Path.GetDirectoryName(strFilePath))
            File.WriteAllText(strFilePath, strElementTree)
        End If

    End Sub

    Private Shared Function GetUIElementNoDuplicateSiblings(element As UIElement, Optional bIgnoreTrueFalse As Boolean = False) As UIElement
        If element Is Nothing Then Return Nothing

        ' Process children and remove duplicates among siblings
        Dim seenSignatures As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim childrenParse1 As New List(Of UIElement)
        For Each child In element.lstChildren
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing AndAlso Not IsSignatureInSet(cleanedChild.strSignature, seenSignatures, bIgnoreTrueFalse) Then
                seenSignatures.Add(cleanedChild.strSignature)
                childrenParse1.Add(cleanedChild)
            End If
        Next

        'loop backwards through new children to see if we can remove any more siblings
        seenSignatures = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim childrenParse2 As New List(Of UIElement)
        For iChildIndex As Integer = childrenParse1.Count - 1 To 0 Step -1
            Dim child = childrenParse1(iChildIndex)
            Dim cleanedChild = GetUIElementNoDuplicateSiblings(child, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing AndAlso Not IsSignatureInSet(cleanedChild.strSignature, seenSignatures, bIgnoreTrueFalse) Then
                seenSignatures.Add(cleanedChild.strSignature)
                childrenParse2.Add(cleanedChild)
            End If
        Next

        ' Create a new node to avoid mutating the original
        Dim newElement As New UIElement(element.strElementName)
        newElement.lstChildren = childrenParse2
        Return newElement
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
        If IsSignatureInSet(element.strSignature, ancestorsSignatures, bIgnoreTrueFalse) Then
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
        For Each child In element.lstChildren
            siblingSignatures.Add(child.strSignature)
        Next

        Dim newChildren As New List(Of UIElement)
        For i As Integer = 0 To element.lstChildren.Count - 1
            Dim child = element.lstChildren(i)
            Dim cleanedChild = GetUIElementNoDuplicateAncestors(child, ancestorsSignaturesNew, siblingSignatures, bIgnoreTrueFalse)
            If cleanedChild IsNot Nothing Then
                newChildren.Add(cleanedChild)
            End If
        Next

        ' Create a new node to avoid mutating the original
        Dim newElement As New UIElement(element.strElementName)
        newElement.lstChildren = newChildren
        Return newElement
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
        If Not lcaClone.lstChildren.Any(Function(child) child.strSignature = nodeToAdd.strSignature) Then
            lcaClone.lstChildren.Add(nodeToAdd)
        End If

        Return newRoot
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

    Private Shared Function GetSignatureIgnoreTrueFalse(strSignature As String) As String
        If strSignature.EndsWith(" F", StringComparison.Ordinal) Then
            strSignature = strSignature.Substring(0, strSignature.Length - 2)
        End If
        strSignature = strSignature.Replace(" F,", ",")
        Return strSignature
    End Function

    ''' <summary>
    ''' Recursively records all paths to each node by strSignature.
    ''' </summary>
    Private Shared Sub RecordSignaturePaths(node As UIElement, path As List(Of UIElement), signatureToPaths As Dictionary(Of String, List(Of List(Of UIElement))))
        If node Is Nothing Then Return
        Dim newPath = New List(Of UIElement)(path) From {node}
        If Not String.IsNullOrEmpty(node.strSignature) Then
            If Not signatureToPaths.ContainsKey(node.strSignature) Then
                signatureToPaths(node.strSignature) = New List(Of List(Of UIElement))()
            End If
            signatureToPaths(node.strSignature).Add(newPath)
        End If
        For Each child In node.lstChildren
            RecordSignaturePaths(child, newPath, signatureToPaths)
        Next
    End Sub

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
            Dim nextName = path(i).strElementName
            current = current.lstChildren.FirstOrDefault(Function(child) child.strElementName = nextName)
            If current Is Nothing Then Exit For
        Next
        Return current
    End Function

    ''' <summary>
    ''' Clones a subtree (deep copy) of a UIElement node.
    ''' </summary>
    Private Shared Function CloneSubtree(node As UIElement) As UIElement
        If node Is Nothing Then Return Nothing
        Dim newNode As New UIElement(node.strElementName)
        For Each child In node.lstChildren
            Dim newChild = CloneSubtree(child)
            If newChild IsNot Nothing Then
                newNode.lstChildren.Add(newChild)
            End If
        Next
        Return newNode
    End Function

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
                    element.lstChildren.Add(childElement)
                End If
            Next
        End If
        Return element
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

    ' Helper: Clone the tree and build a map from old node to new node
    Private Shared Function CloneTree(node As UIElement, cloneMap As Dictionary(Of UIElement, UIElement)) As UIElement
        If node Is Nothing Then Return Nothing
        Dim newNode As New UIElement(node.strElementName)
        cloneMap(node) = newNode
        For Each child In node.lstChildren
            Dim newChild = CloneTree(child, cloneMap)
            If newChild IsNot Nothing Then
                newNode.lstChildren.Add(newChild)
            End If
        Next
        Return newNode
    End Function

    ''' <summary>
    ''' Returns a string representation of the UIElement tree structure, with indentation for each level.
    ''' </summary>
    Private Shared Function OutputUIElementTree(element As UIElement, level As Integer) As String
        If element Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()
        If Not String.IsNullOrEmpty(element.strElementName) Then
            sb.AppendLine($"{New String(" "c, level * 2)}Level {level}: {element.strElementName}")
        End If
        For Each child In element.lstChildren
            sb.Append(OutputUIElementTree(child, level + 1))
        Next
        Return sb.ToString()
    End Function

End Class
