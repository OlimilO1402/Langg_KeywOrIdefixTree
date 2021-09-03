Attribute VB_Name = "MKwPTree"
Option Explicit

Private Type TKwPTreeNode
    Character   As String             ' 4 actually not necessary
    CharIndex   As Integer            ' 2 the Index to this element in SubNodesI
    TreeIndex   As Integer            ' 2 the Index to this element in KwPTree.PTree
    IsKeyWord   As Boolean            ' 2 decides about twig or leaf
    SubNodesI(0 To 255) As Integer    ' 256*2 Indices to the next elements in the array KwPTree.PTree()
End Type

Private Type TKwPTree
    Count   As Long
    PTree() As TKwPTreeNode
End Type

Private KwPTree As TKwPTree 'the root is the first element with index 0 in the array KwPTree.PTree()
Private m_root  As Integer 'TKwPTreeNode ' = 0

' ############################## '   Type TKwPTreeNode  ' ############################## '
Private Function New_TKwPTreeNode(c As String, ByVal bIsWord As Boolean) As TKwPTreeNode
    With New_TKwPTreeNode
        .Character = c
        If Len(c) Then .CharIndex = AscW(c)
        .IsKeyWord = bIsWord
    End With
End Function

Private Function TKwPTreeNode_SubNodes_Add(this As TKwPTreeNode, aNodeTreeIndex As Integer) As Integer 'TKwPTreeNode
    With this
        .SubNodesI(KwPTree.PTree(aNodeTreeIndex).CharIndex) = aNodeTreeIndex 'aNode.TreeIndex
        TKwPTreeNode_SubNodes_Add = aNodeTreeIndex 'aNode.TreeIndex
    End With
End Function

Private Function TKwPTreeNode_SubNodes_Contains(this As TKwPTreeNode, c As String) As Boolean
    With this
        Dim i As Integer: i = AscW(c)
        TKwPTreeNode_SubNodes_Contains = .SubNodesI(i) > 0
    End With
End Function

Private Function TKwPTreeNode_SubNodes_Item(this As TKwPTreeNode, c As String) As Integer 'TKwPTreeNode
    With this
        Dim i As Integer: i = AscW(c)
        TKwPTreeNode_SubNodes_Item = .SubNodesI(i)
    End With
End Function

Private Function TKwPTreeNode_ToStr(this As TKwPTreeNode, ByVal indent As String, ByRef i_inout As Long) As String
    Dim s As String
    With this
        If Len(.Character) Then
            i_inout = i_inout + 1
            s = Format(i_inout, "000") & ": " & indent & .Character & IIf(.IsKeyWord, " -> keyword", "") & vbCrLf
            indent = indent & " "
        End If
        Dim i As Long
        For i = 0 To 255
            If .SubNodesI(i) > 0 Then
                s = s & TKwPTreeNode_ToStr(KwPTree.PTree(.SubNodesI(i)), indent, i_inout)
            End If
        Next
    End With
    TKwPTreeNode_ToStr = s
End Function

' ############################## '     Data KwPTree    ' ############################## '
Private Function KwPTree_AddNode(NewNode As TKwPTreeNode) As Integer 'TKwPTreeNode
    With KwPTree
        If .Count = 0 Then
            ReDim .PTree(0 To 1000) 'As TKwPTreeNode
        Else
            Dim u As Long: u = UBound(.PTree)
            'Oh shit, so geht das nicht, da Array gesperrt wenn ein Element darin gerade verwendet wird
            'also OK wir müssen das gesamte Arry vorher groß genug anlegen
            If u < .Count Then ReDim Preserve .PTree(0 To 2 * u - 1)
        End If
        .PTree(.Count) = NewNode
        .PTree(.Count).TreeIndex = .Count
        KwPTree_AddNode = .Count
        .Count = .Count + 1
    End With
End Function

' ############################## '     Type TKwPTree    ' ############################## '
Private Sub KwPTree_Add(aWord As String)
    With KwPTree
        If .Count = 0 Then
            m_root = KwPTree_AddNode(New_TKwPTreeNode("", False))
        End If
        KwPTree_AddRecursive KwPTree.PTree(m_root), aWord
    End With
End Sub

Private Sub KwPTree_AddRecursive(aNode As TKwPTreeNode, aStr As String)
    
    Dim l As Long: l = Len(aStr)
    
    If l = 0 Then Exit Sub
    
    Dim bIsKeyWord As Boolean
    Dim c As String, sRest As String
    
    If l = 1 Then
        c = aStr
        bIsKeyWord = True
    Else
        c = Mid$(aStr, 1, 1)
        sRest = Mid$(aStr, 2)
    End If
    
    Dim nextNode As Integer ' TKwPTreeNode
    
    If TKwPTreeNode_SubNodes_Contains(aNode, c) Then
        
        nextNode = TKwPTreeNode_SubNodes_Item(aNode, c)
        
    Else
        
        nextNode = TKwPTreeNode_SubNodes_Add(aNode, KwPTree_AddNode(New_TKwPTreeNode(c, bIsKeyWord)))
        
    End If
    
    If Not bIsKeyWord Then
        
        KwPTree_AddRecursive KwPTree.PTree(nextNode), sRest
        
    End If
    
End Sub

Private Function KwPTree_Contains(aStr As String) As Boolean
    Dim i As Long, l As Long: l = LenB(aStr)
    Dim nextNode As Integer: nextNode = m_root
    Dim c As String
    For i = 1 To l
        'entweder Mid() oder Mid$() oder MidB() oder MidB$()
        'c = MidB$(aStr, i, 1)
        c = Mid$(aStr, i, 1)
        If Len(c) > 0 Then
            KwPTree_Contains = TKwPTreeNode_SubNodes_Contains(KwPTree.PTree(nextNode), c)
            If Not KwPTree_Contains Then Exit Function
            nextNode = TKwPTreeNode_SubNodes_Item(KwPTree.PTree(nextNode), c)
        End If
    Next
    KwPTree_Contains = KwPTree.PTree(nextNode).IsKeyWord
End Function

Public Function KwPTree_ToStr() As String
    Dim indent As String, i As Long
    KwPTree_ToStr = TKwPTreeNode_ToStr(KwPTree.PTree(m_root), indent, i)
End Function

' ############################## '     Module MKwPTree    ' ############################## '
Public Sub KeyWords_Fill()
    If Len(MKeywords.VBKeywordsUCase) = 0 Then MKeywords.VBKeywordsUCase = UCase(MKeywords.VBKeywords)
    Dim sa() As String: sa = Split(MKeywords.VBKeywordsUCase, " ")
    Dim i As Long
    For i = LBound(sa) To UBound(sa)
        KwPTree_Add sa(i)
    Next
End Sub

Public Function IsVBKeyword_TPTree(aWord As String) As Boolean
    IsVBKeyword_TPTree = KwPTree_Contains(aWord)
End Function
