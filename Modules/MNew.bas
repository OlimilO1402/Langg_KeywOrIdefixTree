Attribute VB_Name = "MNew"
Option Explicit

Public Function KeywordsPTreeNode(c As String, ByVal bIsWord As Boolean) As KeywordsPTreeNode
    Set KeywordsPTreeNode = New KeywordsPTreeNode: KeywordsPTreeNode.New_ c, bIsWord
End Function
