VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordsPTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_root As KeywordsPTreeNode

Private Sub Class_Initialize()
    Set m_root = New KeywordsPTreeNode
End Sub

Public Sub Add(aWord As String)
    AddRecursive m_root, aWord
End Sub

Private Sub AddRecursive(aNode As KeywordsPTreeNode, aStr As String)
    
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
    
    Dim nextNode As KeywordsPTreeNode
    
    If aNode.SubNodes_Contains(c) Then
        
        Set nextNode = aNode.SubNodes_Item(c)
        
    Else
        
        Set nextNode = aNode.SubNodes_Add(MNew.KeywordsPTreeNode(c, bIsKeyWord))
        
    End If
    
    If Not bIsKeyWord Then
        
        AddRecursive nextNode, sRest
        
    End If
    
End Sub

Public Function Contains(aStr As String) As Boolean
    Dim i As Long, l As Long: l = LenB(aStr)
    Dim nextNode As KeywordsPTreeNode: Set nextNode = m_root
    Dim c As String
    For i = 1 To l
        'entweder Mid() oder Mid$() oder MidB() oder MidB$()
        'c = MidB$(aStr, i, 1)
        c = Mid$(aStr, i, 1)
        If Len(c) > 0 Then
            Contains = nextNode.SubNodes_Contains(c)
            If Not Contains Then Exit Function
            Set nextNode = nextNode.SubNodes_Item(c)
        End If
    Next
    Contains = nextNode.IsKeyword
End Function

Friend Function ToStr() As String
    Dim indent As String, i As Long
    ToStr = m_root.ToStr(indent, i)
End Function

