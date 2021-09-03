VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeywordsPTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_c As String
Public Index As Integer
'Private m_subNodes  As Collection
Private m_subNodes() As KeywordsPTreeNode
Private m_subNdLng() As Long
Public IsKeyWord As Boolean
'
' #,  A -  Z,   a -  z
'35, 65 - 90,  97 - 122
'
Private Sub Class_Initialize()
    'Set m_subNodes = New Collection
    ReDim m_subNodes(0 To 255) As KeywordsPTreeNode
    SAPtr(ArrPtr(m_subNdLng)) = SAPtr(ArrPtr(m_subNodes))
End Sub

Private Sub Class_Terminate()
    ZeroSAPtr ArrPtr(m_subNdLng)
End Sub

Friend Sub New_(c As String, ByVal bIsKeyWord As Boolean)
    m_c = c: IsKeyWord = bIsKeyWord
    Index = CharToIndex(c)
End Sub

Public Property Get Character() As String
    Character = m_c
End Property

Private Function CharToIndex(c As String) As Integer
    CharToIndex = AscW(c) '- 65
End Function

'Public Property Get Index() As Integer
'    Index = m_i
'End Property
'Public Property Get IsKeyword() As Boolean
'    IsKeyword = m_isKeyword
'End Property

Public Function SubNodes_Add(aNode As KeywordsPTreeNode) As KeywordsPTreeNode
'    m_subNodes.Add aNode, aNode.Character
    Set SubNodes_Add = aNode
    Set m_subNodes(aNode.Index) = SubNodes_Add
End Function

Public Function SubNodes_Contains(c As String) As Boolean
    'On Error Resume Next
    'SubNodes_Contains = Not IsEmpty(m_subNodes(c))
    Dim i As Integer: i = AscW(c) 'CharToIndex(c)
    If 0 < i And i < 255 Then
        'SubNodes_Contains = Not (m_subNodes(i) Is Nothing)
        SubNodes_Contains = m_subNdLng(i) <> 0
    End If
End Function

'Public Property Get SubNodes() As Collection
'    Set SubNodes = m_subNodes
'End Property
Public Property Get SubNodes_Item(c As String) As KeywordsPTreeNode
    'Dim i As Integer: i = CharToIndex(c)
    'Set SubNodes_Item = m_subNodes(i)
    Set SubNodes_Item = m_subNodes(Asc(c))
End Property

Friend Function ToStr(ByVal indent As String, ByRef i_inout As Long) As String
    Dim s As String
    If Len(m_c) Then
        i_inout = i_inout + 1
        s = Format(i_inout, "000") & ": " & indent & m_c & IIf(Me.IsKeyWord, " -> keyword", "") & vbCrLf
        indent = indent & " "
    End If
    Dim v
    Dim node As KeywordsPTreeNode
    For Each v In m_subNodes
        Set node = v
        If Not (node Is Nothing) Then
            s = s & node.ToStr(indent, i_inout)
        End If
    Next
    ToStr = s
End Function
