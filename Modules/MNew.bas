Attribute VB_Name = "MNew"
Option Explicit
Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef Dst As Any, ByRef Src As Any, ByVal BytLength As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" ( _
    ByRef Dst As Any, ByVal BytLength As Long)
'die Funktion ArrPtr geht bei allen Arrays außer bei String-Arrays
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" ( _
    ByRef Arr() As Any) As Long

Public Function KeywordsPTreeNode(c As String, ByVal bIsWord As Boolean) As KeywordsPTreeNode
    Set KeywordsPTreeNode = New KeywordsPTreeNode: KeywordsPTreeNode.New_ c, bIsWord
End Function


    
    

'deswegen hier eine Hilfsfunktion für StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As Long
    Call RtlMoveMemory(StrArrPtr, ByVal VarPtr(strArr) + 8, 4)
End Function

'jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
'um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
'anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As Long) As Long
    Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
End Property

Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
    Call RtlMoveMemory(ByVal pArr, RHS, 4)
End Property

Public Sub ZeroSAPtr(ByVal pArr As Long)
    Call RtlZeroMemory(ByVal pArr, 4)
End Sub

