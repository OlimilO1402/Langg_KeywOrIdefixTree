VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrefixTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Olaf Schmidt in January 2018

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements1D As Long
    lLbound1D As Long
End Type

Private Declare Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc&, Optional ByVal CB& = 4)

Private WChars() As Integer
Private saWChars As SAFEARRAY1D

Private Type tNode 'this type-def will only take up 136 Bytes per Node
    KeyIndex As Long
    ChildIndexes(3 To 45 * 3 - 1) As Byte '(the ASCII-26Char-Alphabet + 18 ExtraChars (as e.g. [0-9], [#], [_], etc.)
End Type

Private Nodes() As tNode
Private NodesUB As Long
Private NodesCount As Long

Private mKeys() As String
Private mKeysUB As Long
Private mKeysCount As Long

Private AlphabetIdxLUT(-32768 To 32767) As Byte

Private Sub Class_Initialize()
    'ensure a permanent (static) virtually bound "Loop-Container for Strings" (as an Int16-WChar-Array)...
    saWChars.cDims = 1
    saWChars.fFeatures = &H11 '...by using FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10 ...
    saWChars.cbElements = 2
    saWChars.cLocks = 1 '...and also a Lock (to prevent the runtime from destroying the virtual array when the instance terminates)
    saWChars.cElements1D = 2147483647
    saWChars.lLbound1D = -2
    BindArray WChars, VarPtr(saWChars) 'so we only need to bind this once in Class_Initialize (no need for us to destroy the Binding in Class_Terminate)
    
    'now we pre-allocate a reasonable amount of:
    NodesUB = 512: ReDim Preserve Nodes(0 To NodesUB) 'Nodes (512 * 136) = only about 68kByte
    mKeysUB = 512: ReDim Preserve mKeys(0 To mKeysUB) 'and mKeys
    
    'and finally init the 64kByte large Base-Alphabet-Mapping-LUT already (for case-insensitive comparisons)
    Dim i As Long
    For i = 65 To 90
        AlphabetIdxLUT(i) = 3 * (i - 64) 'will store 'A' as index 1, 'B' as index 2, etc... up to 'Z'
        AlphabetIdxLUT(i + 32) = AlphabetIdxLUT(i) 'put the same map-value also into the [a-z]-range
    Next
End Sub

Public Sub InitExtraChars(ExtraChars As String) 'call this with max 14 Extra-Chars, when you need this in the allowed AlphabetChars
    'the mapping-indexes for the normal (case-insensitive ASCII-alphabet) are [1-26],
    'so we have room for max. 44-26 = 18 extra-chars (with the current tNode-Types ChildIndexes-Array-Ubound)
    If Len(ExtraChars) > (UBound(Nodes(0).ChildIndexes) \ 3) - 26 Then Err.Raise vbObjectError, , "too many ExtraChars"
    Dim Idx As Long: Idx = 27
    Dim i As Long
    For i = 1 To Len(ExtraChars)
        AlphabetIdxLUT(AscW(Mid$(ExtraChars, i, 1))) = 3 * Idx: Idx = Idx + 1
    Next
End Sub

Public Function Add(Key As String) As Long 'returns 0, when the Key was so far unknown in the Tree, otherwise the one-based Index under which it was stored
    saWChars.pvData = StrPtr(Key) - 4 'set pvData to the real allocation-pointer of the BSTR (the Int16-Array was defined with an LBound of -2)
    If saWChars.pvData = -4 Then saWChars.pvData = 0: Exit Function 'avoid looping on passed NullPointer-BStrings
    If WChars(-2) = 0 Then Exit Function 'also avoid any further actions, when an empty String was passed (WChars(-2) accesses the BSTR-Len-descriptor)
    
    Dim i As Long, AlphIdx As Byte, NodeIdx As Long
    For i = 0 To WChars(-2) \ 2 - 1
      'check and prolong the NodesArray-Buffer by 50%, if necessary
      If NodesCount >= NodesUB Then NodesUB = NodesUB * 1.5: ReDim Preserve Nodes(0 To NodesUB)
    
      AlphIdx = AlphabetIdxLUT(WChars(i))
      If AlphIdx = 0 Then Err.Raise vbObjectError, , "the Char [" & ChrW(WChars(i)) & "] is not part of the Alphabet"
      
      With Nodes(NodeIdx) 'the With-Construct is holding the "Parent-Node" here
        'so that we are free to retrieve the next Child-Index (stored in 3 Bytes as 24Bit-Index)
        NodeIdx = .ChildIndexes(AlphIdx) + 256& * .ChildIndexes(AlphIdx + 1) + 65536 * .ChildIndexes(AlphIdx + 2)
        If NodeIdx = 0 Then 'this WCharNode does not exist yet, so we add a new one
           NodesCount = NodesCount + 1 'by incrementing the Nodes-Counter
           NodeIdx = NodesCount 'and setting the current NodeIdx to the new Node-Entry-Slot
           'Nodes(NodeIdx).NodeChar = WChars(i) 'now we Update the current Node-Char at the new Node
           'finally we write the new NodeIdx as a 24Bit-Value into the Parent-Nodes ChildIndexes-Array
           .ChildIndexes(AlphIdx + 0) = (NodeIdx) And &HFF&
           .ChildIndexes(AlphIdx + 1) = (NodeIdx \ 256) And &HFF&
           .ChildIndexes(AlphIdx + 2) = (NodeIdx \ 65536) And &HFF&
        End If
      End With
    Next
    
    Add = Nodes(NodeIdx).KeyIndex 'try to return the End-Marker(Counter)
    
    If Add = 0 Then 'if it's not there yet, then flag the new End-Node properly
       
       mKeys(mKeysCount) = Key 'store the original Key in an Array-allocation which is separate from the Nodes-Arr
       'check and prolong the mKeysArray-Buffer by 100%, if necessary
       If mKeysCount >= mKeysUB Then mKeysUB = mKeysUB * 2: ReDim Preserve mKeys(0 To mKeysUB)
       
       mKeysCount = mKeysCount + 1
       Nodes(NodeIdx).KeyIndex = mKeysCount
    End If
End Function

Public Function Exists(Key As String) As Long 'returns 0, when the Key is unknown in the Tree, otherwise the one-based Index under which it was stored
    
    saWChars.pvData = StrPtr(Key) - 4 'set pvData to the real allocation-pointer of the BSTR (the Int16-Array was defined with an LBound of -2)
    If saWChars.pvData = -4 Then saWChars.pvData = 0: Exit Function 'avoid looping on passed NullPointer-BStrings
    If WChars(-2) = 0 Then Exit Function 'also avoid any further actions, when an empty String was passed (WChars(-2) accesses the BSTR-Len-descriptor)
    
    Dim i As Long
    Dim AlphIdx As Byte
    Dim NodeIdx As Long
    
    For i = 0 To WChars(-2) \ 2 - 1
    
        AlphIdx = AlphabetIdxLUT(WChars(i))
        
        If AlphIdx = 0 Then Exit Function 'exit with 0, because the Char is not part of the Alphabet"
        
        With Nodes(NodeIdx)
        
            NodeIdx = .ChildIndexes(AlphIdx) + 256& * .ChildIndexes(AlphIdx + 1) + 65536 * .ChildIndexes(AlphIdx + 2)
            
        End With
        
        If NodeIdx = 0 Then Exit Function 'exit with 0, because this WCharNode does not exist
        
    Next
    
    Exists = Nodes(NodeIdx).KeyIndex
    
End Function

Public Property Get CountOfNodes() As Long
    CountOfNodes = NodesCount
End Property
Public Property Get CountOfKeys() As Long
    CountOfKeys = mKeysCount
End Property

Public Function Keys()
    If mKeysCount Then mKeysUB = mKeysCount + 512 Else Keys = Split(""): Exit Function
    ReDim Preserve mKeys(0 To mKeysCount - 1) 'ensure the correct amount of Keys in the Array
    Keys = mKeys 'place a  copy of the resized array in the return-value
    ReDim Preserve mKeys(0 To mKeysUB) 'now ensure, that the add-logic will work properly again
End Function
