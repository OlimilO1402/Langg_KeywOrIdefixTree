Attribute VB_Name = "MKeywords"
Option Explicit
'https://docs.microsoft.com/de-de/dotnet/visual-basic/language-reference/keywords/

Public Const VBReservedKeywords As String = _
"AddHandler AddressOf Alias And AndAlso As " & _
"Boolean ByRef Byte ByVal " & _
"Call Case Catch CBool CByte CChar CDate CDbl CDec Char CInt Class CLng CObj Const Continue CSByte CShort CSng CStr CType CUInt CULng CUShort " & _
"Date Decimal Declare Default Delegate Dim DirectCast Do Double " & _
"Each Else ElseIf End EndIf Enum Erase Error Event Exit " & _
"False Finally For Friend Function " & _
"Get GetType GetXMLNamespace Global GoSub GoTo " & _
"Handles " & _
"If Implements Imports In Inherits Integer Interface Is IsNot " & _
"Let Lib Like Long Loop " & _
"Me Mod Modul Module MustInherit MustOverride MyBase MyClass " & _
"NameOf Namespace Narrowing New " & _
"Next Not Nothing NotInheritable NotOverridable " & _
"Objekt Of On Operator Option Optional Or OrElse out Overloads Overrides " & _
"ParamArray Partial Private Property Protected Public " & _
"RaiseEvent ReadOnly ReDim REM RemoveHandler Resume Return " & _
"SByte Select Set Shadows Shared Short Single Static Step Stop String Structure Sub SyncLock " & _
"Then Throw True Try TryCast TypeOf " & _
"UInteger ULong UShort Using " & _
"Variant " & _
"Wend When While Widening With WithEvents WriteOnly " & _
"XOr " & _
"#Else #ElseIf #End #If "



Public Const VBNonReservedKeywords As String = _
"Aggregat ANSI Assembly Async Auto Await " & _
"Binary " & _
"Compare " & _
"Distinct " & _
"Equals Explicit " & _
"From " & _
"GroupBy " & _
"Into IsFalse IsTrue Iterator " & _
"Join " & _
"Key " & _
"Mid " & _
"Off OrderBy " & _
"Preserve " & _
"Skip SkipWhile Strict " & _
"Take TakeWhile Text " & _
"Unicode " & _
"Until " & _
"Where " & _
"Yield " & _
"#ExternalSource " & _
"#Region "

'"Benutzerdefiniert " & _

Public Const VBOperators As String = "= & &= * *= / /= \ \= ^ ^= + += - -= >> >>= << <<= "

Public Const VBKeywords = VBReservedKeywords & VBNonReservedKeywords ' & VBOperators

Public VBKeywordsUCase As String
'"AddHandler AddressOf Alias And AndAlso As " & _
'"Boolean ByRef Byte ByVal" & _
'"Call Case Catch CBool CByte CChar CDate CDbl CDec Char CInt Class-Einschr‰nkung Class-Anweisung CLng CObj Const Continue CSByte CShort CSng CStr CType CUInt CULng CUShort" & _
'"Date Decimal Declare Default Delegate Dim DirectCast Do Double" & _
'"Each Else ElseIf End-Anweisung End <Schl¸sselwort> EndIf Enum Erase Error (On Error) Event Exit" & _
'"False Finally For (in ForÖNext) For EachÖNext Friend Function" & _
'"Get GetType GetXMLNamespace Global GoSub GoTo" & _
'"Handles" & _
'"If If() Implements Implements-Anweisung Imports (.NET-Namespace und Typ) Imports (XML-Namespace) In In (generischer Modifizierer) Inherits Integer Interface Is IsNot" & _
'"Let Lib Like Long Loop" & _
'"Me Mod Modul Module-Anweisung MustInherit MustOverride MyBase MyClass" & _
'"NameOf Namespace Narrowing new-Einschr‰nkung New-Operator" & _
'"Next Next (in Resume) Not Nothing NotInheritable NotOverridable" & _
'"Objekt Of On Operator Option Optional Or OrElse out (generischer Modifizierer) Overloads Overrides Overrides" & _
'"ParamArray Partial Private Property Protected Public" & _
'"RaiseEvent ReadOnly ReDim REM RemoveHandler Resume Return" & _
'"SByte Select Set Shadows Shared Short Single Static Step Stop String Structure-Einschr‰nkung Structure-Anweisung Sub SyncLock" & _
'"Then Throw Aktion True Try TryCast TypeOfÖIs" & _
'"UInteger ULong UShort Using" & _
'"Variant" & _
'"Wend When While Widening With WithEvents WriteOnly" & _
'"XOr const" & _
'"#Else #ElseIf #End #If" & _
'"= & &= * *= / /= \ \= ^ ^= + += - -= >> >>= << <<="
Private m_VBKeyWordsCol As Collection

Private m_VBKeyWordsPTree As KeywordsPTree

Private m_VBKeyWordsOS As cPrefixTree


Public Sub KeyWords_Fill()
    If Len(VBKeywordsUCase) = 0 Then VBKeywordsUCase = UCase(VBKeywords)
    Set m_VBKeyWordsCol = New Collection
    Dim sa() As String: sa = Split(VBKeywordsUCase, " ")
    Dim i As Long
    For i = LBound(sa) To UBound(sa)
        If Len(sa(i)) Then
            m_VBKeyWordsCol.Add sa(i), sa(i)
        End If
    Next
    KeyWords_FillTree
    KeyWords_FillTreeOS
End Sub

Public Sub KeyWords_FillTree()
'    VBKeywordsUCase = UCase(VBKeywords)
    Set m_VBKeyWordsPTree = New KeywordsPTree
    Dim sa() As String: sa = Split(VBKeywordsUCase, " ")
    Dim i As Long
    For i = LBound(sa) To UBound(sa)
        If Len(sa(i)) Then
            m_VBKeyWordsPTree.Add sa(i)
        End If
    Next
End Sub

Public Sub KeyWords_FillTreeOS()
    Set m_VBKeyWordsOS = New cPrefixTree
    m_VBKeyWordsOS.InitExtraChars "#"
    Dim sa() As String: sa = Split(VBKeywordsUCase, " ")
    Dim i As Long
    For i = LBound(sa) To UBound(sa)
        If Len(sa(i)) Then
            m_VBKeyWordsOS.Add sa(i)
        End If
    Next
End Sub

Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function

Public Sub KeywordsPTree_ToTextBox(aTB As TextBox)
    aTB.Text = m_VBKeyWordsPTree.ToStr
End Sub

Public Function IsVBKeyword_instr(k As String) As Boolean
    ''http://www.xbeat.net/vbspeed/c_InStr.htm
    IsVBKeyword_instr = InStr(1, VBKeywordsUCase, k)
End Function

Public Function IsVBKeyword_col(k As String) As Boolean
    On Error Resume Next
    IsVBKeyword_col = Not IsEmpty(m_VBKeyWordsCol(k))
End Function

Public Function IsVBKeyword_PfxTree(aWord As String) As Boolean
    IsVBKeyword_PfxTree = m_VBKeyWordsPTree.Contains(aWord)
End Function

Public Function IsVBKeyword_OS(aWord As String) As Boolean
    IsVBKeyword_OS = m_VBKeyWordsOS.Exists(aWord)
End Function


Function RndName() As String
    'erzeugt einen zuf‰lligen Namen
    'der erste Buchstabe is ein Groﬂbuchstabe, alle folgenden sind kleinbuchstaben
    Dim s As String: s = Chr(65 + Rnd * 25)
    Dim i As Long
    For i = 1 To Rnd * 5 + 5
        s = s & Chr(97 + Rnd * 25)
    Next
    RndName = s
End Function

Function CreateTestCase() As String()
    Randomize
    Dim ks() As String: ks = Split(MKeywords.VBKeywordsUCase, " ")
    Dim i As Long, j As Long, u As Long: u = UBound(ks)
    ReDim tc(0 To (u + 1) * 3 - 2) As String
    For i = 0 To u - 2
        tc(j) = ks(i):             j = j + 1
        tc(j) = ks(i) & ks(i + 1): j = j + 1
        tc(j) = UCase(RndName):    j = j + 1
    Next
    tc(j) = ks(i):          j = j + 1
    tc(j) = UCase(RndName): j = j + 1
    tc(j) = UCase(RndName): j = j + 1
    tc(j) = UCase(RndName): j = j + 1
    tc(j) = UCase(RndName): j = j + 1
    CreateTestCase = tc
End Function

'Public Function IsInCollection( _
'    ByRef col As Collection, _
'    ByRef elem As String _
'  ) As Boolean
''http://vb-tec.de/collctns.htm
'  On Error Resume Next
'
'    If IsEmpty(col(elem)) Then: 'DoNothing
'    IsInCollection = (Err.Number = 0)
'
'  On Error GoTo 0
'
'End Function
'
