VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Keyword or Identifier Prefix Tree"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19455
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   19455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestTreeOS 
      Caption         =   "Test TreeOS"
      Height          =   375
      Left            =   17760
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestTPTree 
      Caption         =   "Test TPTree"
      Height          =   375
      Left            =   12480
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton BtnIsEqual 
      Caption         =   "b1 = b2 = b3 ?"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton BtnTestPTree1 
      Caption         =   "Test PfxTree"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestCol 
      Caption         =   "Test Col"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestInstr 
      Caption         =   "Test Instr"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox TxtUJ 
      Alignment       =   2  'Zentriert
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "uj"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   17760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   15
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   14160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   13
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   12480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   8
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Anzahl Testdurchläufe"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tc() As String
Dim uj   As Long
Dim bks1() As Boolean
Dim bks2() As Boolean
Dim bks3() As Boolean
Dim bks4() As Boolean
Dim bks5() As Boolean

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    MKeywords.KeyWords_Fill
    MKwPTree.KeyWords_Fill
    tc = MKeywords.CreateTestCase
    'Text1.Text = Join(tc, vbCrLf)
    TestCase_ToTextBox Text1
    BtnTestInstr.Caption = "Test Instr"
    BtnTestCol.Caption = "Test Col"
    TxtUJ.Text = "1"
    TxtUJ_LostFocus
    MKeywords.KeywordsPTree_ToTextBox Text5
    
    Text7.Text = MKwPTree.KwPTree_ToStr
    
End Sub

Private Sub Form_Resize()
    Dim T As Single: T = Text1.Top
    Dim H As Single: H = Me.ScaleHeight - T
    Dim px1 As Single: px1 = Screen.TwipsPerPixelX
    Dim L As Single, W As Single
    L = L + W - px1: W = Text1.Width
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
    L = L + W - px1: W = Text2.Width
    If W > 0 And H > 0 Then Text2.Move L, T, W, H
    L = L + W - px1: W = Text3.Width
    If W > 0 And H > 0 Then Text3.Move L, T, W, H
    L = L + W - px1: W = Text4.Width
    If W > 0 And H > 0 Then Text4.Move L, T, W, H
    L = L + W - px1: W = Text5.Width
    If W > 0 And H > 0 Then Text5.Move L, T, W, H
    L = L + W - px1: W = Text6.Width
    If W > 0 And H > 0 Then Text6.Move L, T, W, H
    L = L + W - px1: W = Text7.Width
    If W > 0 And H > 0 Then Text7.Move L, T, W, H
    L = L + W - px1: W = Me.ScaleWidth - L 'Text8.Width
    If W > 0 And H > 0 Then Text8.Move L, T, W, H
    
End Sub

Private Sub TestCase_ToTextBox(aTB As TextBox)
    Dim i As Long, s As String
    For i = LBound(tc) To UBound(tc)
        s = s & Format(i, "000") & ": " & tc(i) & IIf(i < UBound(tc), vbCrLf, "")
    Next
    aTB.Text = s
End Sub
'Private Sub BtnTreeUpdateView_Click()
'    MKeywords.Keywords_ToTB Text6
'End Sub

Private Sub TxtUJ_LostFocus()
    uj = CLng(TxtUJ.Text)
    ReDim bks1(0 To uj * UBound(tc) + uj - 1)
    bks2 = bks1
    bks3 = bks2
    bks4 = bks3
    bks5 = bks4
End Sub

Private Sub BtnIsEqual_Click()
    Dim i As Long
    For i = LBound(bks1) To UBound(bks1)
        If bks1(i) <> bks2(i) Then
            If MsgBox("i: " & i, vbOKCancel) = vbCancel Then Exit Sub
        End If
        If bks2(i) <> bks3(i) Then
            If MsgBox("i: " & i, vbOKCancel) = vbCancel Then Exit Sub
        End If
    Next
    MsgBox "OK"
End Sub

Function JoinB(bools() As Boolean, ByVal delimiter) As String
    ReDim s(LBound(bools) To UBound(bools)) As String
    Dim i As Long
    For i = 0 To UBound(s)
        s(i) = Format(i, "000") & ": " & CStr(bools(i))
    Next
    JoinB = Join(s, delimiter)
End Function

Private Sub BtnTestInstr_Click()
    Dim dt As Single: dt = Timer
    Dim i As Long, j As Long, k As Long
    For j = 1 To uj
        For i = LBound(tc) To UBound(tc)
            bks1(k) = MKeywords.IsVBKeyword_instr(tc(i))
            k = k + 1
        Next
    Next
    dt = Timer - dt
    MsgBox Format(dt, "0.000000000")
    Text2.Text = JoinB(bks1, vbCrLf)
End Sub

Private Sub BtnTestCol_Click()
    Dim dt As Single: dt = Timer
    Dim i As Long, j As Long, k As Long
    For j = 1 To uj
        For i = LBound(tc) To UBound(tc)
            bks2(k) = MKeywords.IsVBKeyword_col(tc(i))
            k = k + 1
        Next
    Next
    dt = Timer - dt
    MsgBox Format(dt, "0.000000000")
    Text3.Text = JoinB(bks2, vbCrLf)
End Sub

Private Sub BtnTestPTree1_Click()
    Dim dt As Single: dt = Timer
    Dim i As Long, j As Long, k As Long
    For j = 1 To uj
        For i = LBound(tc) To UBound(tc)
            bks3(k) = MKeywords.IsVBKeyword_PfxTree(tc(i))
            k = k + 1
        Next
    Next
    dt = Timer - dt
    MsgBox Format(dt, "0.000000000")
    Text4.Text = JoinB(bks3, vbCrLf)
End Sub

Private Sub BtnTestTPTree_Click()
    Dim dt As Single: dt = Timer
    Dim i As Long, j As Long, k As Long
    For j = 1 To uj
        For i = LBound(tc) To UBound(tc)
            bks4(k) = MKwPTree.IsVBKeyword_TPTree(tc(i))
            k = k + 1
        Next
    Next
    dt = Timer - dt
    MsgBox Format(dt, "0.000000000")
    Text6.Text = JoinB(bks4, vbCrLf)
End Sub

Private Sub BtnTestTreeOS_Click()
    Dim dt As Single: dt = Timer
    Dim i As Long, j As Long, k As Long
    For j = 1 To uj
        For i = LBound(tc) To UBound(tc)
            bks5(k) = MKeywords.IsVBKeyword_OS(tc(i))
            k = k + 1
        Next
    Next
    dt = Timer - dt
    MsgBox Format(dt, "0.000000000")
    Text8.Text = JoinB(bks5, vbCrLf)
End Sub


