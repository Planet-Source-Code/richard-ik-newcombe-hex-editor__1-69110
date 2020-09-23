VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_Search 
   Caption         =   "Find String in file."
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Ignore Case"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox Hex_Val 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1508
      _Version        =   393217
      MaxLength       =   100
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Frm_Search.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox HexD 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      MaxLength       =   299
      TextRTF         =   $"Frm_Search.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search String.."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Frm_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Skip_A_U As Boolean
Private Loop_1 As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Search_Type = vbTextCompare
Else
    Search_Type = vbBinaryCompare
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Search_Len = 0
Unload Me
End Sub

Private Sub Form_Load()
Search_Type = vbBinaryCompare
End Sub

Private Sub HexD_Change()
Dim Tmp_Pos As Long
Dim Tmp_Actual As Long
With HexD
    Tmp_Actual = Get_S_Actual(.SelStart)
    Tmp_Pos = .SelStart
    .SelStart = Set_S_Actual(Tmp_Actual)
    .SelLength = 2
    If Skip_A_U Then ' if this is a code triggered change.. Exit..
        Exit Sub
    End If
    Search_H(Tmp_Actual) = Hex_2_Byte(.SelText)
        
    Skip_A_U = True
    Update_Hex_Val (Tmp_Actual)
    Skip_A_U = False
    If Tmp_Pos = Set_S_Actual(Tmp_Actual) + 2 Then
        .SelStart = Tmp_Pos
        .SelLength = 1
        .SelText = " "
    End If
    .SelStart = Tmp_Pos
    .SelLength = 1
    If .SelText = " " Then
        .SelStart = .SelStart + 1
        .SelLength = 1
        Hex_Val.SelStart = Get_S_Actual(.SelStart)
        Hex_Val.SelLength = 1
    End If
End With
Search_Len = Len(Hex_Val.Text)
End Sub

Private Sub HexD_Click()
HexD.SelLength = 1
If HexD.SelText = " " Then
    HexD.SelStart = HexD.SelStart + 1
    HexD.SelLength = 1
End If
Hex_Val.SelStart = Set_Actual(Get_H_Actual(HexD.SelStart, True), True)
Hex_Val.SelLength = 1
End Sub

Private Sub HexD_GotFocus()
HexD.SelLength = 1
End Sub

Private Sub HexD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Tmp_Sel_Start As Long
Dim Tot_Count As Long
Dim Line_Count As Long
Dim Loop_Step As Long
Tmp_Sel_Start = Get_S_Actual(HexD.SelStart)
If KeyCode = vbKeyBack Then ' Delete the Hex code move left 1
    If Tmp_Sel_Start >= 1 Then
        HexD.SelStart = Set_S_Actual(Tmp_Sel_Start - 1)
        Call HexD_KeyDown(vbKeyDelete, 0) ' Call to delete previous pos
    End If
    KeyCode = 0
End If
If KeyCode = vbKeyDelete Then ' From current position shift everything to the left 1 position
    Skip_A_U = True
    HexD.SelStart = Set_S_Actual(Tmp_Sel_Start)
    HexD.SelLength = 3
    HexD.SelText = ""
    Hex_Val.SelStart = Tmp_Sel_Start
    Hex_Val.SelLength = 1
    Hex_Val.SelText = ""
    For Loop_Step = Tmp_Sel_Start To Search_Len - 1
        Search_H(Loop_Step) = Search_H(Loop_Step + 1)
    Next Loop_Step
    Skip_A_U = False
    KeyCode = 0
End If
If KeyCode = vbKeyHome Then ' Move to the first box
    HexD.SetFocus
    HexD.SelStart = 0
    HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyEnd Then ' Move to the last box
    HexD.SetFocus
    HexD.SelStart = Set_S_Actual(Search_Len)
    HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyLeft Then ' Move to the prev box
    If Tmp_Sel_Start > 0 Then
        HexD.SelStart = Set_S_Actual(Tmp_Sel_Start - 1)
        HexD.SelLength = 1
    End If
    KeyCode = 0
End If
If KeyCode = vbKeyRight Then ' move to next. box
            HexD.SelStart = Set_S_Actual(Tmp_Sel_Start + 1)
            HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyDown Then ' move down 1 line
End If
If KeyCode = vbKeyUp Then ' Move up 1 line
End If
If KeyCode = vbKeyPageUp Then ' Move up 1 Page
    KeyCode = 0
End If
If KeyCode = vbKeyPageDown Then ' Move down 1 Page
    KeyCode = 0
End If
    Hex_Val.SelLength = 1
End Sub

Private Sub HexD_KeyPress(KeyAscii As Integer)
Dim Tmp_Actual As Long
HexD.SelLength = 1
If HexD.SelText = "" Then
    HexD.SelStart = Set_H_Actual(Get_H_Actual(HexD.SelStart, True) + 1, True)
    HexD.SelLength = 1
    Hex_Val.SelStart = Set_Actual(Get_H_Actual(HexD.SelStart, True), True)
    Hex_Val.SelLength = 1
End If
If HexD.SelText = " " Then
    HexD.SelStart = HexD.SelStart + 1
    HexD.SelLength = 1
End If
Select Case KeyAscii 'Filter inputs.. (0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F)
    Case 1 To 7
        KeyAscii = 0
    'Case 8  ' No trap for backspace... Done in Keydown above..
    Case 9 To 47
        KeyAscii = 0
    Case 58 To 64
        KeyAscii = 0
    Case 71 To 96
        KeyAscii = 0
    Case 97 To 102
        KeyAscii = KeyAscii - 32 ' Change to upper case code..
    Case 103 To 255
        KeyAscii = 0
End Select
End Sub

Private Sub Hex_Val_Change()
Dim Tmp_Actual As Long
With Hex_Val
    Tmp_Actual = .SelStart
    '.SelLength = 1
    If Skip_A_U Then ' if this is a code triggered change.. Exit..
        Exit Sub
    End If
    Search_Len = Len(.Text)
    Skip_A_U = True
    HexD.Text = ""
    For Loop_1 = 1 To Search_Len
        Search_H(Loop_1 - 1) = Asc(Mid(.Text, Loop_1, 1))
        Update_HexD (Loop_1 - 1)
    Next Loop_1
    Skip_A_U = False
    'Hex_Val_KeyDown vbKeyRight, 0
    .SelStart = Tmp_Actual
End With
HexD.SelStart = Set_S_Actual(Hex_Val.SelStart)
HexD.SelLength = 1
Search_Len = Len(Hex_Val.Text)
End Sub

Private Sub Hex_Val_Click()
'Hex_Val.SelLength = 1
HexD.SelStart = Set_S_Actual(Hex_Val.SelStart)
HexD.SelLength = 2
End Sub


Private Sub Hex_Val_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Tmp_Sel_Start As Long
Dim Tot_Count As Long
Dim Line_Count As Long
Dim Loop_Step As Long
Tmp_Sel_Start = Hex_Val.SelStart
'Hex_Val.SelLength = 1
If KeyCode = vbKeyBack Then
    If Tmp_Sel_Start >= 1 Then
        Hex_Val.SelStart = Tmp_Sel_Start - 1
        Call Hex_Val_KeyDown(vbKeyDelete, 0) ' Call to delete previous pos
    End If
    KeyCode = 0
End If
If KeyCode = vbKeyDelete Then ' Delete contents
    Skip_A_U = True
    HexD.SelStart = Set_S_Actual(Tmp_Sel_Start)
    HexD.SelLength = 3
    HexD.SelText = ""
    Hex_Val.SelStart = Tmp_Sel_Start
    Hex_Val.SelLength = 1
    Hex_Val.SelText = ""
    For Loop_Step = Tmp_Sel_Start To Search_Len - 1
        Search_H(Loop_Step) = Search_H(Loop_Step + 1)
    Next Loop_Step
    Skip_A_U = False
    KeyCode = 0
End If
If KeyCode = vbKeyHome Then ' Move to the first box
    Hex_Val.SetFocus
    Hex_Val.SelStart = 0
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyEnd Then ' Move to the first box
    Hex_Val.SetFocus
'    Hex_Val.SelStart = Set_Actual(Tot_Count, True)
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyLeft Then ' Move to the Prev Char
    Select Case Tmp_Sel_Start
        Case 1 To 100
            Hex_Val.SelStart = Tmp_Sel_Start - 1
            Hex_Val.SelLength = 1
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyRight Then ' move to Next. Char
            Hex_Val.SelStart = Tmp_Sel_Start + 1
            Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyDown Then ' move down 1 line
    Call Hex_Val_KeyDown(vbKeyLeft, 0)
End If
If KeyCode = vbKeyUp Then ' Move up 1 line
    Call Hex_Val_KeyDown(vbKeyLeft, 0)
End If
If KeyCode = vbKeyDelete Then ' Delete key pressed. Put H00 in place
'    HexD_Set(Tmp_Sel_Start) = "00"
    KeyCode = 0
End If
If KeyCode = vbKeyPageUp Then ' Move up 1 Page
    KeyCode = 0
End If
If KeyCode = vbKeyPageDown Then ' Move down 1 Page
    KeyCode = 0
End If
'    HexD.SelLength = 1
End Sub

Private Sub Hex_Val_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 1 To 31
        KeyAscii = 0
    Case 127 To 186
        KeyAscii = 0
End Select
End Sub

Private Sub Update_Hex_Val(Index As Long)
With Hex_Val
    .SelStart = Index
    .SelLength = 1
    .SelText = Valid_Char(Search_H(Index))
    .SelStart = Index
    .SelLength = 1
End With
End Sub

Private Sub Update_HexD(Index As Long)
With HexD
    .SelStart = Set_S_Actual(Index)
    .SelLength = 3
    .SelText = Right("00" & Hex(Search_H(Index)), 2) & " "
    .SelStart = Set_S_Actual(Index)
    .SelLength = 2
End With
End Sub

