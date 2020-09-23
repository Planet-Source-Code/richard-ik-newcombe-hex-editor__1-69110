VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_Hex_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hex Editor"
   ClientHeight    =   1935
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar Scroll_Pos 
      Height          =   1575
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin RichTextLib.RichTextBox Hex_Pos 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Frm_Hex_Edit2.frx":0000
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
   Begin RichTextLib.RichTextBox Hex_Val 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   65535
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Frm_Hex_Edit2.frx":0080
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
   Begin VB.PictureBox Text_Function 
      Height          =   255
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox HexD 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   65535
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Frm_Hex_Edit2.frx":0100
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Type_Lbl 
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Pos 
      Caption         =   "Pos - "
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Offset 
      Caption         =   "Offset - "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_F_Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Mnu_F_Close 
         Caption         =   "Close"
         Visible         =   0   'False
      End
      Begin VB.Menu Split_File_1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_F_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Mnu_Functions 
      Caption         =   "Functions"
      Begin VB.Menu Mnu_F_Write 
         Caption         =   "Write To File"
         Shortcut        =   ^W
      End
      Begin VB.Menu Mnu_F_Clear 
         Caption         =   "Clear Changes"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu Mnu_Tools 
      Caption         =   "Tools"
      Begin VB.Menu Mnu_Tool_File 
         Caption         =   "File Tools"
         Begin VB.Menu Mnu_F_Find 
            Caption         =   "Find in file"
            Shortcut        =   ^F
         End
         Begin VB.Menu Mnu_F_Find_A 
            Caption         =   "Find Again"
            Shortcut        =   ^G
         End
         Begin VB.Menu Mnu_F_BlCpy 
            Caption         =   "Block Copy"
            Shortcut        =   ^B
         End
         Begin VB.Menu Mnu_F_EOF 
            Caption         =   "Set As EOF"
         End
      End
      Begin VB.Menu PopUp_Menu 
         Caption         =   "PopUp_Menu"
         Begin VB.Menu Popup_Copy 
            Caption         =   "Copy"
         End
         Begin VB.Menu Popup_Paste 
            Caption         =   "Paste"
         End
         Begin VB.Menu Popup_delete 
            Caption         =   "Delete"
         End
         Begin VB.Menu Popup_Fill 
            Caption         =   "Fill With"
         End
         Begin VB.Menu Popup_Clear 
            Caption         =   "Clear Changes"
         End
         Begin VB.Menu Popup_Logics 
            Caption         =   "Logics Edit"
         End
      End
      Begin VB.Menu Mnu_Options 
         Caption         =   "Options"
         Begin VB.Menu Mnu_Opt_Wide 
            Caption         =   "32 Byte Width"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnu_Opt_Long 
            Caption         =   "48 Line Length"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnu_Opt_Fwrite 
            Caption         =   "No Prompt for File Write"
         End
      End
      Begin VB.Menu Mnu_Jump_Pos 
         Caption         =   "Jump to Pos"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
   End
End
Attribute VB_Name = "Frm_Hex_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loop_1 As Long
Private Loop_2 As Long
Private Loop_3 As Long
Private Skip_A_F As Boolean
Private Skip_F_P As Boolean
Private Skip_Read As Boolean
Private File_Pos As Currency  ' See note in Public_Function
Private File_Num As Long
Private File_Open As Boolean
Private File_Change As Boolean
Private File_BLock As Long
Private Hex_Array(4095) As Byte
Private Hex_A_Ori(4095) As Byte
Private Search_Pos As Long
Private Tot_Count As Long
Private Line_Count As Long
Private Loop_Step As Long
Private Scroll_Mul As Long
Private C_S_Open As Boolean
'90 % of the core code was tested and debugged by Wizbang of CG..
'Thanks to Wizbang for hanging in there and spending his time to test this stupid little app..
'There's simply too much code to mark everything Wizbang has helped me with..

Private Sub Form_Load()
Change_View
Hex_Enable (False)
Mnu_Functions.Enabled = False
PopUp_Menu.Visible = False
If Command$ <> "" Then
    C_S_Open = True
    Call Mnu_F_Open_Click
End If
C_S_Open = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If File_Change Then Write_Data File_Pos
If File_Open Then
    API_CloseFile File_Num
End If
End Sub

Private Sub Hex_Val_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If Hex_Val.Enabled = True Then
        PopupMenu PopUp_Menu, , x + Hex_Val.Left, y + Hex_Val.Top
    End If
End If
End Sub

Private Sub HexD_Change()
Dim Tmp_Pos As Long
Dim Tmp_Actual As Long
Dim Char_code As Byte
If Skip_Read Then ' if this is a file read triggered change.. Exit..
    Exit Sub
End If
With HexD
    Tmp_Pos = .SelStart
    Tmp_Actual = Get_H_Actual(.SelStart, Mnu_Opt_Wide.Checked)
    .SelStart = Set_H_Actual(Tmp_Actual, Mnu_Opt_Wide.Checked)
    .SelLength = 2
    Char_code = Hex_2_Byte(.SelText)
    If Char_code <> Hex_A_Ori(Tmp_Actual) Then
        HexD.SelColor = vbRed
    Else
        If Tmp_Actual + File_Pos + 1 <= File_Len Then
            HexD.SelColor = vbBlack
        Else
            HexD.SelColor = RGB(&HD0, &HD0, &HD0)
        End If
    End If
    If Skip_A_F Then ' if this is a code triggered change.. Exit..
        Exit Sub
    End If
    Hex_Array(Tmp_Actual) = Char_code
    If Not File_Change Then
        File_Change = True
        Mnu_Functions.Enabled = True
    End If
    Skip_A_F = True
    Update_Hex_Val Tmp_Actual, Char_code
    Skip_A_F = False
    .SelStart = Tmp_Pos
    .SelLength = 1
    If .SelText = "" Then
        HexD_KeyDown vbKeyRight, 0
    End If
    If .SelText = " " Then
        .SelStart = .SelStart + 1
        .SelLength = 1
        Hex_Val.SelStart = Set_Actual(Get_H_Actual(.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
        Hex_Val.SelLength = 1
    End If
End With
End Sub

Private Sub HexD_Click()
HexD.SelLength = 1
If HexD.SelText = " " Then
    HexD.SelStart = HexD.SelStart + 1
    HexD.SelLength = 1
End If
Hex_Val.SelStart = Set_Actual(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
Hex_Val.SelLength = 1
Pos.Caption = "Pos - " & Str(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + File_Pos)
End Sub

Private Sub HexD_GotFocus()
HexD.BackColor = vbWhite
Pos.Caption = "Pos - " & Str(File_Pos + Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked))
HexD.SelLength = 1
End Sub

Private Sub HexD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Tmp_Sel_Start As Long
Tmp_Sel_Start = Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked)
If Shift And vbCtrlMask Then
    If KeyCode = vbKeyV Then
        KeyCode = 0
    End If
End If
    
If KeyCode = vbKeyBack Then ' Change backspace to left arrow and replace with original data
    Skip_A_F = True
    Hex_Array(Tmp_Sel_Start) = Hex_A_Ori(Tmp_Sel_Start)
    Update_Hex_Val Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Update_HexD Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Call HexD_KeyDown(vbKeyLeft, 0) ' Call to goto previous pos
    Skip_A_F = False
    KeyCode = 0
End If
If KeyCode = vbKeyDelete Then ' replace contents with original data
    Skip_A_F = True
    Hex_Array(Tmp_Sel_Start) = Hex_A_Ori(Tmp_Sel_Start)
    Update_Hex_Val Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Update_HexD Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Skip_A_F = False
    KeyCode = 0
End If
If KeyCode = vbKeyHome Then ' Move to the first box
    If Shift And vbCtrlMask Then
        Home_Block
    End If
    HexD.SetFocus
    HexD.SelStart = 0
    HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyEnd Then ' Move to the last box
    If Shift And vbCtrlMask Then
        End_Block
    End If
    HexD.SetFocus
    If Shift And vbCtrlMask Then
        Hex_Val.SelStart = Set_Actual(Tot_Count / 2, Mnu_Opt_Wide.Checked)
    Else
        Hex_Val.SelStart = Set_Actual(Tot_Count - 1, Mnu_Opt_Wide.Checked)
    End If
    HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyLeft Then ' Move to the next box
    Select Case Tmp_Sel_Start
        Case 1 To (Tot_Count - 1)
            HexD.SelStart = Set_H_Actual(Tmp_Sel_Start - 1, Mnu_Opt_Wide.Checked)
            HexD.SelLength = 1
        Case 0
            If Prev_Line Then
                HexD.SetFocus
                HexD.SelStart = Set_H_Actual(Tmp_Sel_Start + Loop_Step - 1, Mnu_Opt_Wide.Checked)
                HexD.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyRight Then ' move to prev. box
    Select Case Tmp_Sel_Start
        Case 0 To (Tot_Count - 2)
            HexD.SelStart = Set_H_Actual(Tmp_Sel_Start + 1, Mnu_Opt_Wide.Checked)
            HexD.SelLength = 1
        Case (Tot_Count - 1)
            If Next_Line Then
                HexD.SetFocus
                HexD.SelStart = Set_H_Actual(Tmp_Sel_Start - Loop_Step + 1, Mnu_Opt_Wide.Checked)
                HexD.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyDown Then ' move down 1 line
    Select Case Tmp_Sel_Start
        Case 0 To Line_Count - 1
            HexD.SelStart = Set_H_Actual(Tmp_Sel_Start + Loop_Step, Mnu_Opt_Wide.Checked)
            HexD.SelLength = 1
        Case Line_Count To (Tot_Count - 1)
            If Next_Line Then
                HexD.SetFocus
                HexD.SelStart = Set_H_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
                HexD.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyUp Then ' Move up 1 line
    Select Case Tmp_Sel_Start
        Case Loop_Step To (Tot_Count - 1)
            HexD.SelStart = Set_H_Actual(Tmp_Sel_Start - Loop_Step, Mnu_Opt_Wide.Checked)
            HexD.SelLength = 1
        Case 0 To (Loop_Step - 1)
            If Prev_Line Then
                HexD.SetFocus
                HexD.SelStart = Set_H_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
                HexD.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyPageUp Then ' Move up 1 Page
    If Shift And vbCtrlMask Then
        If Shift And vbShiftMask Then
            Prev_Block (100)
        Else
            Prev_Block (10)
        End If
    Else
        Prev_Block (1)
    End If
    HexD.SetFocus
    HexD.SelStart = Set_H_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyPageDown Then ' Move down 1 Page
    If Shift And vbCtrlMask Then
        If Shift And vbShiftMask Then
            Next_Block (100)
        Else
            Next_Block (10)
        End If
    Else
        Next_Block (1)
    End If
    HexD.SetFocus
    HexD.SelStart = Set_H_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 1
    KeyCode = 0
End If
Hex_Val.SelStart = Set_Actual(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
Hex_Val.SelLength = 1
Pos.Caption = "Pos - " & Str(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + File_Pos)
End Sub

Private Sub HexD_KeyPress(KeyAscii As Integer)
Dim Tmp_Actual As Long
HexD.SelLength = 1
If HexD.SelText = "" Then
    HexD.SelStart = Set_H_Actual(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + 1, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 1
    Hex_Val.SelStart = Set_Actual(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
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

Private Sub HexD_LostFocus()
HexD.BackColor = vbYellow
End Sub

Private Sub Hex_Val_Change()
Dim Tmp_Pos As Long
Dim Pre_Actual As Long
Dim Tmp_Actual As Long
Dim Char_code As Byte
If Skip_Read Then ' if this is a file read triggered change.. Exit..
    Exit Sub
End If
With Hex_Val
    Tmp_Pos = .SelStart
    Pre_Actual = Get_Actual(.SelStart, Mnu_Opt_Wide.Checked)
    If .SelStart <> 0 Then
        .SelStart = .SelStart - 1
    End If
    Tmp_Actual = Get_Actual(.SelStart, Mnu_Opt_Wide.Checked)
    .SelLength = 1
    Char_code = Asc(.SelText)
    If .SelText <> Valid_Char(Hex_A_Ori(Tmp_Actual)) Then
        Hex_Val.SelColor = vbRed
    Else
        If Tmp_Actual + File_Pos + 1 <= File_Len Then
            Hex_Val.SelColor = vbBlack
        Else
            Hex_Val.SelColor = RGB(&HD0, &HD0, &HD0)
        End If
    End If
    If Skip_A_F Then ' if this is a code triggered change.. Exit..
        Exit Sub
    End If
    Hex_Array(Tmp_Actual) = Char_code
    If Not File_Change Then
        File_Change = True
        Mnu_Functions.Enabled = True
    End If
    Skip_A_F = True
    Update_HexD Tmp_Actual, Char_code
    Skip_A_F = False
    If Tmp_Actual = Tot_Count - 1 Then
        Hex_Val_KeyDown vbKeyRight, 0
        Exit Sub
    End If
End With
HexD.SelStart = Set_H_Actual(Pre_Actual, Mnu_Opt_Wide.Checked)
HexD.SelLength = 1
Hex_Val.SelStart = Set_Actual(Pre_Actual, Mnu_Opt_Wide.Checked)
Hex_Val.SelLength = 1
Pos.Caption = "Pos - " & Str(File_Pos + Pre_Actual)
End Sub

Private Sub Hex_Val_Click()
HexD.SelStart = Set_H_Actual(Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
HexD.SelLength = 2
Pos.Caption = "Pos - " & Str(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + File_Pos)
End Sub

Private Sub Hex_Val_GotFocus()
Hex_Val.BackColor = vbWhite
End Sub

Private Sub Hex_Val_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Tmp_Sel_Start As Long
Tmp_Sel_Start = Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked)
If Shift And vbCtrlMask Then
    If KeyCode = vbKeyV Then
        Popup_Paste_Click
        KeyCode = 0
    End If
End If
Hex_Val.SelLength = 1
If Hex_Val.SelText = "" Then
    Hex_Val.SelStart = Set_Actual(Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    HexD.SelStart = Set_H_Actual(Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked)
    HexD.SelLength = 1
End If
Tmp_Sel_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
If KeyCode = vbKeyBack Then ' Change backspace to left arrow and replace with original data
    Skip_A_F = True
    Hex_Array(Tmp_Sel_Start) = Hex_A_Ori(Tmp_Sel_Start)
    Update_Hex_Val Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Update_HexD Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Call Hex_Val_KeyDown(vbKeyLeft, 0)
    Skip_A_F = False
    KeyCode = 0
End If
If KeyCode = vbKeyDelete Then ' replace contents with original data
    Skip_A_F = True
    Hex_Array(Tmp_Sel_Start) = Hex_A_Ori(Tmp_Sel_Start)
    Update_Hex_Val Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Update_HexD Tmp_Sel_Start, Hex_A_Ori(Tmp_Sel_Start)
    Skip_A_F = False
    KeyCode = 0
End If
If KeyCode = vbKeyHome Then ' Move to the first box
    If Shift And vbCtrlMask Then
        Home_Block
    End If
    Hex_Val.SetFocus
    Hex_Val.SelStart = 0
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyEnd Then ' Move to the first box
    If Shift And vbCtrlMask Then
        End_Block
    End If
    Hex_Val.SetFocus
    If Shift And vbCtrlMask Then
        Hex_Val.SelStart = Set_Actual(Tot_Count / 2, Mnu_Opt_Wide.Checked)
    Else
        Hex_Val.SelStart = Set_Actual(Tot_Count - 1, Mnu_Opt_Wide.Checked)
    End If
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyLeft Then ' Move to the Prev Char
    Select Case Tmp_Sel_Start
        Case 1 To Tot_Count - 1
            Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start - 1, Mnu_Opt_Wide.Checked)
            Hex_Val.SelLength = 1
        Case 0
            If Prev_Line Then
                Hex_Val.SetFocus
                Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start + Loop_Step - 1, Mnu_Opt_Wide.Checked)
                Hex_Val.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyRight Then ' move to Next. Char
    Select Case Tmp_Sel_Start
        Case 0 To Tot_Count - 2
            Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start + 1, Mnu_Opt_Wide.Checked)
            Hex_Val.SelLength = 1
        Case Tot_Count - 1 To Tot_Count
            If Next_Line Then
                Hex_Val.SetFocus
                Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start - Loop_Step + 1, Mnu_Opt_Wide.Checked)
                Hex_Val.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyDown Then ' move down 1 line
    Select Case Tmp_Sel_Start
        Case 0 To Line_Count - 1
            Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start + Loop_Step, Mnu_Opt_Wide.Checked)
            Hex_Val.SelLength = 1
        Case Line_Count To Tot_Count - 1
            If Next_Line Then
                Hex_Val.SetFocus
                Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
                Hex_Val.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyUp Then ' Move up 1 line
    Select Case Tmp_Sel_Start
        Case Loop_Step To Tot_Count - 1
            Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start - Loop_Step, Mnu_Opt_Wide.Checked)
            Hex_Val.SelLength = 1
        Case 0 To Loop_Step - 1
            If Prev_Line Then
                Hex_Val.SetFocus
                Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
                Hex_Val.SelLength = 1
            End If
    End Select
    KeyCode = 0
End If
If KeyCode = vbKeyDelete Then ' Delete key pressed. Put H00 in place
    HexD_Set(Tmp_Sel_Start) = "00"
    KeyCode = 0
End If
If KeyCode = vbKeyPageUp Then ' Move up 1 Page
    Prev_Block (1)
    Hex_Val.SetFocus
    Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
If KeyCode = vbKeyPageDown Then ' Move down 1 Page
    Next_Block (1)
    Hex_Val.SetFocus
    Hex_Val.SelStart = Set_Actual(Tmp_Sel_Start, Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    KeyCode = 0
End If
    HexD.SelStart = Set_H_Actual(Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked) ' Update Hex position
    HexD.SelLength = 1
Pos.Caption = "Pos - " & Str(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + File_Pos)
End Sub

Private Sub Hex_Val_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 1 To 31
        KeyAscii = 0
    Case 127 To 186
        KeyAscii = 0
End Select
End Sub

Private Sub Hex_Val_LostFocus()
Hex_Val.BackColor = vbYellow
End Sub

Private Sub Mnu_F_BlCpy_Click()
Dim Block_S As Currency
Dim Block_E As Currency
Dim Block_T As Currency
Dim Block_L As Long
Dim Data() As Byte
Block_S = Val(InputBox("Please Enter Start Pos (Dec)", , "0"))
Block_E = Val(InputBox("Please Enter End Pos (Dec)", , "0"))
Block_T = Val(InputBox("Please Enter Target Pos (Dec)", , "0"))
If Block_E < Block_S Then Exit Sub
Block_L = CLng(Block_E - Block_S)
ReDim Data(Block_L)
API_ReadFile File_Num, Block_S, Block_L, Data()
API_WriteFile File_Num, Block_T, Block_L, Data()
End Sub

Private Sub Mnu_F_Clear_Click()
Skip_A_F = True
Read_Data File_Pos
Mnu_Functions.Enabled = False
End Sub

Private Sub Mnu_F_Close_Click()
If File_Change Then Write_Data File_Pos
API_CloseFile File_Num
File_Open = False
Me.Caption = "Hex Editor"
Mnu_F_Close.Visible = False
Hex_Enable (False)
Me.Mnu_Jump_Pos.Enabled = False
End Sub

Private Sub Mnu_F_EOF_Click()
Dim New_EOF As Currency
Dim Tmp_Msg As Long
New_EOF = File_Pos + Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + 1
Tmp_Msg = MsgBox("All Data from Position:" & New_EOF & " will be lost." & vbCrLf & "Are you Sure?", vbOKCancel)
If Tmp_Msg = vbCancel Then Exit Sub
If File_Change Then Write_Data File_Pos
API_SetEndOfFile File_Num, New_EOF
File_Len = New_EOF
    Scroll_Pos.Max = IIf(File_Len - (Tot_Count / 2) > 1, ((File_Len - (Tot_Count / 2)) / Scroll_Mul), 0)
Read_Data File_Pos
End Sub

Private Sub Mnu_F_Exit_Click()
Unload Me
End Sub
Private Function Search_Func(Search_Value As String, Search_Loc As Long) As Boolean
Dim File_string As String
Dim Tmp_F_Pos As Long
Dim Find_Pos As Long
Hex_Val.SetFocus
Tmp_F_Pos = File_Pos
Search_Func = False
Do
File_string = StrConv(Hex_A_Ori(), vbUnicode)
File_string = Left(File_string, File_BLock)
Find_Pos = InStr(Search_Loc, File_string, Search_Value, Search_Type)
If Find_Pos > 0 Then
    If Find_Pos < Tot_Count Then
        Read_Data File_Pos, False
        HexD.SelStart = Set_H_Actual(Find_Pos - 1, Mnu_Opt_Wide.Checked)
        HexD.SelLength = 1
        Hex_Val.SelStart = Set_Actual(Find_Pos - 1, Mnu_Opt_Wide.Checked)
        Hex_Val.SelLength = Search_Len
    Else
        Next_Block (1)
        File_string = StrConv(Hex_A_Ori(), vbUnicode)
        File_string = Left(File_string, Tot_Count)
        Find_Pos = InStr(1, File_string, Search_Value, Search_Type) ' find the new location of the searched item
        HexD.SelStart = Set_H_Actual(Find_Pos - 1, Mnu_Opt_Wide.Checked)
        HexD.SelLength = 1
        Hex_Val.SelStart = Set_Actual(Find_Pos - 1, Mnu_Opt_Wide.Checked)
        Hex_Val.SelLength = Search_Len
    End If
    Search_Func = True
    Hex_Val.SetFocus
    Exit Function
End If
Search_Loc = Tot_Count - Search_Len
Loop Until Not (Next_Search_Block(True))
File_Pos = Tmp_F_Pos
Read_Data File_Pos, True
Hex_Val.SetFocus
End Function

Private Sub Mnu_F_Find_A_Click()
Dim Search_string As String
If Search_Len = 0 Then
    Mnu_F_Find_Click
    Exit Sub
End If
Search_Pos = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked) + 2
Hex_Val.SetFocus
Search_string = StrConv(Search_H(), vbUnicode)
Search_string = Left(Search_string, Search_Len)
If Not (Search_Func(Search_string, Search_Pos)) Then
    MsgBox "No more found in File"
End If
End Sub

Private Sub Mnu_F_Find_Click()
Dim Search_string As String
Dim File_string As String
Dim Tmp_F_Pos As Long
Dim Find_Pos As Long
Search_Pos = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked) + 1
Search_Len = 0
Frm_Search.Show vbModal, Me
If Search_Len = 0 Then Exit Sub
Search_string = StrConv(Search_H(), vbUnicode)
Search_string = Left(Search_string, Search_Len)
If Not (Search_Func(Search_string, Search_Pos)) Then
    MsgBox "Search String Not Found in File"
End If
End Sub

Private Sub Mnu_F_Open_Click()
Dim TmpVal1 As Currency
Dim Tmp_File As String
On Error Resume Next
If Not (C_S_Open) Then
    CD.CancelError = True
    CD.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
End If
If File_Open Then
    Call Mnu_F_Close_Click
End If
If C_S_Open Then
    Tmp_File = Replace(Command$, Chr(34), "", 1, 2, vbBinaryCompare)
    API_OpenFile Tmp_File, File_Num, File_Len
Else
    API_OpenFile CD.FileName, File_Num, File_Len
End If
If File_Num = -1 Then
    MsgBox "Error Opening file - " & Command$, vbCritical
    Exit Sub
End If
'Scroll_Mul = 16
Scroll_Mul = (Int((File_Len / &H80000)) + 1) * 16
Scroll_Pos.Min = 0
    Scroll_Pos.Max = IIf(File_Len - (Tot_Count / 2) > 1, ((File_Len - (Tot_Count / 2)) / Scroll_Mul), 0)
File_sig = GetFileType(CD.FileName)
Type_Lbl.Caption = "Type: =" & F_Type_String(File_sig)
Mnu_F_Close.Visible = True
Mnu_Jump_Pos.Enabled = True
Hex_Enable (True)
File_Open = True
If C_S_Open Then
    Me.Caption = "Hex Editor - " & Tmp_File
Else
    Me.Caption = "Hex Editor - " & CD.FileName
End If
File_Pos = 0
Read_Data File_Pos
HexD.SetFocus
On Error GoTo 0
End Sub

Private Sub Update_Hex_Val(Index As Long, Asc_Code As Byte)
Dim Val_index As Integer
Dim Tmp_Text
Dim Tmp_Pos As Long
With Hex_Val
    .SelStart = Set_Actual(Index, Mnu_Opt_Wide.Checked)
    .SelLength = 1
    .SelText = Valid_Char(Asc_Code)
    .SelStart = Set_Actual(Index, Mnu_Opt_Wide.Checked)
'    .SelLength = 1
End With
End Sub

Private Sub Update_HexD(Index As Long, Asc_Code As Byte)
Dim Val_index As Integer
Dim Tmp_Text As String
Dim Tmp_Pos As Long
With HexD
    .SelStart = Set_H_Actual(Index, Mnu_Opt_Wide.Checked)
    .SelLength = 2
    .SelText = Right("00" & Hex(Asc_Code), 2)
    .SelStart = Set_H_Actual(Index, Mnu_Opt_Wide.Checked)
'    .SelLength = 2
End With
End Sub

Private Sub Hex_Enable(State As Boolean)
Hex_Pos.Enabled = False
HexD.Enabled = State
Hex_Val.Enabled = State
Scroll_Pos.Enabled = State
Mnu_Tool_File.Enabled = State
End Sub

Private Sub Read_Data(Offset_val As Currency, Optional Skip_Display As Boolean = False)
Dim Tmp_Count As Long
Dim Tmp_Data As Byte
Dim Tmp_Pos As String
Dim Tmp_Val As String
Dim Tmp_Text As String
MousePointer = 11
Offset.Caption = "Offset - " & Str(Offset_val)
Tmp_Count = 0
Skip_A_F = True ' Turn on Skip funtions - Speeds up reading data
Skip_Read = True ' Turn on Skip funtions - Speeds up reading data
File_BLock = Tot_Count * 2
API_ReadFile File_Num, Offset_val, File_BLock, Hex_Array()
If File_BLock < Tot_Count * 2 Then ' Fill in blank spaces with "00"
    For Loop_2 = File_BLock To (Tot_Count * 2) - 1
        Hex_Array(Loop_2) = 0
    Next Loop_2
End If
' Use copy mem API here... Set Hex_A_Ori() = Hex_Array()
CopyMemory Hex_A_Ori(0), Hex_Array(0), (Tot_Count * 2) ' Quick and easy
If Not (Skip_Display) Then
    HexD.Visible = False
    Hex_Val.Visible = False
    Hex_Pos.Text = ""
    HexD.Text = ""
    Hex_Val.Text = ""
    HexD.SelStart = 0                    ' the next 6 lines of code should take care of the color problem picked up by Wizbang.
    HexD.SelLength = Len(HexD.Text)      ' RTbox does not have a explicit Text color (ForeColor) and has to be set in code..
    HexD.SelColor = RGB(&H0, &H0, &H0)   ' Set default Black text...
    Hex_Val.SelStart = 0
    Hex_Val.SelLength = Len(Hex_Val.Text)
    Hex_Val.SelColor = RGB(&H0, &H0, &H0)
    For Loop_3 = 0 To (Tot_Count) - 1 Step Loop_Step
        Tmp_Text = ""
        Tmp_Val = ""
        Tmp_Pos = Right("000000000000" & DecadeToHex(Str(Offset_val + Loop_3)), 12)
        For Loop_2 = 0 To Loop_Step - 1
            Tmp_Text = Tmp_Text & " " & Right("00" & Hex(Hex_Array(Loop_3 + Loop_2)), 2)
            Tmp_Val = Tmp_Val & Valid_Char(Hex_Array(Loop_3 + Loop_2))
        Next Loop_2
        Hex_Pos.Text = Hex_Pos.Text & IIf(Hex_Pos.Text = "", "", vbCrLf) & Trim(Tmp_Pos)
        HexD.Text = HexD.Text & IIf(HexD.Text = "", "", vbCrLf) & Trim(Tmp_Text)
        Hex_Val.Text = Hex_Val.Text & IIf(Hex_Val.Text = "", "", vbCrLf) & Tmp_Val
    Next Loop_3
    ' If Hex is past file end , Use a different display color...
    If File_Pos + Tot_Count > File_Len Then
        HexD.SelStart = Set_H_Actual((File_Len - File_Pos), Mnu_Opt_Wide.Checked)
        HexD.SelLength = Len(HexD.Text) - HexD.SelStart
        HexD.SelColor = RGB(&HD0, &HD0, &HD0)
        Hex_Val.SelStart = Set_Actual((File_Len - File_Pos), Mnu_Opt_Wide.Checked)
        Hex_Val.SelLength = Len(Hex_Val.Text) - Hex_Val.SelStart
        Hex_Val.SelColor = RGB(&HD0, &HD0, &HD0)
    End If
    HexD.SelLength = 2
    Hex_Val.SelLength = 1
    Pos.Caption = "Pos - " & Str(Get_H_Actual(HexD.SelStart, Mnu_Opt_Wide.Checked) + File_Pos)

    HexD.Visible = True
    Hex_Val.Visible = True
End If
Tmp_Count = (File_Pos / Scroll_Mul)
If Tmp_Count > Scroll_Pos.Max Then Tmp_Count = Scroll_Pos.Max
Scroll_Pos.Value = Tmp_Count
Skip_A_F = False
Skip_Read = False
File_Change = False
Mnu_Functions.Enabled = False
MousePointer = 0
End Sub

Private Sub Write_Data(Offset_val As Currency)
Dim Tmp_Count As Long
Dim Tmp_Msg As Long
If Not Skip_F_P Then
    Tmp_Msg = MsgBox("Data has Changed - Write to file", vbYesNo)
    If Tmp_Msg = vbCancel Then Exit Sub
End If
If File_BLock < Tot_Count Then
    Tmp_Count = File_BLock
    For Loop_1 = File_BLock + 1 To Tot_Count
        If Hex_Array(Loop_1 - 1) <> 0 Then
            Tmp_Count = Loop_1
        End If
    Next Loop_1
    If Tmp_Count <> File_BLock Then
        If Skip_F_P Then
            Tmp_Msg = vbOK
        Else
            Tmp_Msg = MsgBox("File Size has Changed - Write  Extended data to file", vbOKCancel)
        End If
        If Tmp_Msg = vbOK Then
            File_BLock = Tmp_Count
        End If
    End If
End If
            
API_WriteFile File_Num, Offset_val, File_BLock, Hex_Array()
API_FileSize File_Num, File_Len
    Scroll_Pos.Max = IIf(File_Len - (Tot_Count / 2) > 1, ((File_Len - (Tot_Count / 2)) / Scroll_Mul), 0)
File_Change = False
Mnu_Functions.Enabled = False
End Sub

Private Property Let HexD_Set(Offset As Long, Hexadecimal As String)
With HexD
    .SelStart = Set_H_Actual(Offset, Mnu_Opt_Wide.Checked)
    .SelLength = 2
    .SelColor = vbBlack
    .SelText = Right("00" & Hexadecimal, 2)
End With
End Property

'Private Property Get HexA_Set(Offset As Long) As String
'HexA_Set = (Right("00" & Hex(Hex_Array(Offset)), 2))
'End Property

Private Property Let HexA_Set(Offset As Long, Hexadecimal As String)
Hex_Array(Offset) = Hex_2_Byte(Right("00" & Hexadecimal, 2))
End Property

Private Sub Check_Hex_Val_Col(Offset As Long)
Hex_Val.SelStart = Set_Actual(Offset, Mnu_Opt_Wide.Checked)
Hex_Val.SelLength = 1
If Hex_A_Ori(Offset) <> Hex_Array(Offset) Then
    Hex_Val.SelColor = vbRed
Else
    Hex_Val.SelColor = vbBlack
End If
End Sub

Private Sub Next_Block(Num_Blocks As Integer)
If File_Change Then Write_Data File_Pos
File_Pos = File_Pos + (Tot_Count * Num_Blocks)
If File_Pos > File_Len - (Tot_Count / 2) Then
    File_Pos = IIf(File_Len > (Tot_Count / 2), (File_Len - (Tot_Count / 2)), 0)
End If
Read_Data File_Pos
End Sub

Private Function Next_Search_Block(Optional Skip_Display As Boolean = False) As Boolean
Next_Search_Block = False
If File_Change Then Write_Data File_Pos
If File_Pos < (File_Len - (Tot_Count / 2)) Then
    File_Pos = File_Pos + Tot_Count
    Read_Data File_Pos, Skip_Display
    Next_Search_Block = True
End If
End Function

Private Function Next_Line() As Boolean
If File_Change Then Write_Data File_Pos
If File_Pos + Loop_Step > File_Len - (Tot_Count / 2) Then
    Next_Line = False
    Exit Function
End If
File_Pos = File_Pos + Loop_Step
Read_Data File_Pos
Next_Line = True
End Function

Private Sub Home_Block()
If File_Change Then Write_Data File_Pos
File_Pos = 0
Read_Data File_Pos
End Sub

Private Sub End_Block()
If File_Change Then Write_Data File_Pos
File_Pos = IIf(File_Len > (Tot_Count / 2), (File_Len - (Tot_Count / 2)), 0)
Read_Data File_Pos
End Sub

Private Sub Prev_Block(Num_Blocks As Integer)
If File_Change Then Write_Data File_Pos
File_Pos = File_Pos - (Tot_Count * Num_Blocks)
If File_Pos < 0 Then File_Pos = 0
Read_Data File_Pos
End Sub

Private Function Prev_Line() As Boolean
If File_Change Then Write_Data File_Pos
If File_Pos = 0 Then
    Prev_Line = False
    Exit Function
End If
File_Pos = File_Pos - Loop_Step
If File_Pos < 0 Then File_Pos = 0
Read_Data File_Pos
Prev_Line = True
End Function

Private Sub Mnu_F_Write_Click()
Skip_F_P = True
Write_Data File_Pos
Skip_F_P = True
Read_Data File_Pos, False
Skip_F_P = True
'With HexD
'    .SelStart = 0
'    .SelLength = Len(.Text)
'    .SelColor = vbBlack
'    .SelLength = 1
'End With
'With Hex_Val
'    .SelStart = 0
'    .SelLength = Len(.Text)
'    .SelColor = vbBlack
'    .SelLength = 1
'End With
Mnu_Functions.Enabled = False
Skip_F_P = Mnu_Opt_Fwrite.Checked
File_Change = False
End Sub

Private Sub Mnu_Jump_Pos_Click()
Frm_Jump.Show vbModal, Me
If Jump_Loc = -1 Then Exit Sub
If Mnu_Opt_Wide.Checked Then
    File_Pos = Jump_Loc And &HFFFFFFC0
Else
    File_Pos = Jump_Loc And &HFFFFFFE0
End If
If File_Pos > (File_Len - (Tot_Count / 2)) And &H7FFFFFF0 Then
    File_Pos = IIf(File_Len > (Tot_Count / 2), (File_Len - (Tot_Count / 2)) And &H7FFFFFF0, 0)
End If
Read_Data File_Pos
Hex_Val.SetFocus
    Hex_Val.SelStart = Set_Actual(Jump_Loc - File_Pos, Mnu_Opt_Wide.Checked)
Hex_Val.SelLength = 1
HexD.SelStart = Set_H_Actual(Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked), Mnu_Opt_Wide.Checked) ' Update Hex position
HexD.SelLength = 1
End Sub

Private Sub Mnu_Opt_Fwrite_Click()
Mnu_Opt_Fwrite.Checked = Not Mnu_Opt_Fwrite.Checked
Skip_F_P = Mnu_Opt_Fwrite.Checked
End Sub

Private Sub Mnu_Opt_Long_Click()
Mnu_Opt_Long.Checked = Not Mnu_Opt_Long.Checked
Change_View
If HexD.Enabled Then
    Read_Data File_Pos
    HexD.SetFocus
End If
End Sub

Private Sub Mnu_Opt_Wide_Click()
Mnu_Opt_Wide.Checked = Not Mnu_Opt_Wide.Checked
Mnu_Opt_Long.Enabled = Mnu_Opt_Wide.Checked
Change_View
If HexD.Enabled Then
    Read_Data File_Pos
    HexD.SetFocus
End If
End Sub

Private Sub Change_View()
Dim Tmp_Text As String
Dim MultiP As Long
Dim MultiL As Long
Text_Function.Font = HexD.Font
Tmp_Text = "123456789012"
Hex_Pos.Width = (Text_Function.TextWidth(Tmp_Text)) + 40
HexD.Left = Hex_Pos.Left + Hex_Pos.Width + 25
If Mnu_Opt_Wide.Checked Then
    MultiP = 2
    Loop_Step = 32
    If Mnu_Opt_Long.Checked Then
        Tot_Count = 1536
        Line_Count = 1504
        MultiL = 3
        Scroll_Pos.SmallChange = 2
        Scroll_Pos.LargeChange = 96
    Else
        Tot_Count = 1024
        Line_Count = 992
        MultiL = 2
        Scroll_Pos.SmallChange = 2
        Scroll_Pos.LargeChange = 64
    End If
Else
    Tot_Count = 512
    Line_Count = 496
    Loop_Step = 16
    MultiP = 1
    MultiL = 2
    Scroll_Pos.SmallChange = 1
    Scroll_Pos.LargeChange = 32
End If
Tmp_Text = "00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F "
HexD.Width = (Text_Function.TextWidth(Tmp_Text) * MultiP) + 40
Tmp_Text = "00" & vbCrLf & "00" & vbCrLf & "00" & vbCrLf & "00" & vbCrLf & "00" & vbCrLf & "00" & vbCrLf & "00" & vbCrLf & "00"
Hex_Pos.Height = Text_Function.TextHeight(Tmp_Text) * 2 * MultiL
HexD.Top = Hex_Pos.Top
HexD.Height = Hex_Pos.Height
Tmp_Text = "0123456789ABCDEF"
Hex_Val.Width = (Text_Function.TextWidth(Tmp_Text) * MultiP) + 50
Hex_Val.Height = Hex_Pos.Height
Hex_Val.Left = HexD.Left + HexD.Width + 25
Hex_Val.Top = HexD.Top
Scroll_Pos.Top = HexD.Top
Scroll_Pos.Left = Hex_Val.Left + Hex_Val.Width + 10
Scroll_Pos.Height = Hex_Pos.Height
Me.Width = Scroll_Pos.Left + Scroll_Pos.Width + 200 'Size the form to fit around the TextBoxes..
Me.Height = HexD.Top + HexD.Height + 800
If File_Open Then
    Scroll_Pos.Max = IIf(File_Len - (Tot_Count / 2) > 1, ((File_Len - (Tot_Count / 2)) / Scroll_Mul), 0)
End If
End Sub

Private Sub Popup_Clear_Click()
' Custom Popup Clear changes code
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Get_Actual(Hex_Val.SelStart + Hex_Val.SelLength, Mnu_Opt_Wide.Checked) - 1
Skip_A_F = True
For Loop_1 = Tmp_Start To Tmp_Fin
    HexD.SelStart = Set_H_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 2
    HexD.SelText = Right("00" & Hex(Hex_A_Ori(Loop_1)), 2)
    Hex_Val.SelStart = Set_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    Hex_Val.SelText = Valid_Char(Hex_A_Ori(Loop_1))
Next Loop_1
Hex_Val.SetFocus
Skip_A_F = False
End Sub

Private Sub Popup_Copy_Click()
' Custom Popup copy code
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Dim Tmp_Str As String
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Get_Actual(Hex_Val.SelStart + Hex_Val.SelLength, Mnu_Opt_Wide.Checked) - 1
Tmp_Str = ""
For Loop_1 = Tmp_Start To Tmp_Fin
    Tmp_Str = Tmp_Str + Chr(Hex_Array(Loop_1))
Next Loop_1
Clipboard.Clear
Clipboard.SetText Tmp_Str
End Sub

Private Sub Popup_delete_Click()
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Get_Actual(Hex_Val.SelStart + Hex_Val.SelLength, Mnu_Opt_Wide.Checked) - 1
Skip_A_F = True
For Loop_1 = Tmp_Start To Tmp_Fin
    HexD.SelStart = Set_H_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 2
    HexD.SelText = "00"
    Hex_Val.SelStart = Set_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    Hex_Val.SelText = Valid_Char(Hex_Array(Loop_1))
Next Loop_1
Skip_A_F = False
File_Change = True
Mnu_Functions.Enabled = True
End Sub

Private Sub Popup_Fill_Click()
' Custom Popup Fill with ? code
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Dim Tmp_Hex As String
Frm_Fill.Show vbModal, Me
If Fill_Code = -1 Then Exit Sub
Tmp_Hex = Right("00" & Hex(Fill_Code), 2)
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Get_Actual(Hex_Val.SelStart + Hex_Val.SelLength, Mnu_Opt_Wide.Checked) - 1
For Loop_1 = Tmp_Start To Tmp_Fin
    HexD.SelStart = Set_H_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 2
    HexD.SelText = Tmp_Hex
Next Loop_1
File_Change = True
Mnu_Functions.Enabled = True
End Sub

Private Sub Popup_Logics_Click()
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Dim Hex_Count As Long
Frm_Logics.Show vbModal, Me
If Logic_Val.None Then Exit Sub
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Get_Actual(Hex_Val.SelStart + Hex_Val.SelLength, Mnu_Opt_Wide.Checked) - 1
If Logic_Val.Endian = Little_Endian Then
    Hex_Count = 0
Else
    Hex_Count = Logic_Val.Count
End If
Skip_A_F = False
For Loop_1 = Tmp_Start To Tmp_Fin
    HexD.SelStart = Set_H_Actual(Loop_1, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 2
    Select Case Logic_Val.Logics
        Case Logic_And
            HexD.SelText = Right("00" & Hex(Hex_Array(Loop_1) And Logic_Val.HexC(Hex_Count)), 2)
        Case logic_Nand
            HexD.SelText = Right("00" & Hex(Not (Hex_Array(Loop_1) And Logic_Val.HexC(Hex_Count))), 2)
        Case Logic_Or
            HexD.SelText = Right("00" & Hex(Hex_Array(Loop_1) Or Logic_Val.HexC(Hex_Count)), 2)
        Case Logic_Nor
            HexD.SelText = Right("00" & Hex(Not (Hex_Array(Loop_1) Or Logic_Val.HexC(Hex_Count))), 2)
        Case Logic_Xor
            HexD.SelText = Right("00" & Hex(Hex_Array(Loop_1) Xor Logic_Val.HexC(Hex_Count)), 2)
        Case Logic_XNor
            HexD.SelText = Right("00" & Hex(Not (Hex_Array(Loop_1) Xor Logic_Val.HexC(Hex_Count))), 2)
        Case Logic_Not
            HexD.SelText = Right("00" & Hex(Not (Hex_Array(Loop_1))), 2)  ' Logic 'not' has single input..
    End Select
    Select Case Logic_Val.Endian
        Case Little_Endian
            Hex_Count = Hex_Count + 1
            If Hex_Count > Logic_Val.Count Then
                Hex_Count = 0
            End If
        Case Big_Endian
            Hex_Count = Hex_Count - 1
            If Hex_Count < 0 Then
                Hex_Count = Logic_Val.Count
            End If
    End Select
Next Loop_1
Skip_A_F = False
File_Change = True
Mnu_Functions.Enabled = True
End Sub

Private Sub Popup_Paste_Click()
' Custom Popup Paste code
Dim Char_code As Byte
Dim Tmp_Loop As Long
Dim Tmp_Start As Long
Dim Tmp_Fin As Long
Dim Tmp_Str As String
Dim Msg_ret As Long
If Not Clipboard.GetFormat(vbCFText) Then Exit Sub
Tmp_Str = Clipboard.GetText(vbCFText)
Tmp_Start = Get_Actual(Hex_Val.SelStart, Mnu_Opt_Wide.Checked)
Tmp_Fin = Len(Tmp_Str)
If Tmp_Fin + Tmp_Start > Tot_Count - 1 Then
    Msg_ret = MsgBox("Text in Clipboard is bigger than buffer size" & vbCrLf & "The Paste canot be undone" _
            & vbCrLf & " Continue anyway ?", vbOKCancel)
    If Msg_ret = vbCancel Then
        Exit Sub
    End If
    Skip_F_P = True
End If
Skip_A_F = True
For Tmp_Loop = 0 To Tmp_Fin - 1
    If (Tmp_Start + Tmp_Loop) = Tot_Count Then
        File_Change = True
        Next_Block (1)
        Tmp_Start = Tmp_Start - Tot_Count
        Skip_A_F = True
    End If
    Char_code = Asc(Mid(Tmp_Str, Tmp_Loop + 1, 1))
    Hex_Array(Tmp_Start + Tmp_Loop) = Char_code
    HexD.SelStart = Set_H_Actual(Tmp_Start + Tmp_Loop, Mnu_Opt_Wide.Checked)
    HexD.SelLength = 2
    HexD.SelText = Right("00" & Hex(Char_code), 2)
    Hex_Val.SelStart = Set_Actual(Tmp_Start + Tmp_Loop, Mnu_Opt_Wide.Checked)
    Hex_Val.SelLength = 1
    Hex_Val.SelText = Valid_Char(Char_code)
Next Tmp_Loop
Skip_F_P = Mnu_Opt_Fwrite.Checked
Skip_A_F = False
File_Change = True
Mnu_Functions.Enabled = True
End Sub

Private Sub Scroll_Pos_Change()
If File_Change Then Write_Data File_Pos
If Skip_Read Then
    Exit Sub
End If
File_Pos = CLng(Scroll_Pos.Value)
File_Pos = File_Pos * Scroll_Mul
If File_Pos > File_Len - (Tot_Count / 2) Then
    File_Pos = IIf(File_Len > (Tot_Count / 2), (File_Len - (Tot_Count / 2)) And &H7FFFFFF0, 0)
End If
Read_Data File_Pos
End Sub

Private Sub Scroll_Pos_Scroll()
If Skip_Read Then
    Exit Sub
End If
File_Pos = Scroll_Pos.Value
File_Pos = File_Pos * Scroll_Mul
If File_Pos > File_Len - (Tot_Count / 2) Then
    File_Pos = IIf(File_Len > (Tot_Count / 2), (File_Len - (Tot_Count / 2)) And &H7FFFFFF0, 0)
End If
Read_Data File_Pos
End Sub
