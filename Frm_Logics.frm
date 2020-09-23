VERSION 5.00
Begin VB.Form Frm_Logics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logic Operator Changes."
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   1800
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
      Begin VB.OptionButton Endian 
         Caption         =   "Big Endian"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Endian 
         Caption         =   "Little Endian"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2775
      Begin VB.OptionButton LogicT 
         Caption         =   "XNor"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "Nor"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "Nand"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   21
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "Xor"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "Not"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "Or"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton LogicT 
         Caption         =   "And"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proccess"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox LongB 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox HexB 
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox HexB 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox HexB 
      Height          =   285
      Index           =   1
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox HexB 
      Height          =   285
      Index           =   0
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton Bits 
      Caption         =   "32 Bit Data"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton Bits 
      Caption         =   "24 Bit Data"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton Bits 
      Caption         =   "16 Bit Data"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton Bits 
      Caption         =   "8 Bit Data"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Logic opperators"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Data type and value"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Frm_Logics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Auto_Update As Boolean

Private Sub Bits_Click(Index As Integer)
HexB(0).Enabled = True
Select Case Index
    Case 0
        Logic_Val.Count = 0
        HexB(1).Enabled = False
        HexB(2).Enabled = False
        HexB(3).Enabled = False
        HexB(1).Text = ""
        HexB(2).Text = ""
        HexB(3).Text = ""
    Case 1
        Logic_Val.Count = 1
        HexB(1).Enabled = True
        HexB(2).Enabled = False
        HexB(3).Enabled = False
        HexB(2).Text = ""
        HexB(3).Text = ""
    Case 2
        Logic_Val.Count = 2
        HexB(1).Enabled = True
        HexB(2).Enabled = True
        HexB(3).Enabled = False
        HexB(3).Text = ""
    Case 3
        Logic_Val.Count = 3
        HexB(1).Enabled = True
        HexB(2).Enabled = True
        HexB(3).Enabled = True
End Select
End Sub

Private Sub Command1_Click()
Dim Tmp_Loop As Long
For Tmp_Loop = 0 To 3
    Logic_Val.HexC(Tmp_Loop) = Hex_2_Byte(HexB(Tmp_Loop).Text)
Next Tmp_Loop
Logic_Val.None = False
Unload Me
End Sub

Private Sub Command2_Click()
Logic_Val.None = True
Unload Me
End Sub

Private Sub Endian_Click(Index As Integer)
Select Case Index
    Case 0
        Logic_Val.Endian = Little_Endian
    Case 1
        Logic_Val.Endian = Big_Endian
End Select
End Sub

Private Sub Form_Load()
LogicT(0).Value = True
Endian(0).Value = True
Bits(0).Value = True
End Sub

Private Sub HexB_Change(Index As Integer)
Dim Tmp_Hex As String
Dim Tmp_Loop As Long
If Auto_Update Then Exit Sub
Tmp_Hex = ""
For Tmp_Loop = 3 To 0 Step -1
    Tmp_Hex = Tmp_Hex & Right("00" & Hex_Check(HexB(Tmp_Loop).Text), 2)
Next Tmp_Loop
Auto_Update = True
LongB.Text = (HexToDecade(Tmp_Hex))
Auto_Update = False
If Len(HexB(Index)) >= 2 Then
    Select Case Index
        Case 1 To 3
            HexB(Index - 1).SetFocus
        Case 0
            Command1.SetFocus
    End Select
End If
End Sub

Private Sub HexB_GotFocus(Index As Integer)
HexB(Index).SelStart = 0
HexB(Index).SelLength = 2

End Sub

Private Sub HexB_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii 'Filter inputs.. (0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F)
    Case 1 To 7
        KeyAscii = 0
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

Private Sub LogicT_Click(Index As Integer)
Select Case Index
    Case 0
        Logic_Val.Logics = Logic_And
    Case 1
        Logic_Val.Logics = Logic_Or
    Case 2
        Logic_Val.Logics = Logic_Not
    Case 3
        Logic_Val.Logics = Logic_Xor
    Case 4
        Logic_Val.Logics = logic_Nand
    Case 5
        Logic_Val.Logics = Logic_Nor
    Case 6
        Logic_Val.Logics = Logic_XNor
        
End Select
End Sub

Private Sub LongB_Change()
Dim Tmp_Hex As String
If Auto_Update Then Exit Sub
Auto_Update = True
Tmp_Hex = Right("        " & DecadeToHex(LongB.Text), 8)
HexB(3).Text = Mid(Tmp_Hex, 1, 2) ' Left() could work here too, but mid lets you see what's happening here at a glance
HexB(2).Text = Mid(Tmp_Hex, 3, 2)
HexB(1).Text = Mid(Tmp_Hex, 5, 2)
HexB(0).Text = Mid(Tmp_Hex, 7, 2) ' Right() could work here too.

Auto_Update = False
End Sub

Private Sub LongB_KeyPress(KeyAscii As Integer)
Select Case KeyAscii 'Filter inputs.. (0,1,2,3,4,5,6,7,8,9)
    Case 1 To 7
        KeyAscii = 0
    'Case 8  ' No trap for backspace...
    Case 9 To 47
        KeyAscii = 0
    Case 58 To 205
        KeyAscii = 0
End Select
End Sub
