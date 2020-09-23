VERSION 5.00
Begin VB.Form Frm_Jump 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jump To Position"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Goto"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Dec_val 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Hex_Val 
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Dec Value"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hex Value"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_Jump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Auto_Update As Boolean

Private Sub Command1_Click()
Jump_Loc = Val(Dec_val.Text)
If Jump_Loc > File_Len Then Jump_Loc = File_Len
Unload Me
End Sub

Private Sub Command2_Click()
Jump_Loc = -1
Unload Me
End Sub

Private Sub Dec_val_Change()
If Auto_Update Then Exit Sub
Auto_Update = True
Hex_Val.Text = DecadeToHex(Dec_val.Text)
Auto_Update = False
End Sub

Private Sub Dec_val_KeyPress(KeyAscii As Integer)
Select Case KeyAscii 'Filter inputs.. (0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F)
    Case 1 To 7
        KeyAscii = 0
    'Case 8  ' No trap for backspace...
    Case 9 To 47
        KeyAscii = 0
    Case 58 To 205
        KeyAscii = 0
End Select
End Sub


Private Sub Hex_Val_Change()
If Auto_Update Then Exit Sub
Auto_Update = True
Dec_val.Text = HexToDecade(Hex_Val.Text)
Auto_Update = False
End Sub

Private Sub Hex_Val_KeyPress(KeyAscii As Integer)
Select Case KeyAscii 'Filter inputs.. (0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F)
    Case 1 To 7
        KeyAscii = 0
    'Case 8  ' No trap for backspace...
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
