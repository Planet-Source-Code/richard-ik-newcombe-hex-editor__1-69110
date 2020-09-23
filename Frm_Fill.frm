VERSION 5.00
Begin VB.Form Frm_Fill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill Code"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Char 
      Height          =   285
      Left            =   360
      MaxLength       =   1
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.ListBox Hex_Code 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Please select Character or hex code to fill with"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_Fill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loop_1 As Long
Private Auto_Chan As Boolean

Private Sub Char_Change()
If Auto_Chan Then Exit Sub
Auto_Chan = True
If Char.Text <> "" Then
    Hex_Code.ListIndex = Asc(Char.Text)
End If
Auto_Chan = False
End Sub

Private Sub Char_GotFocus()
Char.SelStart = 0
Char.SelLength = 1
End Sub

Private Sub Command1_Click()
Fill_Code = Hex_Code.ListIndex
Unload Me
End Sub

Private Sub Command2_Click()
Fill_Code = -1
Unload Me
End Sub

Private Sub Form_Load()
For Loop_1 = 0 To 255
    Hex_Code.AddItem Right("00" & Hex(Loop_1), 2), Loop_1
Next Loop_1
Hex_Code.Selected(0) = True
End Sub

Private Sub Hex_Code_Click()
If Auto_Chan Then Exit Sub
Auto_Chan = True
Char.Text = Valid_Char(Hex_Code.ListIndex)
Auto_Chan = False
End Sub

