Attribute VB_Name = "R_T_B_Functions"
Option Explicit

Public Function Get_Actual(Index As Long, mode As Boolean) As Long
Dim Char_No As Long
Dim Line_No As Long
If mode Then
    Char_No = Index Mod 34
    Line_No = (Index - Char_No) / 34
    Get_Actual = (Line_No * 32) + Char_No
Else
    Char_No = Index Mod 18
    Line_No = (Index - Char_No) / 18
    Get_Actual = (Line_No * 16) + Char_No
End If
End Function

Public Function Set_Actual(Index As Long, mode As Boolean) As Long
Dim Char_No As Long
Dim Line_No As Long
If mode Then
    Char_No = Index Mod 32
    Line_No = (Index - Char_No) / 32
    Set_Actual = (Line_No * 34) + Char_No
Else
    Char_No = Index Mod 16
    Line_No = (Index - Char_No) / 16
    Set_Actual = (Line_No * 18) + Char_No
End If
If Set_Actual < 0 Then Set_Actual = 0
End Function

Public Function Get_H_Actual(Index As Long, mode As Boolean) As Long
Dim Hex_Pos As Long
Dim Char_No As Long
Dim Line_No As Long
If mode Then
    Hex_Pos = Index Mod 97
    Char_No = (Hex_Pos - (Hex_Pos Mod 3)) / 3
    Line_No = (Index - Hex_Pos) / 97
    Get_H_Actual = (Line_No * 32) + Char_No
Else
    Hex_Pos = Index Mod 49
    Char_No = (Hex_Pos - (Hex_Pos Mod 3)) / 3
    Line_No = (Index - Hex_Pos) / 49
    Get_H_Actual = (Line_No * 16) + Char_No
End If
End Function

Public Function Set_H_Actual(Index As Long, mode As Boolean) As Long
Dim Char_No As Long
Dim Line_No As Long
If mode Then
    Char_No = Index Mod 32
    Line_No = (Index - Char_No) / 32
    Set_H_Actual = (Line_No * 97) + (Char_No * 3)
Else
    Char_No = Index Mod 16
    Line_No = (Index - Char_No) / 16
    Set_H_Actual = (Line_No * 49) + (Char_No * 3)
End If
If Set_H_Actual < 0 Then Set_H_Actual = 0
End Function

Public Function Get_S_Actual(Index As Long) As Long
Dim Char_No As Long
Char_No = (Index - (Index Mod 3)) / 3
Get_S_Actual = Char_No

End Function

Public Function Set_S_Actual(Index As Long) As Long
Dim Char_No As Long
Char_No = Index
Set_S_Actual = (Char_No * 3)

End Function


