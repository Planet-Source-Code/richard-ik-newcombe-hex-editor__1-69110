Attribute VB_Name = "Hex_Functions"
Option Explicit

Public Function Hex_2_Byte(Hexadecimal As String) As Byte
Dim Tmp_Hex As String
Tmp_Hex = Right("00" & Trim(Hexadecimal), 2) ' split 2 digits for Byte...
Hex_2_Byte = CByte("&H" & Tmp_Hex)
End Function

Public Function Hex_2_Int(Hexadecimal As String) As Integer
Dim Tmp_Hex As Long
Tmp_Hex = Right("0000" & Hexadecimal, 4) ' split 4 digits for integer... (2 Bytes)
Hex_2_Int = CInt("&H" & Tmp_Hex)
End Function

Public Function Hex_2_Long(Hexadecimal As String) As Long
Dim Tmp_Hex As String
Tmp_Hex = Right("00000000" & Hexadecimal, 8) ' split 8 digits for Long... (4 Bytes)
Hex_2_Long = CLng("&H" & Tmp_Hex)
End Function

Public Function Hex_Check(Hexadecimal As String) As String
Dim Tmp_Char As String
Dim Tmp_Loop As Long
For Tmp_Loop = 1 To Len(Hexadecimal)
    Tmp_Char = Mid(Hexadecimal, Tmp_Loop, 1)
    Select Case Tmp_Char ' Filter and convert to uppercase digits..
        Case "0" To "9", "A" To "F", "a" To "f"
            Hex_Check = Hex_Check & UCase(Tmp_Char)
    End Select
Next Tmp_Loop
End Function

Public Function Valid_Char(Ascii As Byte) As String
Select Case Ascii And &HFF ' Valid codes are 0 to 255 (Hex : 00 to FF)
    Case 0 To 31
        Valid_Char = "."
    Case 32 To 127
        Valid_Char = Chr(Ascii)
    Case 128 To 144
        Valid_Char = "."
    Case 145 To 172
        Valid_Char = Chr(Ascii)
    Case 173 To 179
        Valid_Char = "."
    Case 180 To 255
        Valid_Char = Chr(Ascii)
End Select
End Function

Public Function Hex_C(ByVal In_data As Currency) As String
Hex_C = ""
next_Hex:
    Select Case (In_data And &HF)
     Case 0
        Hex_C = "0" & Hex_C
     Case 1
        Hex_C = "1" & Hex_C
     Case 2
        Hex_C = "2" & Hex_C
     Case 3
        Hex_C = "3" & Hex_C
     Case 4
        Hex_C = "4" & Hex_C
     Case 5
        Hex_C = "5" & Hex_C
     Case 6
        Hex_C = "6" & Hex_C
     Case 7
        Hex_C = "7" & Hex_C
     Case 8
        Hex_C = "8" & Hex_C
     Case 9
        Hex_C = "9" & Hex_C
     Case 10
        Hex_C = "A" & Hex_C
     Case 11
        Hex_C = "B" & Hex_C
     Case 12
        Hex_C = "C" & Hex_C
     Case 13
        Hex_C = "D" & Hex_C
     Case 14
        Hex_C = "E" & Hex_C
     Case 15
        Hex_C = "F" & Hex_C
    End Select
In_data = (In_data And &HFFFFFFF0) / &H10
    
If In_data <> 0 Then GoTo next_Hex
End Function

