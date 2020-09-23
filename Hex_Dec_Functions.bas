Attribute VB_Name = "Hex_Dec_Functions"
Option Explicit
' Hex_Dec_Funtions.bas....
' Code in this Module provided By JonnyPoet from CodeGuru..
' I may have tweeked or ajusted the code sligtly to suit the application..

Private Function AddDecadeStrings(sTxt1 As String, sTxt2 As String) As String
Dim a(100) As Integer, i As Integer, j As Integer, b(100) As Integer
Dim lg1 As Integer, lg2 As Integer, iLastPos As Integer
'read  sTxt1 into array
j = 0
lg1 = Len(sTxt1)
For i = lg1 To 1 Step -1 ' from right to left
    a(j) = Mid(sTxt1, i, 1)
    j = j + 1
Next
' read sText2 into array
j = 0
lg2 = Len(sTxt2)
For i = lg2 To 1 Step -1  ' from right to left
    b(j) = Mid(sTxt2, i, 1)
    j = j + 1
Next
iLastPos = IIf(lg1 > lg2, lg1, lg2)
AddDecadeStrings = AddArrays(iLastPos, a, b)
End Function

Private Function AddArrays(iLastPos As Integer, ByRef a() As Integer, ByRef b() As Integer) As String
Dim valR As Integer ' the digit
Dim Pos As Integer, r As String
Dim c As Integer ' carryflag
Dim lgR As Integer, i As Integer
' Values may have different length so
' iLastPos is last digit to add beginning from right to left

For Pos = 0 To iLastPos  'until possible last carryflag is added up
    valR = a(Pos) + b(Pos) + c
    c = 0
    If valR > 9 Then
        r = CStr(valR - 10) & r
        c = 1
    Else
        r = CStr(valR) & r
    End If
Next
' take away left values of 0
lgR = Len(r)
For i = 1 To lgR
    If Mid(r, i, 1) <> "0" Then Exit For
Next
i = i - 1
If i > 0 Then
    r = Right(r, lgR - i)
End If
AddArrays = r
End Function


Public Function HexToBin(sHex As String) As String
Dim sBin As String, sDig As String, i As Integer
For i = 1 To Len(sHex)
    Select Case Mid(sHex, i, 1)
        Case "1"
            sDig = "0001"
        Case "2"
            sDig = "0010"
        Case "3"
            sDig = "0011"
        Case "4"
            sDig = "0100"
        Case "5"
            sDig = "0101"
        Case "6"
            sDig = "0110"
        Case "7"
            sDig = "0111"
        Case "8"
            sDig = "1000"
        Case "9"
            sDig = "1001"
        Case "a", "A"
            sDig = "1010"
        Case "b", "B"
            sDig = "1011"
        Case "c", "C"
            sDig = "1100"
        Case "d", "D"
            sDig = "1101"
        Case "e", "E"
            sDig = "1110"
        Case "f", "F"
            sDig = "1111"
        Case "0"
            sDig = "0000"
    End Select
    sBin = sBin & sDig
Next
HexToBin = sBin
End Function

Private Function BinToDecade(sBin As String) As String
Dim lgBin As Integer, sDecade As String, j As Integer, sDigit As String
Dim i As Integer
lgBin = Len(sBin)
j = 0
For i = lgBin To 1 Step -1
    If Mid(sBin, i, 1) = 1 Then
        sDigit = ExpTwo(j)
    Else
        sDigit = "0"
    End If
    j = j + 1
    sDecade = AddDecadeStrings(sDecade, sDigit)
Next
BinToDecade = sDecade
End Function

Public Function HexToDecade(sHex As String) As String
Dim sBin As String, sDecade As String
sBin = HexToBin(sHex)
sDecade = BinToDecade(sBin)
HexToDecade = sDecade
End Function

Public Function DecadeToBinary(sVal As String) As String
'find out whats highest binary which is in that value
Dim sTwoEx As String, sPrevExTwo As String
Dim sBin As String, sCalc As String, ilg As Integer, sVgl As String
Dim iExp As Integer
' before doing anything check for sVal = "00...0"
ilg = Len(sVal)
sVgl = String(ilg, "0")
If CompareDecades(sVal, sVgl) = 0 Then
    DecadeToBinary = "0000"
    Exit Function
End If
' a real Value > 0
sCalc = sVal
iExp = 0
'Find Maximum Exponent
Do
    sTwoEx = ExpTwo(iExp)
    If CompareDecades(sCalc, sTwoEx) = -1 Then
        iExp = iExp - 1 ' one digit before 2-exp was less value
        Exit Do
    End If
    sPrevExTwo = sTwoEx ' store it before doing next one
    iExp = iExp + 1
Loop
sCalc = SubtractDecade(sCalc, sPrevExTwo)
sBin = sBin & "1"
iExp = iExp - 1
Do Until iExp < 0 ' Tweeked Here
    ' the last one before is smaller then sCalc
    sTwoEx = ExpTwo(iExp)
    If CompareDecades(sCalc, sTwoEx) = -1 Then
        sBin = sBin & "0"
    Else
        sCalc = SubtractDecade(sCalc, sTwoEx)
        sBin = sBin & "1"
    End If
    iExp = iExp - 1
Loop
DecadeToBinary = sBin
End Function


Public Function BinaryToHex(sBinVal As String) As String
Dim ilg As Integer, iDiff As Integer, iPos As Integer
Dim sPart As String, sHex As String, sFill As String

ilg = Len(sBinVal)
iDiff = ilg Mod 4
If iDiff > 0 Then ' Quick fix for extra 0 digit infront of Hex Val
    sFill = String(4 - iDiff, "0")
    sBinVal = sFill & sBinVal
End If
ilg = Len(sBinVal)
iPos = 1
Do
  sPart = Mid(sBinVal, iPos, 4)
  sHex = sHex & BinToHex(sPart)
  iPos = iPos + 4
Loop Until iPos > ilg  ' Tweeked Here
BinaryToHex = sHex
End Function

Private Function BinToHex(sBin As String) As String
Select Case sBin
    Case "0000"
        BinToHex = "0"
    Case "0001"
        BinToHex = "1"
    Case "0010"
        BinToHex = "2"
    Case "0011"
        BinToHex = "3"
    Case "0100"
        BinToHex = "4"
    Case "0101"
        BinToHex = "5"
    Case "0110"
        BinToHex = "6"
    Case "0111"
        BinToHex = "7"
    Case "1000"
        BinToHex = "8"
    Case "1001"
        BinToHex = "9"
    Case "1010"
        BinToHex = "A"  ' Changed to Uppercase Letters for Hex
    Case "1011"
        BinToHex = "B"
    Case "1100"
        BinToHex = "C"
    Case "1101"
        BinToHex = "D"
    Case "1110"
        BinToHex = "E"
    Case "1111"
        BinToHex = "F"
End Select
End Function

Public Function DecadeToHex(sDecade As String) As String
Dim y As String
y = DecadeToBinary(Trim(sDecade))
DecadeToHex = BinaryToHex(y)
End Function

Private Function CompareDecades(sVal1 As String, sVal2 As String) As Integer
Dim ilg1 As Integer, ilg2 As Integer, iDeltaLen As Integer, sFill As String
ilg1 = Len(sVal1)
ilg2 = Len(sVal2)
iDeltaLen = ilg1 - ilg2
CompareDecades = 0 'if equal
If iDeltaLen > 0 Then
    sFill = String(iDeltaLen, "0")
    sVal2 = sFill & sVal2
ElseIf iDeltaLen < 0 Then
    sFill = String(-iDeltaLen, "0")
    sVal1 = sFill & sVal1
End If
If sVal1 > sVal2 Then
    CompareDecades = 1 ' sVal1 > sVal2
ElseIf sVal1 < sVal2 Then
    CompareDecades = -1 'sVal2 > sVal1
End If
End Function

Private Function ExpTwo(Exp As Integer) As String
Dim iPos As Integer, sVal As String
sVal = 1
For iPos = 1 To Exp
    sVal = DoubleValue(sVal)
Next
ExpTwo = sVal
End Function

Private Function DoubleValue(sVal As String) As String
DoubleValue = AddDecadeStrings(sVal, sVal)
End Function

Private Function SubtractDecade(sVal As String, sSubtract As String) As String
' the substractor needs to be second value and less then sVal
Dim ilgV As Integer, ilgS As Integer, i As Integer, iDigVal As Integer, iDigSub As Integer
Dim sFill As String, iDelta As Integer, iFlag As Integer, iDiff As Integer, sBin As String

ilgV = Len(sVal)
ilgS = Len(sSubtract)
iDelta = ilgV - ilgS
If iDelta > 0 Then
    sFill = String(iDelta, "0")
    sSubtract = sFill & sSubtract
End If
' now we have same length
iFlag = 0
For i = ilgV To 1 Step -1
    iDigVal = Mid(sVal, i, 1)
    iDigSub = Mid(sSubtract, i, 1)
    iDiff = iDigVal - iDigSub - iFlag
    If iDiff >= 0 Then
        iFlag = 0
    Else
        iDiff = iDiff + 10
        iFlag = 1 ' value needs adding 10
    End If
    sBin = CStr(iDiff) & sBin
Next
SubtractDecade = sBin
End Function

