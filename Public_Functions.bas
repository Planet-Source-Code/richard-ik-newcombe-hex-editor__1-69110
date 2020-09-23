Attribute VB_Name = "Public_Functions"
Option Explicit
Public Enum Logic_Types
    Logic_And
    logic_Nand
    Logic_Or
    Logic_Nor
    Logic_Xor
    Logic_XNor
    Logic_Not
End Enum

Public Enum File_Types
    exe
    ini
    avi
    Wave
    Word_doc
    access_db
    Gif
    mp3
    bmp
    tiff
    Zip
    HTML
    Rar
    arj
    Jpeg
    VB_Proj
    VB_Group
    VB_form
    VB_Mod
    bin
    txt
    None
End Enum

Public Enum Endian_Types ' Thanks to Pink98 from CodeGuru for his article on Data formats including Big/Little Endian format.
    Little_Endian
    Big_Endian
End Enum

Public Type Logic_Data
    HexC(3) As Byte
    Count As Byte
    Endian As Endian_Types
    None As Boolean
    Logics As Logic_Types
End Type
    
Public Type Search_Data
    HexC() As Byte
    Case As Boolean
    None As Boolean
End Type
    

Public Fill_Code As Integer
Public Jump_Loc As Currency
Public File_Len As Currency
Public File_sig As File_Types
Public Logic_Val As Logic_Data
Public Search_H(100) As Byte
Public Search_Len As Long
Public Search_Type As Long

Public Function GetFileType(xFile As String) As File_Types ' Thanks to Wizbang from Codeguru for finding this for me..
    On Error Resume Next
    Dim ID As String * 300
    If Dir$(xFile) = "" Then
        GetFileType = None
        Exit Function
    End If
    Open xFile For Binary Access Read As #1
    Get #1, 1, ID
    Close #1

    If Left(ID, 2) = "MZ" Or Left(ID, 2) = "ZM" Then
        GetFileType = exe
    ElseIf Left(ID, 1) = "[" And InStr(1, Left(ID, 100), "]") > 0 Then
        GetFileType = ini
    ElseIf Mid(ID, 9, 8) = "AVI LIST" Then
        GetFileType = avi
    ElseIf Left(ID, 4) = "RIFF" Then
        GetFileType = Wave
    ElseIf Left(ID, 4) = Chr(208) & Chr(207) & Chr(17) & Chr(224) Then
        GetFileType = Word_doc
    ElseIf Mid(ID, 5, 15) = "Standard Jet DB" Then
        GetFileType = access_db
    ElseIf Left(ID, 3) = "GIF" Or InStr(1, ID, "GIF89") > 0 Then
        GetFileType = Gif
    ElseIf Left(ID, 1) = Chr(255) And Mid(ID, 5, 1) = Chr(0) Then
        GetFileType = mp3
    ElseIf Left(ID, 2) = "BM" Then
        GetFileType = bmp
    ElseIf Left(ID, 3) = "II*" Then
        GetFileType = tiff
    ElseIf Left(ID, 2) = "PK" Then
        GetFileType = Zip
    ElseIf InStr(1, LCase(ID), "<html>") > 0 Or InStr(1, LCase(ID), "<!doctype") > 0 Then
        GetFileType = HTML
    ElseIf UCase(Left(ID, 3)) = "RAR" Then
        GetFileType = Rar
    ElseIf Left(ID, 2) = Chr(96) & Chr(234) Then
        GetFileType = arj
    ElseIf Left(ID, 3) = Chr(255) & Chr(216) & Chr(255) Then
        GetFileType = Jpeg
    ElseIf InStr(1, ID, "Type=") > 0 And InStr(1, ID, "Reference=") > 0 Then
        GetFileType = VB_Proj
    ElseIf Left(ID, 8) = "VBGROUP " Then
        GetFileType = VB_Group
    ElseIf Left(ID, 8) = "VERSION " And InStr(1, ID, vbCrLf & "Begin") > 0 Then
        GetFileType = VB_form
    ElseIf Left(ID, 9) = "Attribute" And InStr(1, ID, "VB_Name") > 0 Then
        GetFileType = VB_Mod
    ElseIf InStr(1, ID, Chr$(255)) > 0 Or InStr(1, ID, Chr$(1)) > 0 Or InStr(1, ID, Chr$(2)) > 0 Or InStr(1, ID, Chr$(3)) > 0 Then
        GetFileType = bin
    Else
        GetFileType = txt
    End If
End Function

Public Function F_Type_String(Index As File_Types) As String
Select Case Index
    Case exe
        F_Type_String = "Executable"
    Case ini
        F_Type_String = "Ini Settings"
    Case avi
        F_Type_String = "Audio Video"
    Case Wave
        F_Type_String = "Wave Audio"
    Case Word_doc
        F_Type_String = "MS Word Document"
    Case access_db
        F_Type_String = "MS Access Database"
    Case Gif
        F_Type_String = "Gif Image"
    Case mp3
        F_Type_String = "MP3 Audio"
    Case bmp
        F_Type_String = "BMP Image"
    Case tiff
        F_Type_String = "TIFF Image"
    Case Zip
        F_Type_String = "Zip Archive"
    Case HTML
        F_Type_String = "Web Document"
    Case Rar
        F_Type_String = "Rar Archive"
    Case arj
        F_Type_String = "ARJ Archive"
    Case Jpeg
        F_Type_String = "JPEG Image"
    Case VB_Proj
        F_Type_String = "VB Project"
    Case VB_Group
        F_Type_String = "VB Project Group"
    Case VB_form
        F_Type_String = "VB Form"
    Case VB_Mod
        F_Type_String = "VB Module"
    Case bin
        F_Type_String = "Data File"
    Case txt
        F_Type_String = "Text File"
End Select
End Function
