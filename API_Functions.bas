Attribute VB_Name = "API_Functions"
Option Explicit

Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Public Sub API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency)
Dim FileH As Long
Dim Ret As Long
On Error Resume Next
FileH = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
If Err.Number > 0 Then
    Err.Clear
    FileNumber = -1
Else
    FileNumber = FileH
    Ret = SetFilePointer(FileH, 0, 0, FILE_BEGIN)
    API_FileSize FileH, FileSize
End If
On Error GoTo 0
End Sub
Public Sub API_FileSize(ByVal FileNumber As Long, ByRef FileSize As Currency)
    Dim FileSizeL As Long
    Dim FileSizeH As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(FileNumber, FileSizeH)
    Long2Size FileSizeL, FileSizeH, FileSize
End Sub

Public Sub API_ReadFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = ReadFile(FileNumber, Data(0), BlockSize, SizeRead, 0&)
BlockSize = SizeRead
End Sub

Public Sub API_CloseFile(ByVal FileNumber As Long)
Dim Ret As Long
Ret = CloseHandle(FileNumber)
End Sub

Public Sub API_WriteFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeWrit As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = WriteFile(FileNumber, Data(0), BlockSize, SizeWrit, 0&)
BlockSize = SizeWrit
End Sub

Private Sub Size2Long(ByVal FileSize As Currency, ByRef LongLow As Long, ByRef LongHigh As Long)
'&HFFFFFFFF unsigned = 4294967295
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
LongHigh = 0
Do Until FileSize < Cutoff
    LongHigh = LongHigh + 1
    FileSize = FileSize - Cutoff
Loop
If FileSize > 2147483647 Then
    LongLow = -CLng(Cutoff - (FileSize - 1))
Else
    LongLow = CLng(FileSize)
End If
End Sub

Private Sub Long2Size(ByVal LongLow As Long, ByVal LongHigh As Long, ByRef FileSize As Currency)
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1
FileSize = Cutoff * LongHigh
If LongLow < 0 Then
    FileSize = FileSize + (Cutoff + (LongLow + 1))
Else
    FileSize = FileSize + LongLow
End If
End Sub

Public Sub API_SetEndOfFile(ByVal FileNumber As Long, ByVal Position As Currency)
Dim PosL As Long
Dim PosH As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = SetEndOfFile(FileNumber)
End Sub
