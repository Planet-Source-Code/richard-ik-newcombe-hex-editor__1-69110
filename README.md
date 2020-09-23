<div align="center">

## Hex Editor


</div>

### Description

Ever needed to peek into a file and view the raw hex data. The older dos applications have limited capabilities to read in the newer OS's. The newer applications require registration, etc... Here is a full featured Hex Editor, that even blows away the 2gig file limit that VB6 has.

Several GUI formats are available to efectivly view your raw hex.. (16 or 32 Byte width, 32 or 48 lines).

Features :

* Full editing is alowable in hex or text.

* Files can be extended or shortend, by setting the EOF.

* Block Fill.

* Block Copy.

* Logic edits with 8, 16, 24 or 32 bit data (Big or Little Endian).

* Copy - Paste to and from the hex.

* Jump to Hex or Dec location.

* Hex or Text search.

* Verify edits defore writing.

Also accepts a filename as a command line parameter so that hex files can be opened with a single click..
 
### More Info
 
Commandline accepts filename

Uses alot of Windows API call's to Load, Save, edit the hexdata.. some of the methods used are for the 2gig file limit work around. and the limited ability of VB to work with data larger than 31 bits (32nd bit is always the sign)..



----

WARNING 

----

Files raw hex can be changed directly on disk, so be warned, once saved, edits cannot be undone..


<span>             |<span>
---                |---
**Submitted On**   |2007-07-20 19:36:00
**By**             |[Richard IK Newcombe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-ik-newcombe.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Hex\_Editor207856872007\.zip](https://github.com/Planet-Source-Code/richard-ik-newcombe-hex-editor__1-69110/archive/master.zip)

### API Declarations

```
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
```





