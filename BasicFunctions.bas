Attribute VB_Name = "BasicFunctions"
Option Compare Text
Option Explicit

' Following code executes a file with a known/registered extension
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Used to return a DOS 8.3 filename format
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
' Used to read INI files
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
' Used to see if a file is in use
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

' Used to wordwrap RTFs
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
' Returns system Temp folder
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const gintMAX_SIZE = 256

Public Function OpenThisFile(stFile As String, lShowHow As Long, sParams As String, LhWnd As Long) As Variant
' Simply attempts to open a file with the Shell command, if errors are encountered then...
'   > tries to open it with an API call using extensions to find associated executables, if error then...
'   > prompts user with the "Open With ..." routine
Dim lRet As Long, stRet As String, ErrID As Long

On Error GoTo TryAPIcall
    lRet = -1   ' set default value -- meaning failure
    If Len(sParams) > 0 Then sParams = " " & sParams    ' if no optional parameters, then format with a space
    lRet = Shell(stFile & sParams, lShowHow)       ' attempt simple shell command
OpenThisFile = lRet
Exit Function

TryAPIcall:
Err.Clear
' if above shell function failed, then try an association open based on the file extension
    ErrID = apiShellExecute(LhWnd, "OPEN", _
            stFile, sParams, App.Path, lShowHow)
    ' Errors will be a retruned value of <32
    If ErrID < 32& Then
        Select Case ErrID
            Case 31&:
                'Try the OpenWith dialog
                lRet = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                        & stFile, 1)
            Case 0&:
                stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
            Case 2&:
                stRet = "Error: File not found.  Couldn't Execute!"
            Case 3&:
                stRet = "Error: Path not found. Couldn't Execute!"
            Case 11&:
                stRet = "Error:  Bad File Format. Couldn't Execute!"
            Case Else:
        End Select
        If ErrID <> 31 Then
            lRet = -1 ' failure
            MsgBox stRet, vbExclamation + vbOKOnly  ' display error
        End If
        OpenThisFile = lRet
    Else
        lRet = 69
    End If
Resume Next
End Function

Public Function StripFile(Pathname As String, DPNEm As String) As String
Dim ChrsIn As String, ChrsOut As String, IdX As Integer, Chrs As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo StripFile_General_ErrTrap
If Pathname = "" Then Exit Function
ChrsIn = Pathname
Select Case InStr("DPNEm", DPNEm)
Case 1:     ' Return the Drive Letter
    GoSub ExtractDrive
Case 2:     ' Return the Path
    GoSub ExtractPath
Case 3:     ' Return the File Name
    GoSub ExtractName
Case 4:     ' Return the File Extension
    GoSub ExtractExtension
Case 5:     ' Return filename less the extension
    GoSub ExtractName
    ChrsIn = StripFile
    GoSub ExtractExtension
    StripFile = Left(ChrsIn, Chrs - 1)
End Select
Exit Function

ExtractDrive:
Chrs = InStr(ChrsIn, ":\") 'check to see if a forward slash exists
If Chrs Then 'if a forward slash is found in the passed string
   ChrsOut = Left(ChrsIn, Chrs + 1) 'get the drive
End If
StripFile = ChrsOut 'return the drive to the user
Return
 
ExtractExtension:
Chrs = InStr(ChrsIn, ".") 'check to see if a full stop exists
If Chrs Then 'if a full stop is found in the passed string
IdX = Chrs
Do While IdX > 0
    IdX = InStr(IdX + 1, ChrsIn, ".")
    If IdX Then Chrs = IdX
Loop
   ChrsOut = Mid(ChrsIn, Chrs + 1) 'get the extension
Else
    ChrsOut = ""
End If
StripFile = ChrsOut 'return the extension to the user
Return

ExtractName:
If InStr(ChrsIn, "\") Then 'check to see if a forward slash exists
   For IdX = Len(ChrsIn) To 1 Step -1 'step though until full name is extracted
       If Mid(ChrsIn, IdX, 1) = "\" Then
          ChrsOut = Mid(ChrsIn, IdX + 1)
          Exit For
       End If
   Next IdX
ElseIf InStr(ChrsIn, ":") = 2 Then 'otherwise, check to see if a colon exists
   ChrsOut = Mid(ChrsIn, 3)        'if so, return the filename
Else
   ChrsOut = ChrsIn 'otherwise, return the original string
End If
StripFile = ChrsOut 'return the filename to the user
Return

ExtractPath:
If InStr(ChrsIn, "\") Then 'check to see if a forward slash exists
   For IdX = Len(ChrsIn) To 1 Step -1 'step though until full name is extracted
       If Mid(ChrsIn, IdX, 1) = "\" Then
          ChrsOut = Left(ChrsIn, IdX)
          Exit For
       End If
   Next IdX
ElseIf InStr(ChrsIn, ":") = 2 Then 'otherwise, check to see if a colon exists
   ChrsOut = CurDir(ChrsIn)
   If Len(ChrsOut) = 0 Then
      ChrsOut = CurDir
   End If
Else
   ChrsOut = CurDir 'otherwise, return the current directory
End If
StripFile = ChrsOut 'return the filenames path to the user
Return
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
StripFile_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: StripFile" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Function ReadWriteINI(Mode As String, BKfile As String, tmpSecName As String, tmpKeyname As String, _
    Optional tmpKeyValue As String = "*****", _
    Optional DeleteSection As Boolean = False) As String

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo ReadWriteINI_General_ErrTrap

If DeleteSection = True Then
    'WritePrivateProfileSection tmpSecName, "", BKfile
    Exit Function
End If
Dim tmpString As String, tmpCounter As Integer
Dim secname As String, tmpTimer As Single
Dim KeyName As String
Dim keyvalue As String
Dim anInt
Dim defaultkey As String
On Error GoTo ReadWriteINIError

ReadWriteINI = tmpKeyValue
If IsNull(Mode) Or Len(Mode) = 0 Then Exit Function
If IsNull(tmpSecName) Or Len(tmpSecName) = 0 Then Exit Function
If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then Exit Function

secname = tmpSecName
KeyName = tmpKeyname
keyvalue = tmpKeyValue
defaultkey = tmpKeyValue
tmpString = tmpKeyValue
' ******* WRITE MODE *************************************
  If UCase(Mode) = "WRITE" Then
        If keyvalue = "" Then keyvalue = vbNullString
      anInt = WritePrivateProfileString(secname, KeyName, keyvalue, BKfile)
      If anInt > 0 Then anInt = 1
  Else
  ' *******  READ MODE *************************************
    If UCase(Mode) = "GET" Then
ReadFileNow:
      keyvalue = String$(255, 32)
      anInt = GetPrivateProfileString(secname, KeyName, defaultkey, keyvalue, Len(keyvalue), BKfile)
      If Left(keyvalue, Len(tmpKeyValue) + 1) <> tmpKeyValue & Chr$(0) Then     ' *** got it
         tmpString = keyvalue
         tmpString = RTrim(tmpString)
         If Len(tmpString) Then tmpString = Left(tmpString, Len(tmpString) - 1)
      End If
   End If
  End If
If anInt > 0 Then ReadWriteINI = tmpString
Exit Function
  ' *******
ReadWriteINIError:
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
ReadWriteINI_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: ReadWriteINI" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function
'-----------------------------------------------------------
' FUNCTION: FileInUse
' Determines whether the specified file is currently in use
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists and is in use, False otherwise
'-----------------------------------------------------------
'
Public Function FileInUse(ByVal strPathName As String) As Boolean
    Dim hFile As Long
    
 Const GENERIC_WRITE As Long = &H40000000
 Const OPEN_EXISTING As Long = 3
 Const FILE_ATTRIBUTE_NORMAL As Long = &H80
 Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
 Const INVALID_HANDLE_VALUE As Long = -1
 Const ERROR_SHARING_VIOLATION As Long = 32
    
    On Error Resume Next
    '
   strPathName = Trim(strPathName)
    '
    ' If the string is quoted, remove the quotes.
    '
    If Len(strPathName) = 0 Then Exit Function
    
    If Left$(strPathName, 1) = Chr$(34) And Right$(strPathName, 1) = Chr$(34) Then
        strPathName = Mid$(strPathName, 2, Len(strPathName) - 2)
    End If
     '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    hFile = CreateFile(strPathName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        FileInUse = Err.LastDllError = ERROR_SHARING_VIOLATION
    Else
        CloseHandle hFile
    End If
    Err.Clear
End Function

'-----------------------------------------------------------
' FUNCTION GetShortPathName
'
' Retrieve the short pathname version of a path possibly
'   containing long subdirectory and/or file names
'-----------------------------------------------------------
'
Public Function GetShortPathName(ByVal strLongPath As String) As String
    Const cchBuffer = 300
    Dim strShortPath As String
    Dim lResult As Long, nPos As Long

' Inserted by LaVolpe
On Error GoTo Function_GetShortPathName_General_ErrTrap_by_LaVolpe
    strShortPath = String$(cchBuffer, 0)
    lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
    If lResult = 0 Then
        'Just use the long name as this is usually good enough
        GetShortPathName = strLongPath
    Else

      nPos = InStr(strShortPath, vbNullChar)
      If nPos > 0 Then
          GetShortPathName = Left$(strShortPath, nPos - 1)
      Else
          GetShortPathName = strShortPath
      End If

    End If
Exit Function

Function_GetShortPathName_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function GetShortPathName]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Public Function GetTempFolder() As String
    Dim strBuf As String, nPos As Long

    strBuf = Space$(gintMAX_SIZE)
    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetTempPath(gintMAX_SIZE, strBuf) Then
      nPos = InStr(strBuf, vbNullChar)
      If nPos > 0 Then
          GetTempFolder = Left$(strBuf, nPos - 1)
      Else
          GetTempFolder = "C:\"
      End If

        If Right$(GetTempFolder, 1) <> "\" Then GetTempFolder = GetTempFolder & "\"
    End If
End Function

