VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' Written by Ellis Dee
' Generic functions for accessing files in Windows
Option Explicit

' Constants for API calls
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200

Private Const MAX_PATH = 260
Private Const FILE_FLAGS = FOF_NOCONFIRMATION Or FOF_CREATEPROGRESSDLG Or FOF_NOCONFIRMMKDIR
Private Const FILE_FLAGS_SILENT = FOF_NOCONFIRMATION Or FOF_CREATEPROGRESSDLG Or FOF_NOCONFIRMMKDIR Or FOF_SILENT

' Enumerations
Public Enum WindowStateEnum
    wsNormal = 1
    wsMaximized = 3
    wsMinimized = 2
End Enum

' Type declarations for API calls
Private Type SHFILEOPSTRUCT
    Hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Long
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
    dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             '  e.g. VFT_DRIVER
    dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           '  e.g. 0
    dwFileDateLS As Long           '  e.g. 0
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Long
End Type

Private Type WIN32_FIND_DataOld
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' API
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DataOld) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathStripPath Lib "shlwapi" Alias "PathStripPathA" (ByVal pPath As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long



' METHODS


Public Sub Copy(ByVal Source As String, ByVal Dest As String, Optional pblnSilent As Boolean = False)
    Const FO_COPY As Long = &H2
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFile Source
    ValidateDest Dest
    With typFO
        .wFunc = FO_COPY
        If pblnSilent Then
            .fFlags = FILE_FLAGS_SILENT
        Else
            .fFlags = FILE_FLAGS
        End If
        .pFrom = Source & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    ' if lngreturn = 1223 then out of disk space
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub Delete(ByVal File As String, Optional ByVal RecycleBin As Boolean = False)
'    Const FO_DELETE As Long = &H3
'    Dim typFO As SHFILEOPSTRUCT
'    Dim lngReturn As Long
'
    ValidateFile File
'    With typFO
'        .wFunc = FO_DELETE
'        .fFlags = FILE_FLAGS
'        If RecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
'        .pFrom = File & vbNullChar & vbNullChar
'    End With
'    lngReturn = SHFileOperation(typFO)
'    ' If failed, Windows already informed the user as to how and why
'    If lngReturn <> 0 Then Err.Raise ErrorIgnore
    On Error Resume Next
    Kill File
    Select Case Err.Number
        Case 0, 70, 53
        Case Else: MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    End Select
End Sub

Public Sub DeleteInUse(ByVal File As String)
    Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
    
    MoveFileEx File, vbNull, MOVEFILE_DELAY_UNTIL_REBOOT
End Sub

Public Function Exists(ByVal File As String) As Boolean
    Exists = (PathFileExists(File) = 1)
    If Exists Then Exists = (PathIsDirectory(File) = 0)
End Function

Public Function FormatSize(ByVal Size As Currency) As String
    Const Kilobyte As Currency = 1024@
    Const HundredK As Currency = 102400@
    Const ThousandK As Currency = 1024000@
    Const Megabyte As Currency = 1048576@
    Const HundredMeg As Currency = 104857600@
    Const ThousandMeg As Currency = 1048576000@
    Const Gigabyte As Currency = 1073741824@
    Const Terabyte As Currency = 1099511627776@
    
    If Size < Kilobyte Then
        FormatSize = Int(Size) & " bytes"
    ElseIf Size < HundredK Then
        FormatSize = Format(Size / Kilobyte, "0.0") & " KB"
    ElseIf Size < ThousandK Then
        FormatSize = Int(Size / Kilobyte) & " KB"
    ElseIf Size < HundredMeg Then
        FormatSize = Format(Size / Megabyte, "0.0") & " MB"
    ElseIf Size < ThousandMeg Then
        FormatSize = Int(Size / Megabyte) & " MB"
    ElseIf Size < Terabyte Then
        FormatSize = Format(Size / Gigabyte, "0.00") & " GB"
    Else
        FormatSize = Format(Size / Terabyte, "0.00") & " TB"
    End If
End Function

Public Sub GetDatetime(ByVal File As String, ByRef pdtmCreated As Date, ByRef pdtmModified As Date, ByRef pdtmAccess As Date)
    Const GENERIC_READ = &H80000000
    Const GENERIC_WRITE = &H40000000
    Const FILE_SHARE_READ = &H1
    Const FILE_SHARE_DELETE As Long = &H4
    Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    Const OPEN_EXISTING = 3
    Dim typCreate As FILETIME
    Dim typAccess As FILETIME
    Dim typModified As FILETIME
    Dim lngHandle As Long
    
    If Me.Exists(File) Then
        lngHandle = CreateFile(File, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
        If GetFileTime(lngHandle, typCreate, typAccess, typModified) = 1 Then
            pdtmCreated = StructureToDate(typCreate)
            pdtmAccess = StructureToDate(typAccess)
            pdtmModified = StructureToDate(typModified)
        End If
        CloseHandle lngHandle
    End If
End Sub

Public Function DateModified(ByVal File As String) As Date
    Dim dtmCreated As Date
    Dim dtmModified As Date
    Dim dtmAccess As Date
    
    GetDatetime File, dtmCreated, dtmModified, dtmAccess
    DateModified = dtmModified
End Function

Public Function GetLongName(ByVal File As String, Optional ByVal IncludePath As Boolean = True) As String
    Dim typ As WIN32_FIND_DataOld
    Dim strToken As String
    Dim lngHandle As Long
    Dim strReturn As String
    Dim lngTemp As Long
    
    ValidateFile File
    Do While lngHandle <> -1
        lngHandle = FindFirstFile(File, typ)
        strToken = Left$(typ.cFileName, InStr(typ.cFileName, vbNullChar) - 1)
        If Len(File) > 2 Then
            strReturn = strToken & "\" & strReturn
            lngTemp = InStrRev(File, "\")
            If lngTemp > 0 Then File = Left$(File, lngTemp - 1)
        Else
            strReturn = File & "\" & strReturn
            Exit Do
        End If
        lngTemp = FindClose(lngHandle)
    Loop
    lngTemp = FindClose(lngHandle)
    'Strip away unwanted characters.
    strReturn = Left$(strReturn, Len(strReturn) - 1)
    If Not IncludePath Then
        Call PathStripPath(strReturn)
        strReturn = Left$(strReturn, InStr(strReturn, vbNullChar) - 1)
    End If
    GetLongName = strReturn
End Function

Public Function GetShortName(ByVal File As String, Optional ByVal IncludePath As Boolean = True) As String
    Dim strReturn As String
    Dim intLen As Integer
    Dim lngReturn As Long
    
    ValidateFile File
    'Set up buffer area for API function call return
    strReturn = Space$(255)
    intLen = Len(strReturn)
    'Call the function
    lngReturn = GetShortPathName(File, strReturn, intLen)
    'Strip away unwanted characters.
    strReturn = Left$(strReturn, lngReturn)
    If Not IncludePath Then
        Call PathStripPath(strReturn)
        strReturn = Left$(strReturn, InStr(strReturn, vbNullChar) - 1)
    End If
    GetShortName = strReturn
End Function

' Currency handles file sizes over 2 gig
' Thanks to Richard Newcombe from codeguru.com
Public Function GetSize(ByVal File As String) As Currency
    Const GENERIC_READ = &H80000000
    Const FILE_SHARE_READ = &H1
    Const FILE_SHARE_WRITE = &H2
    Const OPEN_EXISTING = 3
    Dim lngHandle As Long
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim curFileSize As Currency
    
    ' Open the file
    lngHandle = CreateFile(File, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
    ' Get the file size
    lngLow = GetFileSize(lngHandle, lngHigh)
    CloseHandle lngHandle
    ' Combine the Low and High values into one currency
    ' Must use the '@' currency declaration or IDE will balk
    curFileSize = 4294967295@ * lngHigh
    If lngLow < 0 Then
        curFileSize = curFileSize + (4294967295@ + (lngLow + 1))
    Else
        curFileSize = curFileSize + lngLow
    End If
    GetSize = curFileSize
End Function

Public Function GetType(ByVal File As String) As String
    Dim lngPos As Long
    Dim strExt As String
    Dim strKey As String
    
    lngPos = InStrRev(File, ".")
    If lngPos = 0 Then
        strExt = "." & File
    Else
        strExt = Mid$(File, lngPos)
    End If
    strKey = RegistryRead(strExt)
    If Len(strKey) <> 0 Then GetType = RegistryRead(strKey)
    If Len(GetType) = 0 Then GetType = "<unknown>"
End Function

Public Function GetVersion(ByVal File As String) As String
    Dim lngSize As Long
    Dim lngLen As Long
    Dim lngHandle As Long
    Dim lngBuffer As Long
    Dim lngRes As Long
    Dim lngMajor As Long
    Dim lngMinor As Long
    Dim lngRevision As Long
    Dim typ As VS_FIXEDFILEINFO
    Dim bytArray() As Byte
    
    GetVersion = ""
    lngSize = GetFileVersionInfoSize(File, lngHandle)
    ReDim bytArray(lngSize)
    lngRes = GetFileVersionInfo(File, lngHandle, lngSize, bytArray(0))
    If lngRes <> 0 Then
        lngRes = VerQueryValue(bytArray(0), "\", lngBuffer, lngLen)
        If lngRes <> 0 Then
            ' Copy the fixed file info into the structure
            CopyMemory typ, ByVal lngBuffer, Len(typ)
            lngMajor = typ.dwFileVersionMS \ &HFFFF&
            lngMinor = typ.dwFileVersionMS And &HFFFF&
            lngRevision = typ.dwFileVersionLS And &HFFFF&
            GetVersion = lngMajor & "." & lngMinor
            If lngRevision <> 0 Then GetVersion = GetVersion & "." & lngRevision
        End If
    End If
End Function

Public Function LoadToString(File As String) As String
    Dim strReturn As String
    Dim FileNumber As Long

    FileNumber = FreeFile
    Open File For Binary Access Read As #FileNumber
    If LOF(FileNumber) > 0 Then
        strReturn = Space(LOF(FileNumber))
        Get #FileNumber, , strReturn
    End If
    Close #FileNumber
    Do While Right$(strReturn, 2) = vbNewLine
        strReturn = Left$(strReturn, Len(strReturn) - 2)
    Loop
    LoadToString = strReturn
End Function

Public Function MakeNameDOS(ByVal File As String) As String
    File = Replace(File, ":", "-")
    File = Replace(File, "\", "")
    File = Replace(File, "/", "")
    File = Replace(File, "*", "")
    File = Replace(File, "?", "")
    File = Replace(File, "+", "")
    File = Replace(File, Chr(34), "'")
    File = Replace(File, "<", "(")
    File = Replace(File, ">", ")")
    File = Replace(File, "|", "")
    MakeNameDOS = File
End Function

Public Function MakeNameUnique(ByVal File As String) As String
    Dim strBase As String
    Dim strFile As String
    Dim strExt As String
    Dim bytChar() As Byte
    Dim lngLen As Long
    Dim blnNewChar As Boolean
    Dim i As Long
    
    If Not Exists(File) Then
        MakeNameUnique = File
        Exit Function
    End If
    strFile = File
    i = InStrRev(strFile, ".")
    If i = 0 Then
        strBase = strFile & "_"
    Else
        strExt = Mid$(strFile, i)
        strBase = Left$(strFile, i - 1) & "_"
    End If
    ReDim bytChar(0)
    bytChar(0) = 96
    Do While Exists(strFile)
        blnNewChar = True
        For i = lngLen To 0 Step -1
            If bytChar(i) < 122 Then
                bytChar(i) = bytChar(i) + 1
                blnNewChar = False
                Exit For
            Else
                bytChar(i) = 97
            End If
        Next
        If blnNewChar Then
            lngLen = lngLen + 1
            ReDim Preserve bytChar(lngLen)
            bytChar(lngLen) = 97
        End If
        strFile = strBase & StrConv(bytChar, vbUnicode) & strExt
    Loop
    Erase bytChar
    MakeNameUnique = strFile
End Function

Public Sub Move(ByVal Source As String, ByVal Dest As String)
    On Error Resume Next
    Name Source As Dest
    If Err.Number Then MsgBox Err.Description, vbInformation, "Error " & Err.Number
    Exit Sub

    Const FO_MOVE As Long = &H1
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFile Source
    ValidateDest Dest
    With typFO
        .wFunc = FO_MOVE
        .fFlags = FILE_FLAGS
        .pFrom = Source & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
    End With
    
'    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub Rename(ByVal Source As String, ByVal Dest As String)
    Const FO_RENAME As Long = &H4
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFile Source
    With typFO
        .wFunc = FO_RENAME
        .fFlags = FILE_FLAGS
        .pFrom = Source & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Function Run(ByVal File As String, Optional ByVal WindowState As WindowStateEnum = wsNormal, Optional ByVal DefaultFolder As String) As Long
    Dim lngPos As Long
    Dim lngDesktop As Long
    
    If DefaultFolder = "" Then
        lngPos = InStrRev(File, "\")
        If lngPos > 0 Then DefaultFolder = Left$(File, lngPos - 1)
    End If
    lngDesktop = GetDesktopWindow()
    Run = ShellExecute(lngDesktop, "Open", File, "", DefaultFolder, WindowState) 'SW_SHOW)
End Function

Public Function RunParams(ByVal File As String, ByVal Params As String, Optional ByVal WindowState As WindowStateEnum = wsNormal, Optional ByVal DefaultFolder As String) As Long
    Dim lngPos As Long
    Dim lngDesktop As Long
    
    If DefaultFolder = "" Then
        lngPos = InStrRev(File, "\")
        If lngPos > 0 Then DefaultFolder = Left$(File, lngPos - 1)
    End If
    lngDesktop = GetDesktopWindow()
    RunParams = ShellExecute(lngDesktop, "Open", File, Params, DefaultFolder, WindowState) 'SW_SHOW)
End Function

Public Sub RunAs(ByVal File As String, ByVal AsExt As String, Optional ByVal DefaultFolder As String)
    Dim lngPos As Long
    Dim strExt As String
    Dim strKey As String
    Dim strCommand As String

    If DefaultFolder = "" Then
        lngPos = InStrRev(File, "\")
        If lngPos > 0 Then DefaultFolder = Left$(File, lngPos - 1)
    End If
    lngPos = InStrRev(AsExt, ".")
    If lngPos = 0 Then
        AsExt = "." & AsExt
    Else
        AsExt = Mid$(AsExt, lngPos)
    End If
    strKey = RegistryRead(AsExt)
    If Len(strKey) <> 0 Then strCommand = RegistryRead(strKey & "\shell\open\command")
    If Len(strCommand) <> 0 Then
        If InStr(strCommand, """%1""") <> 0 Then
            strCommand = Replace$(strCommand, "%1", File)
        ElseIf InStr(strCommand, "%1") Then
            strCommand = Replace$(strCommand, "%1", """" & File & """")
        ElseIf InStr(strCommand, """%L""") <> 0 Then
            strCommand = Replace$(strCommand, "%L", File)
        ElseIf InStr(strCommand, "%L") Then
            strCommand = Replace$(strCommand, "%L", """" & File & """")
        Else
            strCommand = strCommand & " """ & File & """"
        End If
        Shell strCommand, vbNormalFocus
    End If
End Sub

Public Function SaveStringAs(File As String, Text As String) As Boolean
On Error GoTo SaveStringAsErr
    Dim FileNumber As Long

    FileNumber = FreeFile
SaveStringAsRetry:
    Open File For Output As #FileNumber
    If Right$(Text, 4) = vbNewLine & vbNewLine Then
        Print #FileNumber, Left$(Text, Len(Text) - 2)
    Else
        Print #FileNumber, Text
    End If
    SaveStringAs = True
    
SaveStringAsExit:
    Close #FileNumber
    Exit Function
    
SaveStringAsErr:
    Select Case MsgBox(Err.Description, vbRetryCancel + vbInformation, "Notice")
        Case vbRetry
            Resume SaveStringAsRetry
        Case vbCancel
            Resume SaveStringAsExit
    End Select
End Function



' INTERNAL FUNCTIONS


Private Function StructureToDate(ptyp As FILETIME) As Date
    Dim typLocal As FILETIME
    Dim typSystem As SYSTEMTIME
    
    If FileTimeToLocalFileTime(ptyp, typLocal) = 1 Then
        If FileTimeToSystemTime(typLocal, typSystem) = 1 Then
            With typSystem
                StructureToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
            End With
        End If
    End If
End Function

' Thanks to schoolbusdriver from vbforums.com
Private Function RegistryRead(ByVal strKey As String) As String
    Const HKEY_CLASSES_ROOT = &H80000000
    Const strKey_QUERY_VALUE = &H1
    Const ERROR_MORE_DataOld = 234 '  dderror
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1
    Const REG_EXPAND_SZ = 2
    Const REG_MULTI_SZ = 7
    Dim lngHandle As Long
    Dim strValue As String
    Dim lngType As Long
    Dim strData As String
    Dim strExpanded As String
    Dim lngLen As Long
    
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, strKey, 0&, strKey_QUERY_VALUE, lngHandle) = ERROR_SUCCESS Then
        If RegQueryValueEx(lngHandle, strValue, 0&, lngType, 0&, lngLen) = ERROR_MORE_DataOld Then
            strData = Space$(lngLen)
            Select Case lngType
                Case REG_SZ
                    If RegQueryValueEx(lngHandle, strValue, 0&, 0&, ByVal strData, lngLen) = ERROR_SUCCESS Then RegistryRead = Left$(strData, lngLen - 1)
                Case REG_EXPAND_SZ
                    If RegQueryValueEx(lngHandle, strValue, 0&, 0&, ByVal strData, lngLen) = ERROR_SUCCESS Then
                        strData = Left$(strData, lngLen - 1)
                        lngLen = lngLen + 1023
                        strExpanded = Space$(lngLen)
                        lngLen = ExpandEnvironmentStrings(strData, strExpanded, lngLen)
                        RegistryRead = Left$(strExpanded, lngLen - 1)
                    End If
                Case REG_MULTI_SZ
                    If RegQueryValueEx(lngHandle, strValue, 0&, 0&, ByVal strData, lngLen) = ERROR_SUCCESS Then RegistryRead = Left$(strData, lngLen - 2)
            End Select
        End If
        RegCloseKey lngHandle
    End If
End Function

Private Sub ValidateDest(ByVal pstrFile As String)
    Dim blnFolder As Boolean
    Dim lngError As Long
    
    ' Allow a retry for removable disks
    Do While True
        ' Check if file doesn't exist or it's actually a folder
        On Error Resume Next
        blnFolder = GetAttr(pstrFile) And vbDirectory
        lngError = Err.Number
        On Error GoTo 0
        Select Case lngError
            Case 0, 53: Exit Do
            Case 5 ' Drive not ready
                If MsgBox(Left$(pstrFile, 2) & " Drive not ready", vbInformation + vbRetryCancel, "Disk Error") = vbCancel Then
                    Err.Raise 71, , "Disk not ready"
                End If
            Case 52 ' Bad command or file name
                Err.Raise 52, , "'" & pstrFile & "' is not a valid file name."
            Case 76 ' Invalid path
                If InStrRev(pstrFile, "\") > 0 Then pstrFile = Left$(pstrFile, InStrRev(pstrFile, "\") - 1)
                Err.Raise 76, , "'" & pstrFile & "' not a valid path"
            Case Else: Err.Raise lngError
        End Select
    Loop
End Sub

Private Sub ValidateFile(ByVal pstrFile As String)
    Dim blnFolder As Boolean
    Dim lngError As Long
    
    ' Allow a retry for removable disks
    Do While True
        ' Check if file doesn't exist or it's actually a folder
        On Error Resume Next
        blnFolder = GetAttr(pstrFile) And vbDirectory
        lngError = Err.Number
        On Error GoTo 0
        Select Case lngError
            Case 0: Exit Do
            Case 5 ' Drive not ready
                If MsgBox(Left$(pstrFile, 2) & " Drive not ready", vbInformation + vbRetryCancel, "Disk Error") = vbCancel Then
                    Err.Raise 71, , "Disk not ready"
                End If
            Case 52 ' Bad command or file name
                Err.Raise 52, , "'" & pstrFile & "' is not a valid file name."
            Case 53 ' File not found
                Err.Raise 53, , "'" & pstrFile & "' not found"
            Case 76 ' Invalid path
                If InStrRev(pstrFile, "\") > 0 Then pstrFile = Left$(pstrFile, InStrRev(pstrFile, "\") - 1)
                Err.Raise 76, , "'" & pstrFile & "' not a valid path"
            Case Else: Err.Raise lngError
        End Select
    Loop
End Sub
