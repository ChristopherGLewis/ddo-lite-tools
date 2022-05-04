VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' Written by Ellis Dee
' Generic functions for accessing folders in Windows
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
Public Enum FolderEnum
    feApp = 0  ' \Program Files\Project (more reliable than App.Path)
    feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application DataOld\Microsoft\CD Burning
    feCommonAppData = 35 ' \Docs & Settings\All Users\Application DataOld
    feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
    feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
    feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
    feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
    feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
    feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
    feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
    feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
    feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
    feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application DataOld
    feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application DataOld\Microsoft\CD Burning
    feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
    feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feProgramFiles = 38 ' \Program Files
    feProgramFilesCommon = 43 ' \Program Files\Common Files
    'feRecycleBin = 10 ' ???
    feUser = 40 ' \Docs & Settings\User
    feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
    feUserAppData = 26 ' \Docs & Settings\User\Application DataOld
    feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feUserCookies = 33 ' \Docs & Settings\User\Cookies
    feUserDesktop = 16 ' \Docs & Settings\User\Desktop
    feUserDocs = 5 ' \Docs & Settings\User\My Documents
    feUserFavorites = 6 ' \Docs & Settings\User\Favorites
    feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
    feUserNetHood = 19 ' \Docs & Settings\User\NetHood
    feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
    feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
    feUserRecent = 8 ' \Docs & Settings\User\Recent
    feUserSendTo = 9 ' \Docs & Settings\User\SendTo
    feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
    feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
    feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
    feUserTemplates = 21 ' \Docs & Settings\User\Templates
    feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
    feWindows = 36 ' \Windows
    feWindowFonts = 20 ' \Windows\Fonts
    feWindowsResources = 56 ' \Windows\Resources
    feWindowsSystem = 37 ' \Windows\System32
End Enum

Public Enum FolderDateEnum
    peCreated
    peModified
    peAccessed
End Enum

' Type declarations for API calls
Private Type BrowseFolderType
   hwnd As Long
   Root As Long
   DisplayName As Long
   Title As Long
   flags As Long
   Callback As Long
   Param As Long
   Image As Long
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Long
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
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
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DataOld) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DataOld) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseFolderType) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


' METHODS


Public Function Abbreviate(ByVal Folder As String) As String
    Dim strUserDocs As String
    Dim strSharedDocs As String
    Dim strUserAppData As String
    Dim strLocalAppData As String
    Dim strSharedAppData As String
    
    If Len(Folder) Then
        strUserDocs = SpecialFolder(feUserDocs)
        strSharedDocs = SpecialFolder(feCommonDocs)
        strUserAppData = SpecialFolder(feUserAppData)
        strLocalAppData = SpecialFolder(feLocalAppData)
        strSharedAppData = SpecialFolder(feCommonAppData)
        If strUserDocs = Left$(Folder, Len(strUserDocs)) Then
            Abbreviate = "User\Docs" & Mid$(Folder, Len(strUserDocs) + 1)
        ElseIf strSharedDocs = Left$(Folder, Len(strSharedDocs)) Then
            Abbreviate = "Shared\Docs" & Mid$(Folder, Len(strSharedDocs) + 1)
        ElseIf strUserAppData = Left$(Folder, Len(strUserAppData)) Then
            Abbreviate = "User\AppData" & Mid$(Folder, Len(strUserAppData) + 1)
        ElseIf strLocalAppData = Left$(Folder, Len(strLocalAppData)) Then
            Abbreviate = "User\Local AppData" & Mid$(Folder, Len(strLocalAppData) + 1)
        ElseIf strSharedAppData = Left$(Folder, Len(strSharedAppData)) Then
            Abbreviate = "Shared\AppData" & Mid$(Folder, Len(strSharedAppData) + 1)
        Else
            Abbreviate = Folder
        End If
    End If
End Function

Public Function Browse(ByVal Title As String, Optional ByVal Root As FolderEnum = 0) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const BIF_DONTGOBELOWDOMAIN = 2
    Dim lngHandle As Long
    Dim strBuffer As String
    Dim typ As BrowseFolderType
    
    With typ
        .hwnd = Screen.ActiveForm.hwnd
        .Title = lstrcat(Title, "")
        .Root = Root
        .flags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
    End With
    lngHandle = SHBrowseForFolder(typ)
    If (lngHandle) Then
        strBuffer = Space$(MAX_PATH)
        If SHGetPathFromIDList(lngHandle, strBuffer) Then strBuffer = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If
    Browse = strBuffer
End Function

Public Sub Copy(ByVal Source As String, ByVal Dest As String)
    Const FO_COPY As Long = &H2
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFolder Source
    ValidateDest Dest
    With typFO
        .wFunc = FO_COPY
        .fFlags = FILE_FLAGS
        .pFrom = Source & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub Create(ByVal Folder As String)
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    MakeSureDirectoryPathExists Folder
End Sub

Public Sub Delete(ByVal Folder As String, Optional ByVal RecycleBin As Boolean = False, Optional pblnSilent As Boolean = False)
    Const FO_DELETE As Long = &H3
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFolder Folder
    With typFO
        .wFunc = FO_DELETE
        If pblnSilent Then
            .fFlags = FILE_FLAGS_SILENT
        Else
            .fFlags = FILE_FLAGS
        End If
        If RecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        .pFrom = Folder & vbNullChar & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub DeleteMultiple(Folder() As String, Optional ByVal RecycleBin As Boolean = False, Optional pblnSilent As Boolean = False)
    Const FO_DELETE As Long = &H3
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    Dim strFolder As String
    Dim i As Long
    
    For i = LBound(Folder) To UBound(Folder)
        ValidateFolder Folder(i)
        strFolder = strFolder & Folder(i) & vbNullChar
    Next
    With typFO
        .wFunc = FO_DELETE
        If pblnSilent Then
            .fFlags = FILE_FLAGS_SILENT
        Else
            .fFlags = FILE_FLAGS
        End If
        If RecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        .pFrom = strFolder & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub EnumerateFoldersToDebugWindow()
    Dim i As Long
    
    For i = 0 To 255
        If Len(SpecialFolder(i)) Then
            Debug.Print i & ": " & SpecialFolder(i)
        End If
    Next
End Sub

Public Function Exists(ByVal Folder As String) As Boolean
    Exists = (PathFileExists(Folder) = 1)
    If Exists Then Exists = (PathIsDirectory(Folder) <> 0)
End Function

Public Function Expand(ByVal Folder As String) As String
    Dim strMyDocs As String
    Dim strSharedDocs As String
    
    If Left$(Folder, 9) = "User\Docs" Then
        Expand = Me.UserDocs & Mid$(Folder, 10)
    ElseIf Left$(Folder, 11) = "Shared\Docs" Then
        Expand = Me.CommonDocs & Mid$(Folder, 12)
    ElseIf Left$(Folder, 12) = "User\AppData" Then
        Expand = Me.UserAppData & Mid$(Folder, 13)
    ElseIf Left$(Folder, 14) = "Shared\AppData" Then
        Expand = Me.CommonAppData & Mid$(Folder, 15)
    Else
        Expand = Folder
    End If
End Function

Public Function IsRelative(ByVal Folder As String) As Boolean
    IsRelative = (PathIsRelative(Folder) = 1)
End Function

Public Function RelativeToAbsolute(ByVal Folder As String) As String
    Dim strBuffer As String
    
    strBuffer = Space$(MAX_PATH)
    If PathCanonicalize(strBuffer, Folder) Then
        RelativeToAbsolute = Left$(strBuffer, InStr(strBuffer & vbNullChar, vbNullChar) - 1)
    Else
        RelativeToAbsolute = Folder
    End If
End Function

Public Function Explore(ByVal Folder As String, Optional ByVal FolderPane As Boolean, Optional ByVal Root As Boolean, Optional ByVal WindowState As WindowStateEnum) As Long
    Dim strCommand As String
    
    strCommand = "Explorer "
    If FolderPane Then strCommand = strCommand & "/e,"
    If Root Then strCommand = strCommand & "/root,"
    strCommand = strCommand & Chr(34) & Folder & Chr(34)
    Shell strCommand, WindowState
End Function

Public Sub GetDatetime(ByVal Folder As String, ByRef pdtmCreated As Date, ByRef pdtmModified As Date, ByRef pdtmAccess As Date)
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
    
    If Me.Exists(Folder) Then
        lngHandle = CreateFile(Folder, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
        If GetFileTime(lngHandle, typCreate, typAccess, typModified) = 1 Then
            pdtmCreated = StructureToDate(typCreate)
            pdtmAccess = StructureToDate(typAccess)
            pdtmModified = StructureToDate(typModified)
        End If
        CloseHandle lngHandle
    End If
End Sub

Public Function GetLongName(ByVal Folder As String) As String
    Dim strToken As String
    Dim lngHandle As Long
    Dim strReturn As String
    Dim lngTemp As Long
    Dim typ As WIN32_FIND_DataOld

    If Right$(Folder, 1) = "\" Then Folder = Left$(Folder, Len(Folder) - 1)
    ValidateFolder Folder
    Do While lngHandle <> -1
        lngHandle = FindFirstFile(Folder, typ)
        strToken = Left$(typ.cFileName, InStr(typ.cFileName, vbNullChar) - 1)
        If Len(Folder) > 2 Then
            strReturn = strToken & "\" & strReturn
            lngTemp = InStrRev(Folder, "\")
            If lngTemp > 0 Then Folder = Left$(Folder, lngTemp - 1)
        Else
            strReturn = Folder & "\" & strReturn
            Exit Do
        End If
        lngTemp = FindClose(lngHandle)
    Loop
    lngTemp = FindClose(lngHandle)
    'Strip away unwanted characters.
    GetLongName = Left$(strReturn, Len(strReturn) - 1)
End Function

Public Function GetShortName(ByVal Folder As String) As String
    Dim strReturn As String
    Dim intLen As Integer
    Dim lngReturn As Long
    
    ValidateFolder Folder
    'Set up buffer area for API function call return
    strReturn = Space$(255)
    intLen = Len(strReturn)
    
    'Call the function
    lngReturn = GetShortPathName(Folder, strReturn, intLen)
    'Strip away unwanted characters.
    GetShortName = Left$(strReturn, lngReturn)
End Function

Public Function HasSubFolders(ByVal Folder As String) As Boolean
    Const INVALID_HANDLE_VALUE = -1
    Dim typ As WIN32_FIND_DataOld
    Dim lngFile As Long
    
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    lngFile = FindFirstFile(Folder & "*.*", typ)
    If lngFile <> INVALID_HANDLE_VALUE Then
        Do
            If (typ.dwFileAttributes And vbDirectory) Then
                If (Left$(typ.cFileName, 1) <> ".") Then
                    HasSubFolders = True
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lngFile, typ)
    End If
    Call FindClose(lngFile)
End Function

Public Function IsEmpty(ByVal Folder As String) As Boolean
    IsEmpty = (PathIsDirectoryEmpty(Folder) = 1)
End Function

Public Function MakeNameUnique(ByVal Folder As String) As String
    Dim strBase As String
    Dim strFolder As String
    Dim bytChar() As Byte
    Dim lngLen As Long
    Dim blnNewChar As Boolean
    Dim i As Long

    strFolder = Folder
    strBase = strFolder & "_"
    ReDim bytChar(0)
    bytChar(0) = 96
    Do While Exists(strFolder)
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
        strFolder = strBase & StrConv(bytChar, vbUnicode)
    Loop
    Erase bytChar
    MakeNameUnique = strFolder
End Function

Public Sub Move(ByVal Source As String, ByVal Dest As String)
    Const FO_MOVE As Long = &H1
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFolder Source
    ValidateDest Dest
    With typFO
        .wFunc = FO_MOVE
        .fFlags = FILE_FLAGS
        .pFrom = Source & vbNullChar & vbNullChar
        .pTo = Dest & vbNullChar & vbNullChar
    End With
    lngReturn = SHFileOperation(typFO)
    ' If failed, Windows already informed the user as to how and why
    If lngReturn <> 0 Then Err.Raise ErrorIgnore
End Sub

Public Sub Rename(ByVal Source As String, ByVal Dest As String)
    Const FO_RENAME As Long = &H4
    Dim typFO As SHFILEOPSTRUCT
    Dim lngReturn As Long
    
    ValidateFolder Source
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

Public Sub SetDatetime(ByVal Folder As String, ByVal pdtmCreated As Date, ByVal pdtmModified As Date, ByVal pdtmAccess As Date)
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
    
    If Me.Exists(Folder) Then
        DateToStructure pdtmCreated, typCreate
        DateToStructure pdtmModified, typModified
        DateToStructure pdtmAccess, typAccess
        lngHandle = CreateFile(Folder, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
        SetFileTime lngHandle, typCreate, typAccess, typModified
        CloseHandle lngHandle
    End If
End Sub



' PROPERTIES


' Application
Public Property Get Application() As String
    Application = SpecialFolder(feApp)
End Property


' CommonAppData
Public Property Get CDBurnArea() As String
    CDBurnArea = SpecialFolder(feCDBurnArea)
End Property


' CommonAppData
Public Property Get CommonAppData() As String
    CommonAppData = SpecialFolder(feCommonAppData)
End Property


' CommonDesktop
Public Property Get CommonDesktop() As String
    CommonDesktop = SpecialFolder(feCommonDesktop)
End Property


' CommonDocs
Public Property Get CommonDocs() As String
    CommonDocs = SpecialFolder(feCommonDocs)
End Property


' CommonMusic
Public Property Get CommonMusic() As String
    CommonMusic = SpecialFolder(feCommonMusic)
End Property


' CommonPics
Public Property Get CommonPics() As String
    CommonPics = SpecialFolder(feCommonPics)
End Property


' CommonStartMenu
Public Property Get CommonStartMenu() As String
    CommonStartMenu = SpecialFolder(feCommonStartMenu)
End Property


' CommonStartMenuPrograms
Public Property Get CommonStartMenuPrograms() As String
    CommonStartMenuPrograms = SpecialFolder(feCommonStartMenuPrograms)
End Property


' CommonVideos
Public Property Get CommonVideos() As String
    CommonVideos = SpecialFolder(feCommonVideos)
End Property


' LocalAppData
Public Property Get LocalAppData() As String
    LocalAppData = SpecialFolder(feLocalAppData)
End Property


' ProgramFiles
Public Property Get ProgramFiles() As String
    ProgramFiles = SpecialFolder(feProgramFiles)
End Property


' ProgramFilesCommon
Public Property Get ProgramFilesCommon() As String
    ProgramFilesCommon = SpecialFolder(feProgramFilesCommon)
End Property


' User
Public Property Get User() As String
    User = SpecialFolder(feUser)
End Property


' UserAppData
Public Property Get UserAppData() As String
    UserAppData = SpecialFolder(feUserAppData)
End Property


' UserCache
Public Property Get UserCache() As String
    UserCache = SpecialFolder(feUserCache)
End Property


' UserCookies
Public Property Get UserCookies() As String
    UserCookies = SpecialFolder(feUserCookies)
End Property


' UserDesktop
Public Property Get UserDesktop() As String
    UserDesktop = SpecialFolder(feUserDesktop)
End Property


' UserDocs
Public Property Get UserDocs() As String
    UserDocs = SpecialFolder(feUserDocs)
End Property


' UserFavorites
Public Property Get UserFavorites() As String
    UserFavorites = SpecialFolder(feUserFavorites)
End Property


' UserHistory
Public Property Get UserHistory() As String
    UserHistory = SpecialFolder(feLocalHistory)
End Property


' UserMusic
Public Property Get UserMusic() As String
    UserMusic = SpecialFolder(feUserMusic)
End Property


' UserPics
Public Property Get UserPics() As String
    UserPics = SpecialFolder(feUserPics)
End Property


' UserRecent
Public Property Get UserRecent() As String
    UserRecent = SpecialFolder(feUserRecent)
End Property


' UserSendTo
Public Property Get UserSendTo() As String
    UserSendTo = SpecialFolder(feUserSendTo)
End Property


' UserStartMenu
Public Property Get UserStartMenu() As String
    UserStartMenu = SpecialFolder(feUserStartMenu)
End Property


' UserStartMenuPrograms
Public Property Get UserStartMenuPrograms() As String
    UserStartMenuPrograms = SpecialFolder(feUserStartMenuPrograms)
End Property


' UserStartup
Public Property Get UserStartup() As String
    UserStartup = SpecialFolder(feUserStartup)
End Property


' UserTemp
Public Property Get UserTemp() As String
    Dim lngLen As Long
    Dim strReturn As String
    
    strReturn = String(MAX_PATH, 0)
    lngLen = GetTempPath(MAX_PATH, strReturn)
    If lngLen > 0 Then UserTemp = Me.GetLongName(Left$(strReturn, lngLen))
End Property


' UserVideos
Public Property Get UserVideos() As String
    UserVideos = SpecialFolder(feUserVideos)
End Property


' Windows
Public Property Get Windows() As String
    Windows = SpecialFolder(feWindows)
End Property


' WindowsSystem
Public Property Get WindowsSystem() As String
    WindowsSystem = SpecialFolder(feWindowsSystem)
End Property



' INTERNAL FUNCTIONS


Private Sub DateToStructure(pdtm As Date, ByRef ptyp As FILETIME)
    Dim typSystem As SYSTEMTIME
    
    With typSystem
        .wYear = Year(pdtm)
        .wMonth = Month(pdtm)
        .wDay = Day(pdtm)
        .wHour = Hour(pdtm)
        .wMinute = Minute(pdtm)
        .wSecond = Second(pdtm)
    End With
    If SystemTimeToFileTime(typSystem, ptyp) = 1 Then LocalFileTimeToFileTime ptyp, ptyp
End Sub

Private Function SpecialFolder(pfe As FolderEnum) As String
    Dim strPath As String
    Dim strBuffer As String
    Dim lngHandle As Long
    Dim lngLen As Long
    
    strBuffer = Space$(MAX_PATH)
    If pfe = feApp Then
        lngHandle = GetModuleHandle(App.EXEName)
        lngLen = GetModuleFileName(lngHandle, strBuffer, MAX_PATH)
        strPath = Left$(strBuffer, lngLen)
        strPath = Left$(strPath, InStrRev(strPath, "\") - 1)
        If InStr(strPath, "Microsoft Visual Studio") > 0 Then strPath = App.Path
    Else
        If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    SpecialFolder = strPath
End Function

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

Private Sub ValidateDest(ByVal pstrFolder As String)
    Dim blnFolder As Boolean
    Dim lngError As Long
    
    ' Allow a retry for removable disks
    Do While True
        ' Check if file doesn't exist or it's actually a folder
        On Error Resume Next
        blnFolder = GetAttr(pstrFolder) And vbDirectory
        lngError = Err.Number
        On Error GoTo 0
        Select Case lngError
            Case 0, 53: Exit Do
            Case 5 ' Drive not ready
                If MsgBox(Left$(pstrFolder, 2) & " Drive not ready", vbInformation + vbRetryCancel, "Disk Error") = vbCancel Then
                    Err.Raise 71, , "Disk not ready"
                End If
            Case 52, 76 ' Invalid path or Bad command or file name
                Err.Raise lngError, , "'" & pstrFolder & "' not a valid path"
            Case Else: Err.Raise lngError
        End Select
    Loop
End Sub

Private Sub ValidateFolder(ByVal pstrFolder As String)
    Dim blnFolder As Boolean
    Dim lngError As Long
    
    ' Allow a retry for removable disks
    Do While True
        ' Check if folder doesn't exist or it's actually a file
        On Error Resume Next
        blnFolder = GetAttr(pstrFolder) And vbDirectory
        lngError = Err.Number
        On Error GoTo 0
        Select Case lngError
            Case 0: Exit Do
            Case 5 ' Drive not ready
                If MsgBox(Left$(pstrFolder, 2) & " Drive not ready", vbInformation + vbRetryCancel, "Disk Error") = vbCancel Then
                    Err.Raise 71, , "Disk not ready"
                End If
            Case 52, 76 ' Invalid path or Bad command or file name
                Err.Raise lngError, , "'" & pstrFolder & "' not a valid path"
            Case 53 ' File not found
                Err.Raise 53, , "'" & pstrFolder & "' not found"
            Case Else: Err.Raise lngError
        End Select
    Loop
End Sub
