VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Drive"
Attribute VB_Ext_KEY = "Member1" ,"File"
Attribute VB_Ext_KEY = "Member2" ,"Folder"
Attribute VB_Ext_KEY = "Member3" ,"Mouse"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member4" ,"Registry"
' Written by Ellis Dee
' Generic functions for interacting with Windows. Also includes clsFile and clsFolder instances.
Option Explicit

' Enumerations
Public Enum MouseStateEnum
    msNormal
    msWait
    msAppWait
    msSystemWait
    msBackgroundWait
    msHidden
End Enum

Public Enum MouseCursorEnum
    mcAppStarting = 32650&
    mcArrow = 32512&
    mcCross = 32515&
    mcIBeam = 32513&
    mcHand = 32649&
    mcIcon = 32641&
    mcNo = 32648&
    mcSize = 32640&
    mcSizeAll = 32646&
    mcSizeNew = 32643&
    mcSizeNS = 32645&
    mcSizeNWSE = 32642&
    mcSizeWE = 32644&
    mcUpArrow = 32516&
    mcWait = 32514&
End Enum

Private Enum FlagsEnum
    feOpen
    feSaveAs
End Enum

' Property variables
Private menMouse As MouseStateEnum

' Variables
Private mlngLockCount As Long

' Child classes
Private mclsFile As clsFile
Private mclsFolder As clsFolder

' Type declarations for API calls
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
   Attributes As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PRINTDLGS
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

' API
Private Declare Function AdjustTokenPrivileges Lib "Advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "Advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function OpenProcessToken Lib "Advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Declare Sub SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long



' METHODS



Public Sub AutoHideTaskbar(Optional AutoHide As Boolean = True)
    Const ABM_GETSTATE = &H4
    Const ABM_SETSTATE = &HA
    Const ABS_AUTOHIDE = &H1
    Const ABS_ALWAYSONTOP = &H2
    Dim typAppBar As APPBARDATA
    
    typAppBar.cbSize = Len(typAppBar)
    SHAppBarMessage ABM_GETSTATE, typAppBar
    If AutoHide Then
        typAppBar.lParam = ABS_AUTOHIDE Or ABS_ALWAYSONTOP
    Else
        typAppBar.lParam = ABS_ALWAYSONTOP
    End If
    SHAppBarMessage ABM_SETSTATE, typAppBar
End Sub

' Force Windows to refresh file associations
Public Sub AssociateIcon()
    Const SHCNE_ASSOCCHANGED = &H8000000
    Const SHCNF_IDLIST = 0
    
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0&, 0&
End Sub

Public Sub EmptyRecycleBin()
    Const SHERB_NOCONFIRMATION As Long = &H1
    
    SHEmptyRecycleBin 0&, vbNullString, SHERB_NOCONFIRMATION
End Sub

' Get desktop coords excluding any taskbars
' Thanks to bushmobile from vbforums.com
Public Sub GetDesktop(Left As Long, Top As Long, Width As Long, Height As Long)
    Const SPI_GETWORKAREA As Long = 48
    Dim rc As RECT
    
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
    With rc
        Left = .Left * Screen.TwipsPerPixelX
        Top = .Top * Screen.TwipsPerPixelY
        Width = (.Right - .Left) * Screen.TwipsPerPixelX
        Height = (.Bottom - .Top) * Screen.TwipsPerPixelY
    End With
End Sub

' Suppress paint requests to a window
' (reduces flicker, but don't abuse as a crutch)
Public Sub LockWindow(ByVal hwnd As Long)
    mlngLockCount = mlngLockCount + 1
    If mlngLockCount = 1 Then LockWindowUpdate hwnd
End Sub

' Launch default browser
Public Sub OpenURL(ByVal URL As String)
    ShellExecute 0&, "OPEN", URL, vbNullString, vbNullString, vbNormalFocus
End Sub

Public Function DownloadURL(ByVal URL As String) As String
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0
    Const INTERNET_OPEN_TYPE_DIRECT = 1
    Const INTERNET_OPEN_TYPE_PROXY = 3
    Const INTERNET_FLAG_DONT_CACHE = &H4000000
    Const scUserAgent = "VB Project"
    Const INTERNET_FLAG_RELOAD = &H80000000
    Dim lngOpen As Long
    Dim lngOpenURL As Long
    Dim blnReturn As Boolean
    Dim strReadBuffer As String * 2048
    Dim lngNumberOfBytesRead As Long
    Dim strBuffer As String

    lngOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    lngOpenURL = InternetOpenUrl(lngOpen, URL, vbNullString, 0, INTERNET_FLAG_DONT_CACHE Or INTERNET_FLAG_RELOAD, 0)
    Do
        strReadBuffer = vbNullString
        blnReturn = InternetReadFile(lngOpenURL, strReadBuffer, Len(strReadBuffer), lngNumberOfBytesRead)
        strBuffer = strBuffer & Left$(strReadBuffer, lngNumberOfBytesRead)
        If Not CBool(lngNumberOfBytesRead) Then Exit Do
    Loop
    If lngOpenURL <> 0 Then InternetCloseHandle (lngOpenURL)
    If lngOpen <> 0 Then InternetCloseHandle (lngOpen)
    DownloadURL = strBuffer
End Function

' Engage default screensaver
Public Sub ScreenSaver()
    Const WM_SYSCOMMAND = &H112&
    Const SC_SCREENSAVE = &HF140&
    Dim frm As Form
    Dim hwnd As Long
    
    For Each frm In Forms
        hwnd = frm.hwnd
        Set frm = Nothing
        Exit For
    Next
    SendMessage hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
End Sub

' Set/unset a form as AlwaysOnTop
Public Sub SetAlwaysOnTop(ByVal hwnd As Long, Optional ByVal AlwaysOnTop As Boolean = True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_SHOWWINDOW = &H40
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    
    If AlwaysOnTop Then
        SetWindowPos hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub

' Set mouse cursor from MouseMove()
Public Sub SetMouseCursor(ByVal Cursor As MouseCursorEnum)
    SetCursor LoadCursor(0, Cursor)
End Sub

' Move the mouse cursor
Public Sub SetMousePosition(ByVal PixelX As Long, ByVal PixelY As Long)
    If PixelX < 0 Then PixelX = 0
    If PixelY < 0 Then PixelY = 0
    With Screen
        If PixelX > .Width Then PixelX = .Width
        If PixelY > .Height Then PixelY = .Height
    End With
    SetCursorPos PixelX, PixelY
End Sub

' Show the common dialog for choosing a color.
' Return the chosen color, or -1 if the dialog is canceled
'
' hParent is the handle of the parent form
' bFullOpen specifies whether the dialog will be open with the Full style
' (allows to choose many more colors)
' InitColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorsDialog(Me.hwnd, True, vbRed)
'    If oleNewColor <> -1 Then Me.BackColor = oleNewColor
Function ShowColorDialog(Optional ByVal hParent As Long, Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) As Long
    Const CC_RGBINIT = &H1&
    Const CC_FULLOPEN = &H2&
    Const CC_PREVENTFULLOPEN = &H4&
    Const CC_SHOWHELP = &H8&
    Const CC_ENABLEHOOK = &H10&
    Const CC_ENABLETEMPLATE = &H20&
    Const CC_ENABLETEMPLATEHANDLE = &H40&
    Const CC_SOLIDCOLOR = &H80&
    Const CC_ANYCOLOR = &H100&
    Const CLR_INVALID = &HFFFF
    Dim cc As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With cc
        .lStructSize = Len(cc)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(cc) Then
        'if not canceled, return the color
        ShowColorDialog = cc.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If
End Function

' Common Dialog - Open
Public Function ShowOpenDialog(pstrInitialPath As String, pstrFilter As String, pstrDefaultExt As String) As String
    Dim typFileName As OPENFILENAME
            
    typFileName = GetStructure(pstrInitialPath, "", pstrFilter, pstrDefaultExt, feOpen)
    If GetOpenFileName(typFileName) Then ShowOpenDialog = Left$(typFileName.lpstrFile, InStr(typFileName.lpstrFile, Chr$(0)) - 1)
End Function

' Common Dialog - Print
Public Function ShowPrinterDialog() As Long
    Dim typPrinter As PRINTDLGS
    
    typPrinter.lStructSize = Len(typPrinter)
    ShowPrinterDialog = PrintDlg(typPrinter)
End Function

' Common Dialog - SaveAs
Public Function ShowSaveAsDialog(pstrInitialPath As String, pstrFile As String, pstrFilter As String, pstrDefaultExt As String) As String
    Dim typFileName As OPENFILENAME
    
    typFileName = GetStructure(pstrInitialPath, pstrFile, pstrFilter, pstrDefaultExt, feSaveAs)
    If GetSaveFileNameW(VarPtr(typFileName)) Then ShowSaveAsDialog = Left$(typFileName.lpstrFile, InStr(typFileName.lpstrFile, Chr$(0)) - 1)
End Function

' Shut down windows, and optional reboot it
' if the 2nd argument is True, no WM_QUERYENDSESSION and WM_ENDSESSION
' messages are sent to active applications
Public Sub ShutDownWindows(Optional ByVal Reboot As Boolean = False, Optional ByVal Force As Boolean)
    Const EWX_LOGOFF = 0
    Const EWX_SHUTDOWN = 1
    Const EWX_REBOOT = 2
    Const EWX_FORCE = 4
    Const EWX_POWEROFF = 8
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hToken As Long
    Dim tp As TOKEN_PRIVILEGES
    Dim flags As Long
    
    ' Windows NT/2000 require a special treatment
    ' to ensure that the calling process has the
    ' privileges to shut down the system
    
    ' under NT the high-order bit (that is, the sign bit)
    ' of the value retured by GetVersion is cleared
    If GetVersion() >= 0 Then
        ' Open this process for adjusting its privileges
        OpenProcessToken GetCurrentProcess(), (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hToken
        
        ' Get the LUID for shutdown privilege.
        ' retrieves the locally unique identifier (LUID) used
        ' to locally represent the specified privilege name
        ' (first argument = "" means the local system)
        LookupPrivilegeValue "", "SeShutdownPrivilege", tp.LuidUDT
        
        ' complete the TOKEN_PRIVILEGES structure with the # of
        ' privileges and the desired attribute
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        
        ' enables or disables privileges in the specified access token
        ' last 3 arguments are zero because we aren't interested
        ' in previous privilege attributes.
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
    End If
    
    ' prepare shutdown flags
    flags = EWX_SHUTDOWN
    If Reboot Then
        flags = flags Or EWX_REBOOT
    Else
        flags = flags Or EWX_POWEROFF
    End If
    If Force Then flags = flags Or EWX_FORCE
    
    ' finally, you can shut down Windows
    ExitWindowsEx flags, &HFFFF
End Sub

' Translate system color into its resulting color value
Public Function SystemColorRGB(ByVal plngSystemColor As Long) As Long
    Const S_OK = &H0
    Dim lngReturn As Long
    
    If OleTranslateColor(plngSystemColor, 0, lngReturn) = S_OK Then SystemColorRGB = lngReturn Else SystemColorRGB = plngSystemColor
End Function

Public Sub ColorToRGB(ByVal plngColor As Long, plngRed As Long, plngGreen As Long, plngBlue As Long)
    If plngColor < 0 Then plngColor = SystemColorRGB(plngColor)
    plngRed = &HFF& And plngColor
    plngGreen = (&HFF00& And plngColor) \ 256
    plngBlue = (&HFF0000 And plngColor) \ 65536
End Sub

Public Sub ColorToRGBDebug(plngColor As Long)
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    'Debug.Print "RGB " & lngRed & "," & lngGreen & "," & lngBlue
End Sub

Public Function HexToColor(ByVal pstrHex As String) As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    If Not HexColorIsValid(pstrHex) Then Exit Function
    If Len(pstrHex) = 7 Then pstrHex = Right$(pstrHex, 6)
    lngRed = CLng("&H" & Left$(pstrHex, 2))
    lngGreen = CLng("&H" & Mid$(pstrHex, 3, 2))
    lngBlue = CLng("&H" & Right$(pstrHex, 2))
    HexToColor = RGB(lngRed, lngGreen, lngBlue)
End Function

Public Function HexColorIsValid(ByVal pstrHex As String) As Boolean
    Dim i As Long
    
    If Left$(pstrHex, 1) = "#" Then pstrHex = Mid$(pstrHex, 2)
    If Len(pstrHex) <> 6 Then Exit Function
    pstrHex = UCase$(pstrHex)
    For i = 1 To 6
        If InStr("1234567890ABCDEF", Mid$(pstrHex, i, 1)) = 0 Then Exit Function
    Next
    HexColorIsValid = True
End Function

Public Function ColorToHex(ByVal plngColor As Long) As String
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    If plngColor < 0 Then plngColor = SystemColorRGB(plngColor)
    ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    ColorToHex = HexColorValue(lngRed) & HexColorValue(lngGreen) & HexColorValue(lngBlue)
End Function

Private Function HexColorValue(plngValue As Long) As String
    Dim strReturn As String
    
    strReturn = Hex(plngValue)
    If Len(strReturn) < 2 Then strReturn = "0" & strReturn
    HexColorValue = strReturn
End Function

' Enable paint requests
Public Sub UnlockWindow()
    mlngLockCount = mlngLockCount - 1
    If mlngLockCount = 0 Then LockWindowUpdate 0
End Sub



' PROPERTIES


' Debug Mode
Public Property Get DebugMode() As Boolean
    DebugMode = (App.LogMode = 0)
End Property


' File class
Public Property Get File() As clsFile
    Set File = mclsFile
End Property

Public Property Set File(pcls As clsFile)
    Set mclsFile = pcls
End Property


' Folder class
Public Property Get Folder() As clsFolder
    Set Folder = mclsFolder
End Property

Public Property Set Folder(pcls As clsFolder)
    Set mclsFolder = pcls
End Property


' Set mouse cursor
Public Property Let Mouse(ByVal penState As MouseStateEnum)
    Const IDC_ARROW As Long = 32512&
    Const IDC_WAIT = 32514&
    Const IDC_APPSTARTING = 32650&
    Dim lngSysPointer As Long
    Dim lngFormPointer As Long
    Dim lngCursor As Long
    Static sblnReset As Boolean

    menMouse = penState
    ' Set the cursor
    Select Case menMouse
        Case msNormal
            If Not (Screen.ActiveForm Is Nothing) Then Screen.ActiveForm.MousePointer = vbDefault
            Screen.MousePointer = vbDefault
            If sblnReset Then
                SetCursor LoadCursor(0, IDC_ARROW)
                sblnReset = False
            End If
        Case msWait
            If Not (Screen.ActiveForm Is Nothing) Then Screen.ActiveForm.MousePointer = vbHourglass
        Case msAppWait
            Screen.MousePointer = vbHourglass
        Case msSystemWait
            SetCursor LoadCursor(0, IDC_WAIT)
            sblnReset = True
        Case msBackgroundWait
            SetCursor LoadCursor(0, IDC_APPSTARTING)
            sblnReset = True
    End Select
    ' Show/hide cursor
    If menMouse = msHidden Then
        Do While ShowCursor(False) >= 0
        Loop
    Else
        Do While ShowCursor(True) < 0
        Loop
    End If
End Property

Public Property Get Mouse() As MouseStateEnum
    Mouse = menMouse
End Property


' Get current Windows UserName
Public Property Get UserName() As String
    Dim strUser As String * 50
    Dim lngLen As Long

    lngLen = 50
    If GetUserName(strUser, lngLen) = 0 Then
        UserName = ""
    Else
        UserName = Left$(strUser, lngLen - 1)
    End If
End Property



' INTERNAL FUNCTIONS


Private Sub Class_Initialize()
    menMouse = msNormal
    Set mclsFile = New clsFile
    Set mclsFolder = New clsFolder
    mlngLockCount = 0
End Sub

Private Sub Class_Terminate()
    Me.Mouse = msNormal
    Set mclsFile = Nothing
    Set mclsFolder = Nothing
End Sub

Private Function GetStructure(pstrPath As String, pstrFile As String, pstrFilter As String, pstrDefaultExt As String, penFlags As FlagsEnum) As OPENFILENAME
    Const OFN_FILEMUSTEXIST = &H1000
    Const OFN_PATHMUSTEXIST = &H800
    Const OFN_HIDEREADONLY = &H4
    Const OFN_LONGNAMES = &H200000
    Const OFN_OVERWRITEPROMPT = &H2
    Const OF_WRITE = &H1
    Const MAX_PATH = 260

    With GetStructure
        .lStructSize = Len(GetStructure)
        .hwndOwner = Screen.ActiveForm.hwnd
'        .hInstance = App.hInstance
        .lpstrFilter = Replace(pstrFilter, "|", vbNullChar) & vbNullChar
        .nMaxFile = MAX_PATH + 1
        .nMaxFileTitle = MAX_PATH + 1
        .lpstrFileTitle = String(MAX_PATH, 0)
        .lpstrInitialDir = pstrPath
        .lpstrDefExt = pstrDefaultExt
        Select Case penFlags
            Case feOpen
                .lpstrTitle = "Open"
                .lpstrFile = Space(MAX_PATH)
                .flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_LONGNAMES
            Case feSaveAs
                .lpstrTitle = "Save As"
                .lpstrFile = pstrFile & Space$(MAX_PATH - Len(pstrFile))
                .flags = OFN_PATHMUSTEXIST + OFN_HIDEREADONLY + OFN_LONGNAMES + OF_WRITE ' + OFN_OVERWRITEPROMPT
        End Select
    End With
End Function
