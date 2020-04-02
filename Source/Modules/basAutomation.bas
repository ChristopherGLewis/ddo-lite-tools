Attribute VB_Name = "basAutomation"
Option Explicit

' Declarations for GetEXE(), written by Kaverin of vbForums.com
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
 
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
 
Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH
End Type

' Declarations for EnumMessageWindows, written by Ellis Dee
Private Const MessageWindowCaption As String = "EllisoftLiteMessages"
Private Const WM_SETTEXT As Long = &HC

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private mstrMessage As String
Private mlngCompendiums As Long


' ************* BROADCAST MESSAGES *************


Public Sub Broadcast(ByVal pstrMessage As String)
    mstrMessage = pstrMessage
    EnumWindows AddressOf EnumMessageWindows, 0
End Sub

' Callback function
Public Function EnumMessageWindows(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim lngLen As Long
    Dim strBuffer As String
    Dim strTitleBar As String
    Dim lngReturn As Long ' return value
    Dim lngChild As Long
    
    lngLen = GetWindowTextLength(hWnd) + 1 ' get length of title bar text
    If lngLen > 1 Then ' if return value refers to non-empty string
        strBuffer = Space(lngLen) ' make room in the strBuffer
        lngReturn = GetWindowText(hWnd, strBuffer, lngLen) ' get title bar text
        strTitleBar = Left$(strBuffer, lngLen - 1)
        If strTitleBar = MessageWindowCaption Then
            lngChild = FindWindowEx(hWnd, 0, vbNullString, vbNullString)
            If lngChild Then SendMessageTimeout lngChild, WM_SETTEXT, 0, ByVal mstrMessage, 0, 100, 0
        End If
    End If
    EnumMessageWindows = True ' 1 ' return value of 1 means continue enumeration
End Function


' ************* COUNT COMPENDIUMS *************


Public Function CountCompendiums() As Long
    mlngCompendiums = 0
    EnumWindows AddressOf EnumCompendiums, 0
    CountCompendiums = mlngCompendiums
End Function

' Callback function
Public Function EnumCompendiums(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim lngLen As Long
    Dim strBuffer As String
    Dim strTitleBar As String
    Dim lngReturn As Long ' return value
    Dim lngChild As Long
    
    lngLen = GetWindowTextLength(hWnd) + 1 ' get length of title bar text
    If lngLen > 1 Then ' if return value refers to non-empty string
        strBuffer = Space(lngLen) ' make room in the strBuffer
        lngReturn = GetWindowText(hWnd, strBuffer, lngLen) ' get title bar text
        strTitleBar = Left$(strBuffer, lngLen - 1)
        If strTitleBar = MessageWindowCaption Then
            If LCase$(GetEXE(hWnd)) = "compendium.exe" Then mlngCompendiums = mlngCompendiums + 1
        End If
    End If
    EnumCompendiums = True ' 1 ' return value of 1 means continue enumeration
End Function


' ************* GET EXE *************


' GetEXE() originally written by Kaverin of vbForums.com
' Source: http://www.vbforums.com/showthread.php?198867-Getting-application-name-via-window-handle&p=1176825&viewfull=1#post1176825
 
'returns the exe name of the hWnd's process, or "" if it can't be determined
Private Function GetEXE(ByVal hWnd As Long) As String
   Dim lngPID As Long
   Dim lngProcess As Long
   Dim lngSnapshot As Long
   Dim typInfo As PROCESSENTRY32
   Dim lngReturn As Long
   
    If hWnd = 0 Then Exit Function
    'get the lngPID of the window's thread
    If GetWindowThreadProcessId(hWnd, lngPID) Then
        'attempt to get a handle to the process (this must be closed after use)
        lngProcess = OpenProcess(PROCESS_ALL_ACCESS, False, lngPID)
        If lngProcess Then
            'make a snapshot of the processes (this must be closed after use)
            lngSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, lngPID)
            If lngSnapshot Then
                'set the size of the structure
                typInfo.dwSize = Len(typInfo)
                'get the first process in the snapshot
                lngReturn = Process32First(lngSnapshot, typInfo)
                'loop until there are no more
                Do While lngReturn
                    'if the current lngPID matches the lngPID we have already, stop the loop
                    If typInfo.th32ProcessID = lngPID Then
                        'trim off the null terminator and return it
                        GetEXE = Left$(typInfo.szExeFile, InStr(typInfo.szExeFile & vbNullChar, vbNullChar) - 1)
                        Exit Do
                    End If
                    'get the next process
                    lngReturn = Process32Next(lngSnapshot, typInfo)
                Loop
                'close the snapshot handle
                CloseHandle lngSnapshot
            End If
            'close the process handle
            CloseHandle lngProcess
        End If
    End If
End Function


