Attribute VB_Name = "basLibrary"
' Written by Ellis Dee
' This module contains select helper functions originally found in:
' basControls, basUtils, clsFile, clsWindowsXP
' Not including those full modules saves ~100k in final exe sizes
Option Explicit

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

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long


' ************* CLSWINDOWSXP *************


' Set mouse cursor from MouseMove()
Public Sub SetMouseCursor(ByVal Cursor As MouseCursorEnum)
    SetCursor LoadCursor(0, Cursor)
End Sub

' Set/unset a form as AlwaysOnTop
Public Sub SetAlwaysOnTop(ByVal hWnd As Long, Optional ByVal AlwaysOnTop As Boolean = True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_SHOWWINDOW = &H40
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1

    If AlwaysOnTop Then
        SetWindowPos hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    Else
        SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub

Public Function DebugMode() As Boolean
    DebugMode = (App.LogMode = 0)
End Function


' ************* CONTROLS *************


' Function to make checkboxes behave like command buttons. Graphical checkboxes look like command buttons,
' plus you can change their forecolor and backcolor. (Command buttons cannot change forecolor.)
' Add the following line to start of the checkbox's click event:
' If UncheckButton(CheckBox, mblnOverride) Then Exit Sub
Public Function UncheckButton(pchk As CheckBox, pblnOverride As Boolean) As Boolean
    If pblnOverride Then
        pblnOverride = False
        UncheckButton = True
    Else
        pblnOverride = True
        pchk.Value = vbUnchecked
    End If
End Function


' ************* FORM *************


Public Function GetForm(pfrm As Form, pstrFormName As String) As Boolean
    For Each pfrm In Forms
        If pfrm.Name = pstrFormName Then
            GetForm = True
            Exit Function
        End If
    Next
    Set pfrm = Nothing
End Function

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

Public Sub PositionForm(pfrm As Form, pstrValue As String)
    Dim lngDesktopLeft As Long
    Dim lngDesktopTop As Long
    Dim lngDesktopWidth As Long
    Dim lngDesktopHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    
    GetDesktop lngDesktopLeft, lngDesktopTop, lngDesktopWidth, lngDesktopHeight
    ReadCoords pstrValue, lngLeft, lngTop
    If lngLeft = 0 And lngTop = 0 Then
        Select Case pstrValue
            Case "ADQ1"
                lngLeft = (lngDesktopWidth - pfrm.Width) \ 2
                lngTop = (lngDesktopHeight - pfrm.Height) \ 3
            Case "ADQ2"
                lngLeft = (lngDesktopWidth - pfrm.Width) \ 2
                lngTop = lngDesktopTop
            Case "Solver"
                lngLeft = lngDesktopLeft + lngDesktopWidth - pfrm.Width
                lngTop = lngDesktopTop
            Case "Stopwatch"
                lngLeft = lngDesktopLeft + lngDesktopWidth - pfrm.Width
                lngTop = lngDesktopTop + (lngDesktopHeight - pfrm.Height) \ 2
        End Select
    End If
    If lngLeft < lngDesktopLeft Then lngLeft = lngDesktopLeft
    If lngLeft + pfrm.Width > lngDesktopLeft + lngDesktopWidth Then lngLeft = lngDesktopLeft + lngDesktopWidth - pfrm.Width
    If lngTop < lngDesktopTop Then lngTop = lngDesktopTop
    If lngTop + pfrm.Height > lngDesktopTop + lngDesktopHeight Then lngTop = lngDesktopTop + lngDesktopHeight - pfrm.Height
    pfrm.Move lngLeft, lngTop
End Sub


' ************* FILES *************


Public Function FileExists(pstrFile As String) As Boolean
    FileExists = (PathFileExists(pstrFile) = 1)
    If FileExists Then FileExists = (PathIsDirectory(pstrFile) = 0)
End Function

Public Function LoadToString(pstrFile As String) As String
    Dim strReturn As String
    Dim FileNumber As Long

    FileNumber = FreeFile
    Open pstrFile For Binary Access Read As #FileNumber
    If LOF(FileNumber) > 0 Then
        strReturn = Space(LOF(FileNumber))
        Get #FileNumber, , strReturn
    End If
    Close #FileNumber
    LoadToString = strReturn
End Function

Public Function SaveStringAs(pstrFile As String, pstrText As String) As Boolean
On Error GoTo SaveStringAsErr
    Dim FileNumber As Long

    FileNumber = FreeFile
SaveStringAsRetry:
    Open pstrFile For Output As #FileNumber
    Print #FileNumber, pstrText
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


' ************* SETTINGS (COORDS) *************


Public Sub ReadCoords(ByVal pstrValue As String, plngLeft As Long, plngTop As Long)
    Dim strFile As String
    Dim strLine() As String
    Dim strToken() As String
    Dim i As Long
    
    strFile = CoordsFile()
    If Not FileExists(strFile) Then Exit Sub
    strLine = Split(LoadToString(strFile), vbNewLine)
    pstrValue = pstrValue & "Coords="
    For i = 0 To UBound(strLine)
        If Left$(strLine(i), Len(pstrValue)) = pstrValue Then
            strToken = Split(Mid$(strLine(i), Len(pstrValue) + 1), ",")
            If UBound(strToken) = 1 Then
                plngLeft = Val(strToken(0)) * Screen.TwipsPerPixelX
                plngTop = Val(strToken(1)) * Screen.TwipsPerPixelY
            End If
            Exit Sub
        End If
    Next
End Sub

Public Sub WriteCoords(ByVal pstrValue As String, plngLeft As Long, plngTop As Long)
    Dim strFile As String
    Dim strLine() As String
    Dim strRaw As String
    Dim i As Long
    
    pstrValue = pstrValue & "Coords="
    strFile = CoordsFile()
    strLine = Split(LoadToString(strFile), vbNewLine)
    For i = 0 To UBound(strLine)
        If Left$(strLine(i), Len(pstrValue)) = pstrValue Then Exit For
    Next
    If i > UBound(strLine) Then ReDim Preserve strLine(i)
    strLine(i) = pstrValue & plngLeft \ Screen.TwipsPerPixelX & "," & plngTop \ Screen.TwipsPerPixelY
    strRaw = Join(strLine, vbNewLine)
    strRaw = Replace(strRaw, vbNewLine & vbNewLine, vbNewLine)
    Do While Left$(strRaw, 2) = vbNewLine
        strRaw = Mid$(strRaw, 3)
    Loop
    Do While Right$(strRaw, 2) = vbNewLine
        strRaw = Left$(strRaw, Len(strRaw) - 2)
    Loop
    SaveStringAs CoordsFile(), strRaw
End Sub

Private Function CoordsFile() As String
    CoordsFile = App.Path & "\Config.txt"
End Function
