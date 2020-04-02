Attribute VB_Name = "basRaceCombo"
Option Explicit

Public Const GWL_WNDPROC As Long = -4
Private Const CBN_DROPDOWN As Long = 7
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_KEYDOWN As Long = &H100
Private Const VK_F4 As Long = &H73
Private Const VK_DOWN As Long = &H28

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type COMBOBOXINFO
    cbSize As Long
    rcItem As RECT
    rcButton As RECT
    stateButton  As Long
    hwndCombo  As Long
    hwndEdit  As Long
    hwndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public mlngHooked As Long

Public Sub ComboGetEditRect(pcbo As ComboBox, plngLeft As Long, plngTop As Long, plngWidth As Long, plngHeight As Long)
    Dim typComboBox As COMBOBOXINFO
    
    typComboBox.cbSize = Len(typComboBox)
    GetComboBoxInfo pcbo.hwnd, typComboBox
    With typComboBox.rcItem
        plngLeft = pcbo.Left + (.Left * Screen.TwipsPerPixelX)
        plngTop = pcbo.Top + (.Top * Screen.TwipsPerPixelY)
        plngWidth = (.Right - .Left) * Screen.TwipsPerPixelX
        plngHeight = (.Bottom - .Top) * Screen.TwipsPerPixelY
    End With
End Sub

Public Sub HookRaceCombo(hwnd As Long)
    If mlngHooked = 0 Then mlngHooked = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf RaceComboProc)
End Sub

Public Sub UnhookRaceCombo(hwnd As Long)
    If mlngHooked <> 0 Then SetWindowLong hwnd, GWL_WNDPROC, mlngHooked
    mlngHooked = 0
End Sub

Public Function RaceComboProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hwnd <> frmOverview.cboRace.hwnd Then Exit Function
    Select Case uMsg
        Case CBN_DROPDOWN, WM_LBUTTONDOWN
            frmOverview.RaceToggle
        Case WM_KEYDOWN
        Case Else
            RaceComboProc = CallWindowProc(mlngHooked, hwnd, uMsg, wParam, lParam)
            Exit Function
    End Select
    On Error Resume Next
    frmOverview.txtRace.SetFocus
    On Error GoTo 0
    RaceComboProc = 1
End Function

