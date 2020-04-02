Attribute VB_Name = "basMouseWheel"
' Core functionality written by bushmobile of VBForums.com
' Thread title: VB6 - MouseWheel with Any Control (originally just MSFlexGrid Scrolling)
' Thread link: www.vbforums.com/showthread.php?388222
Option Explicit

' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
' Position Checking
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Check Messages
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lngMouseKeys As Long
    Dim lngRotation As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim frm As Form
    
    If Lmsg = WM_MOUSEWHEEL Then
        lngMouseKeys = wParam And 65535
        lngRotation = wParam / 65536
        lngX = lParam And 65535
        lngY = lParam / 65536
        Set frm = GetForm(Lwnd)
        If frm Is Nothing Then
            ' it's not a form
            If Not IsOver(Lwnd, lngX, lngY) And IsOver(GetParent(Lwnd), lngX, lngY) Then
                ' it's not over the control and is over the form,
                ' so fire mousewheel on form (if it's not a dropped down combo)
                If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
                    GetForm(GetParent(Lwnd)).MouseWheel lngMouseKeys, lngRotation, lngX, lngY
                    Exit Function ' Discard scroll message to control
                End If
            End If
        Else
            ' it's a form so fire mousewheel
            If IsOver(frm.hWnd, lngX, lngY) Then frm.MouseWheel lngMouseKeys, lngRotation, lngX, lngY
        End If
    End If
    WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hWnd As Long)
    On Error Resume Next
    SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hWnd As Long)
    On Error Resume Next
    SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
    RemoveProp hWnd, "PrevWndProc"
End Sub

' Window Checks
' ================================================
Public Function IsOver(ByVal hWnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
    Dim rectCtl As RECT
    
    GetWindowRect hWnd, rectCtl
    With rectCtl
        IsOver = (lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom)
    End With
End Function

Private Function GetForm(ByVal hWnd As Long) As Form
    For Each GetForm In Forms
        If GetForm.hWnd = hWnd Then Exit Function
    Next GetForm
    Set GetForm = Nothing
End Function

Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    picBox.Cls
    picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
End Sub


