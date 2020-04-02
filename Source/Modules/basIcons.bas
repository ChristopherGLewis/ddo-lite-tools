Attribute VB_Name = "basIcons"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

Public Sub SetFormIcon(pfrm As Form, ByVal pstrResourceID As String)
    Dim frm As Form
    Dim lnghWnd As Long
    Dim blnFound As Boolean
    Dim lngIcon As Long
    Dim lngX As Long
    Dim lngY As Long
    
    ' Doesn't work in IDE, so just use crappy LoadResPicture() and bail
    If (App.LogMode = 0) Then
        pfrm.Icon = LoadResPicture(pstrResourceID, vbResIcon)
        Exit Sub
    End If
    lnghWnd = pfrm.hwnd
    ' Set large icon
    lngX = GetSystemMetrics(SM_CXICON)
    lngY = GetSystemMetrics(SM_CYICON)
    lngIcon = LoadImageAsString(App.hInstance, pstrResourceID, IMAGE_ICON, lngX, lngY, LR_SHARED)
    SendMessageLong lnghWnd, WM_SETICON, ICON_BIG, lngIcon
    ' Set small icon
    lngX = GetSystemMetrics(SM_CXSMICON)
    lngY = GetSystemMetrics(SM_CYSMICON)
    lngIcon = LoadImageAsString(App.hInstance, pstrResourceID, IMAGE_ICON, lngX, lngY, LR_SHARED)
    SendMessageLong lnghWnd, WM_SETICON, ICON_SMALL, lngIcon
End Sub

Public Sub SetAppIcon(ByVal pstrResourceID As String)
    Dim frm As Form
    Dim blnFound As Boolean
    Dim lnghWndTop As Long
    Dim lnghWnd As Long
    Dim lngIcon As Long
    Dim lngX As Long
    Dim lngY As Long
    
    ' Doesn't work in IDE, so just bail
    If (App.LogMode = 0) Then Exit Sub
    ' Grab any form
    For Each frm In Forms
        lnghWnd = frm.hwnd
        blnFound = True
        Exit For
    Next
    Set frm = Nothing
    If Not blnFound Then Exit Sub
    ' Find VB's hidden parent window:
    lnghWndTop = lnghWnd
    Do While lnghWnd
        lnghWnd = GetWindow(lnghWnd, GW_OWNER)
        If lnghWnd Then lnghWndTop = lnghWnd
    Loop
    ' Set large icon
    lngX = GetSystemMetrics(SM_CXICON)
    lngY = GetSystemMetrics(SM_CYICON)
    lngIcon = LoadImageAsString(App.hInstance, pstrResourceID, IMAGE_ICON, lngX, lngY, LR_SHARED)
    SendMessageLong lnghWndTop, WM_SETICON, ICON_BIG, lngIcon
    ' Set small icon
    lngX = GetSystemMetrics(SM_CXSMICON)
    lngY = GetSystemMetrics(SM_CYSMICON)
    lngIcon = LoadImageAsString(App.hInstance, pstrResourceID, IMAGE_ICON, lngX, lngY, LR_SHARED)
    SendMessageLong lnghWndTop, WM_SETICON, ICON_SMALL, lngIcon
End Sub

'' Original code from vbaccelerator.com article: "Providing a proper VB Application Icon, Including Large Icons and 32-Bit Alpha Images"
'' http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp
'Private Sub SetIconAPI(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
'    Dim lhWndTop As Long
'    Dim lhWnd As Long
'    Dim cx As Long
'    Dim cy As Long
'    Dim hIconLarge As Long
'    Dim hIconSmall As Long
'
'    If (bSetAsAppIcon) Then
'        ' Find VB's hidden parent window:
'        lhWnd = hWnd
'        lhWndTop = lhWnd
'        Do While Not (lhWnd = 0)
'            lhWnd = GetWindow(lhWnd, GW_OWNER)
'            If Not (lhWnd = 0) Then
'                lhWndTop = lhWnd
'            End If
'        Loop
'    End If
'
'    cx = GetSystemMetrics(SM_CXICON)
'    cy = GetSystemMetrics(SM_CYICON)
'    hIconLarge = LoadImageAsString( _
'    App.hInstance, sIconResName, _
'    IMAGE_ICON, _
'    cx, cy, _
'    LR_SHARED)
'    If (bSetAsAppIcon) Then
'        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
'    End If
'    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
'
'    cx = GetSystemMetrics(SM_CXSMICON)
'    cy = GetSystemMetrics(SM_CYSMICON)
'    hIconSmall = LoadImageAsString( _
'    App.hInstance, sIconResName, _
'    IMAGE_ICON, _
'    cx, cy, _
'    LR_SHARED)
'    If (bSetAsAppIcon) Then
'        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
'    End If
'    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
'
'End Sub

