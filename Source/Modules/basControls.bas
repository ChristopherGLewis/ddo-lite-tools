Attribute VB_Name = "basControls"
' Written by Ellis Dee
' Generic helper functions for native controls
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Combo box constants
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_LIMITTEXT = &H141
Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_GETITEMHEIGHT = &H154

' Listbox constants
Private Const LB_SETTABSTOPS = &H192

' Textbox constants
Private Const EM_GETRECT = &HB2
Private Const EM_SETRECT = &HB3

' Declarations to put progress bar in status bar
Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

' API
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long


' ************* GENERAL *************


Public Sub SetVisible(pctl As Control, pblnVisible As Boolean)
    If pctl.Visible <> pblnVisible Then pctl.Visible = pblnVisible
End Sub


' ************* COMBOBOX  *************


Public Sub ComboAddItem(pcbo As ComboBox, pstrItem As String, plngItemData As Long)
    pcbo.AddItem pstrItem
    pcbo.ItemData(pcbo.NewIndex) = plngItemData
End Sub

Public Sub ComboClear(pcbo As ComboBox)
    pcbo.ListIndex = -1
    pcbo.Clear
End Sub

' Returns
Public Function ComboContainsText(pcbo As ComboBox, pstrText As String) As Boolean
    Dim i As Long
    
    For i = 0 To pcbo.ListCount - 1
        If pcbo.List(i) = pstrText Then
            ComboContainsText = True
            Exit Function
        End If
    Next
End Function

Public Sub ComboDropDown(cbo As ComboBox, Optional blnShow As Boolean = True)
    SendMessage cbo.Hwnd, CB_SHOWDROPDOWN, blnShow, ByVal 0&
End Sub

Public Function ComboExpand(pcbo As ComboBox, ByRef pblnOverride As Boolean) As Boolean
    Dim i As Long
    Dim lngLength As Long
    Dim lngMax As Long
    
    With pcbo
        ComboExpand = False
        lngLength = Len(.Text)
        If lngLength = 0 Then Exit Function
        For i = 0 To .ListCount - 1
            If StrComp(Left(.List(i), lngLength), .Text, vbTextCompare) = 0 Then
                pblnOverride = True
                .ListIndex = i
                pblnOverride = False
                .SelStart = lngLength
                .SelLength = Len(.Text) - lngLength
                ComboExpand = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function ComboGetText(pcbo As ComboBox) As String
    If pcbo.ListIndex <> -1 Then ComboGetText = pcbo.List(pcbo.ListIndex)
End Function

Public Function ComboGetValue(pcbo As ComboBox) As Long
    If pcbo.ListIndex = -1 Then ComboGetValue = -1 Else ComboGetValue = pcbo.ItemData(pcbo.ListIndex)
End Function

Public Function ComboIsOpen(pcbo As ComboBox) As Boolean
    ComboIsOpen = SendMessage(pcbo.Hwnd, CB_GETDROPPEDSTATE, 0, ByVal 0&)
End Function

' Return Values:
' -1 No item
' 0 Item found in list
' 1 Item not found, choosing again
' 2 Item not found, adding to list
Public Function ComboLimit(pcbo As ComboBox, ByRef pblnOverride As Boolean, Optional pblnAllowAddNew As Boolean = False) As Long
    Dim blnAddNew As Boolean
    Dim i As Long
   
    ' No item
    If Len(pcbo.Text) = 0 Then
        ComboLimit = -1
        Exit Function
    End If
    ' Item found?
    For i = 0 To pcbo.ListCount - 1
        If StrComp(pcbo.List(i), pcbo.Text, vbTextCompare) = 0 Then
            pcbo.ListIndex = i
            ComboLimit = 0
            Exit Function
        End If
    Next
    ' Item not found
    If pblnAllowAddNew Then blnAddNew = (MsgBox("Would you like to add " & Chr(34) & pcbo.Text & Chr(34) & " to the list?", vbQuestion + vbYesNo + vbDefaultButton2, "Item Not Found") = vbYes)
    ' Choose again
    If Not blnAddNew Then
        pcbo.ListIndex = -1
'        MsgBox "Please select from the list.", vbInformation, "Notice"
        pcbo.SetFocus
        ComboLimit = 1
        pblnOverride = True
        Exit Function
    Else
    ' Add new
        ComboLimit = 2
    End If
End Function

' NOTE: Does not work if the combobox is inside a container (frame, picturebox, etc...)
Public Sub ComboListHeight(pcbo As ComboBox, plngDropDownRows As Long)
   Dim typPoint As POINTAPI
   Dim typRect As RECT
   Dim lngWidth As Long
   Dim lngHeight As Long
   Dim enScaleMode As ScaleModeConstants
   Dim lngItemHeight As Long
   
   enScaleMode = pcbo.Parent.ScaleMode
   pcbo.Parent.ScaleMode = vbPixels
   lngItemHeight = SendMessage(pcbo.Hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   lngHeight = lngItemHeight * (plngDropDownRows + 2)
   GetWindowRect pcbo.Hwnd, typRect
   typPoint.X = typRect.Left
   typPoint.Y = typRect.Top
   ScreenToClient pcbo.Parent.Hwnd, typPoint
   MoveWindow pcbo.Hwnd, typPoint.X, typPoint.Y, pcbo.Width, lngHeight, True
   pcbo.Parent.ScaleMode = enScaleMode
End Sub

' NOTE: Use this if the combobox is inside a container (frame, picturebox, etc...)
Public Sub ComboListHeightChild(pcbo As ComboBox, plngDropDownRows As Long, pfrm As Form)
   Dim typPoint As POINTAPI
   Dim typRect As RECT
   Dim lngWidth As Long
   Dim lngHeight As Long
   Dim enScaleMode As ScaleModeConstants
   Dim lngItemHeight As Long
   
   enScaleMode = pfrm.ScaleMode
   pfrm.ScaleMode = vbPixels
   lngItemHeight = SendMessage(pcbo.Hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   lngHeight = lngItemHeight * (plngDropDownRows + 2)
   GetWindowRect pcbo.Hwnd, typRect
   typPoint.X = typRect.Left
   typPoint.Y = typRect.Top
   ScreenToClient pfrm.Hwnd, typPoint
   MoveWindow pcbo.Hwnd, typPoint.X, typPoint.Y, pcbo.Width, lngHeight, True
   pfrm.ScaleMode = enScaleMode
End Sub

Public Sub ComboSetMaxLength(pcbo As ComboBox, ByVal plngMaxLength As Long)
    SendMessage pcbo.Hwnd, CB_LIMITTEXT, plngMaxLength, ByVal 0&
End Sub

Public Sub ComboSetText(pcbo As ComboBox, pstrText As String)
    Dim i As Long
    
    For i = 0 To pcbo.ListCount - 1
        If pcbo.List(i) = pstrText Then
            pcbo.ListIndex = i
            Exit Sub
        End If
    Next
    pcbo.ListIndex = -1
End Sub

Public Sub ComboSetValue(pcbo As ComboBox, ByVal plngItemData As Long)
    Dim i As Long
    
    For i = 0 To pcbo.ListCount - 1
        If pcbo.ItemData(i) = plngItemData Then
            pcbo.ListIndex = i
            Exit Sub
        End If
    Next
    pcbo.ListIndex = -1
End Sub

Public Function ComboFindText(pcbo As ComboBox, pstrText As String) As Long
    Dim i As Long
    
    For i = 0 To pcbo.ListCount - 1
        If pcbo.List(i) = pstrText Then
            ComboFindText = i
            Exit Function
        End If
    Next
    ComboFindText = -1
End Function


' ************* LISTBOX  *************


Public Sub ListboxAddItem(plst As ListBox, pstrItem As String, plngItemData As Variant)
    plst.AddItem pstrItem
    plst.ItemData(plst.NewIndex) = plngItemData
End Sub

Public Sub ListboxClear(plst As ListBox)
    plst.ListIndex = -1
    plst.Clear
End Sub

Public Function ListboxContainsItem(plst As ListBox, plngItemData As Long) As Boolean
    Dim i As Long
    
    For i = 0 To plst.ListCount - 1
        If plst.ItemData(i) = plngItemData Then
            ListboxContainsItem = True
            Exit Function
        End If
    Next
End Function

Public Function ListboxContainsText(plst As ListBox, pstrText As String) As Boolean
    Dim i As Long
    
    For i = 0 To plst.ListCount - 1
        If plst.List(i) = pstrText Then
            ListboxContainsText = True
            Exit Function
        End If
    Next
End Function

Public Function ListboxGetText(plst As ListBox) As String
    If plst.ListIndex <> -1 Then ListboxGetText = plst.List(plst.ListIndex)
End Function

Public Function ListboxGetValue(plst As ListBox) As Long
    If plst.ListIndex <> -1 Then ListboxGetValue = plst.ItemData(plst.ListIndex)
End Function

Public Sub ListboxSetText(plst As ListBox, pstrText As String)
    Dim i As Long
    
    For i = 0 To plst.ListCount - 1
        If plst.List(i) = pstrText Then
            plst.ListIndex = i
            Exit Sub
        End If
    Next
    plst.ListIndex = -1
End Sub

Public Sub ListboxSetValue(plst As ListBox, plngItemData As Long)
    Dim i As Long
    
    For i = 0 To plst.ListCount - 1
        If plst.ItemData(i) = plngItemData Then
            plst.ListIndex = i
            Exit Sub
        End If
    Next
    plst.ListIndex = -1
End Sub

Public Function ListboxTabStops(plst As ListBox, plngTabs() As Long)
    Dim lngElements As Long
    
    lngElements = UBound(plngTabs) - LBound(plngTabs) + 1
    SendMessage plst.Hwnd, LB_SETTABSTOPS, lngElements, plngTabs(LBound(plngTabs))
End Function


' ************* TEXTBOX  *************


Public Sub TextboxGotFocus(ptxt As TextBox)
    ptxt.SelStart = 0
    ptxt.SelLength = Len(ptxt.Text)
End Sub

Public Sub TextboxNumericKeypress(KeyAscii As Integer, pblnAllowDash As Boolean, pblnAllowDot As Boolean, pblnAllowColon As Boolean)
    Select Case KeyAscii
        Case 48 To 57 ' 0 to 9
        Case 45: If Not pblnAllowDash Then KeyAscii = 0 ' -
        Case 46: If Not pblnAllowDot Then KeyAscii = 0 ' .
        Case 58: If Not pblnAllowColon Then KeyAscii = 0 ' :
        Case 8 ' backspace
        Case 13 ' enter
        Case Else: KeyAscii = 0
    End Select
End Sub

' Get the formatting rectangle.
Sub TextBoxGetRect(ptxt As TextBox, Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim RECT As RECT
    
    SendMessage ptxt.Hwnd, EM_GETRECT, 0, RECT
    With RECT
        Left = .Left
        Top = .Top
        Right = .Right
        Bottom = .Bottom
    End With
End Sub

' Set the formatting rectangle and refresh the control.
Sub TextBoxSetRect(ptxt As TextBox, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim RECT As RECT
    
    With RECT
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    SendMessage ptxt.Hwnd, EM_SETRECT, 0, RECT
End Sub


' ************* CHECKBOX  *************


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
        pchk.Refresh
    End If
End Function
