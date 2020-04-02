VERSION 5.00
Begin VB.UserControl userTable 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2880
   ScaleWidth      =   3840
End
Attribute VB_Name = "userTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrTableID As String

Private mlngBorderColor As Long
Private mlngGroupColor As Long

Private mlngMarginX As Long
Private mlngMarginY As Long
Private mlngSpace As Long
Private mlngRowHeight As Long

Private mlngTop As Long

Private tbl As TableType

Public Property Let TableID(pstrTableID As String)
    mstrTableID = pstrTableID
    DrawTable
End Property

Public Property Get TableID() As String
    TableID = mstrTableID
End Property

Public Sub DrawTable()
    SetColors
    If Not GetTable() Then Exit Sub
    SizeTable
    DrawTitle
    DrawHeaders
    DrawRows
End Sub

Private Sub SetColors()
    UserControl.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    mlngBorderColor = cfg.GetColor(cgeControls, cveBorderInterior)
    mlngGroupColor = cfg.GetColor(cgeControls, cveBorderExterior)
    UserControl.Cls
End Sub

Private Function GetTable() As Boolean
    Dim lngIndex As Long
    
    lngIndex = SeekTable(mstrTableID)
    If lngIndex Then
        tbl = db.Table(lngIndex)
        GetTable = True
    End If
    AssignGroups
End Function

Private Sub AssignGroups()
    Dim lngGroup As Long
    Dim lngCount As Long
    Dim i As Long
    
    If tbl.Group = 0 Then Exit Sub
    For i = 1 To tbl.Rows
        tbl.Row(i).Group = lngGroup
        lngCount = lngCount + 1
        If lngCount >= tbl.Group Then
            lngGroup = lngGroup + 1
            lngCount = 0
        End If
    Next
End Sub

Private Sub SizeTable()
    Dim lngLeft As Long
    Dim lngRows As Long
    Dim i As Long
    
    ' X
    mlngSpace = UserControl.TextWidth(" ")
    mlngMarginX = UserControl.ScaleX(cfg.MarginX, vbPixels, vbTwips) + mlngSpace
    For i = 1 To tbl.Columns
        With tbl.Column(i)
            .Left = lngLeft
            .Width = UserControl.TextWidth(.Widest) + mlngMarginX * 2
            .Right = .Left + .Width
            lngLeft = .Right
        End With
    Next
    UserControl.Width = tbl.Column(tbl.Columns).Right + PixelX
    ' Y
    mlngMarginY = UserControl.ScaleY(cfg.MarginY, vbPixels, vbTwips)
    mlngRowHeight = UserControl.TextHeight("Q") + mlngMarginY * 2
    lngRows = tbl.Rows
    If Len(tbl.Title) Then lngRows = lngRows + 1
    If tbl.Headers Then lngRows = lngRows + 1
    UserControl.Height = mlngRowHeight * lngRows + PixelY
End Sub

Private Sub DrawTitle()
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    mlngTop = 0
    If Len(tbl.Title) = 0 Then Exit Sub
    UserControl.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    PrintText tbl.Title, (UserControl.Width - UserControl.TextWidth(tbl.Title)) \ 2, mlngMarginY
    mlngTop = mlngTop + mlngRowHeight
End Sub

Private Sub PrintText(pstrText As String, plngX As Long, plngY As Long)
    UserControl.CurrentX = plngX
    UserControl.CurrentY = plngY
    UserControl.Print pstrText
End Sub

Private Sub DrawHeaders()
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim i As Long
    
    If Not tbl.Headers Then Exit Sub
    lngBottom = mlngTop + mlngRowHeight
    UserControl.ForeColor = cfg.GetColor(cgeDropSlots, cveText)
    UserControl.FillColor = cfg.GetColor(cgeDropSlots, cveBackground)
    For i = 1 To tbl.Columns
        With tbl.Column(i)
            UserControl.Line (.Left, mlngTop)-(.Right, lngBottom), mlngGroupColor, B
            PrintText .Header, .Left + (.Width - UserControl.TextWidth(.Header)) \ 2, mlngTop + mlngMarginY
        End With
    Next
    mlngTop = mlngTop + mlngRowHeight
End Sub

Private Sub DrawRows()
    Dim i As Long
    
    UserControl.ForeColor = cfg.GetColor(cgeControls, cveText)
    UserControl.FillColor = cfg.GetColor(cgeControls, cveBackground)
    For i = 1 To tbl.Rows
        DrawRow i
    Next
End Sub

Private Sub DrawRow(plngRow As Long)
    Dim lngCol As Long
    Dim strDisplay As String
    Dim lngLeft As Long
    
    For lngCol = 1 To tbl.Columns
        DrawBorders plngRow, lngCol
        If tbl.Column(lngCol).Style = tcseNumeric Then strDisplay = Format(tbl.Row(plngRow).Value(lngCol), "#,###") Else strDisplay = tbl.Row(plngRow).Value(lngCol)
        Select Case tbl.Column(lngCol).Style
            Case tcseTextLeft: lngLeft = tbl.Column(lngCol).Left + mlngMarginX
            Case tcseTextCenter: lngLeft = tbl.Column(lngCol).Left + (tbl.Column(lngCol).Width - UserControl.TextWidth(strDisplay)) \ 2
            Case Else: lngLeft = tbl.Column(lngCol).Right - UserControl.TextWidth(strDisplay) - mlngMarginX
        End Select
        PrintText strDisplay, lngLeft, mlngTop + mlngMarginY
    Next
    mlngTop = mlngTop + mlngRowHeight
End Sub

Private Sub DrawBorders(plngRow As Long, plngCol As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim blnThick As Boolean
    
    ' Coords
    lngLeft = tbl.Column(plngCol).Left
    lngTop = mlngTop
    lngRight = tbl.Column(plngCol).Right
    lngBottom = lngTop + mlngRowHeight
    UserControl.Line (lngLeft, lngTop)-(lngRight, lngBottom), mlngBorderColor, B
    ' Top
    If plngRow = 1 Then
        blnThick = True
    ElseIf tbl.Row(plngRow).Group <> tbl.Row(plngRow - 1).Group Then
        blnThick = True
    Else
        blnThick = False
    End If
    If blnThick Then UserControl.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), mlngGroupColor
    ' Left
    If plngCol = 1 Then UserControl.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), mlngGroupColor
    ' Right
    If plngCol = tbl.Columns Then UserControl.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), mlngGroupColor
    ' Bottom
    If plngRow = tbl.Rows Then
        blnThick = True
    ElseIf tbl.Row(plngRow).Group <> tbl.Row(plngRow + 1).Group Then
        blnThick = True
    Else
        blnThick = False
    End If
    If blnThick Then UserControl.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), mlngGroupColor
End Sub
