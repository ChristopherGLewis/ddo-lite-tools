VERSION 5.00
Begin VB.Form frmEldritch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eldritch Rituals"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10212
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEldritch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6936
   ScaleWidth      =   10212
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userInfo usrInfo 
      Height          =   6312
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   9552
      _ExtentX        =   16849
      _ExtentY        =   11134
      TitleSize       =   2
   End
End
Attribute VB_Name = "frmEldritch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    cfg.RefreshColors Me
    DrawDetail
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = 3 Else lngValue = -3
    If IsOver(Me.usrInfo.hwnd, Xpos, Ypos) Then Me.usrInfo.Scroll lngValue
End Sub

Private Sub DrawDetail()
    Me.usrInfo.ClearContents
    Me.usrInfo.AddText "You can apply an Eldritch Ritual to any item using the Stone of Change. First the item must be"
    Me.usrInfo.AddLink "Bound and Attuned", lseURL, "https://ddowiki.com/page/Stone_of_Change_-_Recipes#Binding_and_Attuning_of_Items", 0
    Me.usrInfo.AddText " by combining it with 400", 0
    Me.usrInfo.AddLink "Khyber Dragonshard Fragments", lseMaterial, "Khyber Dragonshard Fragment", 2
    Me.usrInfo.AddText "An item can only have one type of alchemical ritual place on it at a time, Adamantine or Eldritch. Applying another alchemical ritual will replace the previous one, or an already-applied Adamantine ritual.", 2
    DrawRecipe "Force Damage"
    DrawRecipe "Force Critical"
    DrawRecipe "Resistance"
    DrawRecipe "Alchemical Armor"
    DrawRecipe "Alchemical Shield"
    Me.usrInfo.AddText vbNullString
    DrawRecipe "Adamantine Ritual I"
    DrawRecipe "Adamantine Ritual II"
    DrawRecipe "Adamantine Ritual III"
    DrawRecipe "Adamantine Ritual IV"
    DrawRecipe "Adamantine Ritual V"
    Me.usrInfo.AddText vbNullString
    Me.usrInfo.AddText "See also", 0
    Me.usrInfo.AddLink "this wiki page", lseURL, "https://ddowiki.com/page/Stone_of_Change_-_Recipes#Alchemical_Rituals", 0
    Me.usrInfo.AddText " for more info."
End Sub

Private Sub DrawRecipe(pstrRitual As String)
    Dim lngIndex As Long
    Dim strItems As String
    Dim i As Long
    
    lngIndex = SeekRitual(pstrRitual)
    If lngIndex Then
        With db.Ritual(lngIndex)
            Me.usrInfo.AddTextFormatted .RitualName, True
            strItems = MakeItemList(.ItemType, .ItemStyle)
            If Len(strItems) Then Me.usrInfo.AddText "Applies to: " & strItems
            Me.usrInfo.AddText .Descrip, 0
            AddRecipeToInfo .Recipe, Me.usrInfo
            Me.usrInfo.AddText vbNullString, 2
        End With
    Else
        Me.usrInfo.AddError pstrRitual & " not found.", 2
    End If
End Sub

Private Function MakeItemList(ptypType() As Boolean, ptypStyle() As Boolean) As String
    Dim strReturn As String
    Dim i As Long
    
    If ptypType(0) Then
        For i = 1 To iteItemTypes - 1
            If ptypType(i) Then
                strReturn = strReturn & ", " & GetItemTypeName(i)
            End If
        Next
    End If
    If ptypStyle(0) Then
        For i = 1 To iseItemStyles - 1
            If ptypStyle(i) Then
                strReturn = strReturn & ", " & GetItemStyleName(i)
            End If
        Next
    End If
    If Len(strReturn) Then strReturn = Mid$(strReturn, 3) Else strReturn = "Any Item"
    MakeItemList = strReturn
End Function
