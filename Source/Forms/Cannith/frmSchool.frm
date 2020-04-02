VERSION 5.00
Begin VB.Form frmSchool 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collectable Schools"
   ClientHeight    =   8664
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8964
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchool.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8664
   ScaleWidth      =   8964
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8964
      _ExtentX        =   15812
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "Arcane;Cultural;Lore;Natural"
   End
   Begin CannithCrafting.userInfo usrInfo 
      Height          =   8304
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8964
      _ExtentX        =   15812
      _ExtentY        =   14647
      TitleSize       =   2
   End
End
Attribute VB_Name = "frmSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private menColorScheme As ColorGroupEnum

Private mtypInfo As InfoType
Private menSchool As SchoolEnum

Private Sub Form_Initialize()
    LoadSchoolInfo mtypInfo
End Sub

Private Sub Form_Load()
    cfg.RefreshColors Me
    menColorScheme = cgeWorkspace
    With Me.usrInfo
        .BackColor = cfg.GetColor(menColorScheme, cveBackground)
        .TextColor = cfg.GetColor(menColorScheme, cveText)
        .LinkColor = cfg.GetColor(menColorScheme, cveTextLink)
        .ErrorColor = cfg.GetColor(menColorScheme, cveTextError)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseApp
End Sub

Private Sub usrHeader_Click(pstrCaption As String)
    Dim strFreq() As String
    Dim lngTier As Long
    Dim lngFreq As Long
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim i As Long
    
    strFreq = Split(",Common:,Uncommon:,Rare:", ",")
    Me.usrInfo.Clear
    With mtypInfo.School(GetSchoolID(pstrCaption))
        For lngTier = 1 To 6
            Me.usrInfo.AddText vbNullString
            Me.usrInfo.AddTextFormatted "   Tier " & lngTier & " ", True, False, False, -1, 0
            lngHigh = lngTier * 5
            lngLow = lngHigh - 4
            If lngLow = 26 Then
                Me.usrInfo.AddTextFormatted "(Quest Level " & lngLow & "+)", False, False, False, cfg.GetColor(menColorScheme, cveTextDim), 2
            Else
                Me.usrInfo.AddTextFormatted "(Quest Level " & lngLow & " to " & lngHigh & ")", False, False, False, cfg.GetColor(menColorScheme, cveTextDim), 2
            End If
            For lngFreq = 1 To 3
                With .Tier(lngTier).Freq(lngFreq)
                    Me.usrInfo.AddText "   " & strFreq(lngFreq), 0
                    For i = 1 To .Materials
                        If i > 1 Then Me.usrInfo.AddText ",", 0
                        Me.usrInfo.AddLink .Material(i), lseMaterial, .Material(i), 0, False
                    Next
                    Me.usrInfo.AddText vbNullString
                End With
            Next
        Next
    End With
End Sub
