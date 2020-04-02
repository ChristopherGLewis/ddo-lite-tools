VERSION 5.00
Begin VB.Form frmValueDetail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Value Details"
   ClientHeight    =   7464
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7452
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValueDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7464
   ScaleWidth      =   7452
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userInfo usrInfo 
      Height          =   7452
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7392
      _ExtentX        =   13039
      _ExtentY        =   13145
      TitleSize       =   2
      TitleIcon       =   0   'False
   End
End
Attribute VB_Name = "frmValueDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DemandType
    Shard() As Long
    Shards As Long
End Type

Private mtypDemand(1 To 6) As DemandType

Private mstrCollectable As String


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    If Not xp.DebugMode Then Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hWnd)
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If Rotation < 0 Then
        Me.usrInfo.Scroll 3
    Else
        Me.usrInfo.Scroll -3
    End If
End Sub

Public Property Get Collectable() As String
    Collectable = mstrCollectable
End Property

Public Property Let Collectable(ByVal pstrCollectable As String)
    mstrCollectable = pstrCollectable
    Me.Caption = mstrCollectable & " Value"
    ShowDetails
End Property

Private Sub usrInfo_Click(strLink As String)
    If strLink = "Chart" Then frmValueChart.Show
End Sub

Private Sub ShowIcon()
    Dim lngIndex As Long
    Dim strID As String
    
    lngIndex = SeekMaterial(mstrCollectable)
    If lngIndex Then
        strID = GetMaterialResource(lngIndex)
        SetFormIcon Me, strID
    End If
End Sub


' ************* DRAWING *************


Private Sub ShowDetails()
    Dim lngIndex As Long
    Dim strSchool As String
    
    lngIndex = SeekMaterial(mstrCollectable)
    If lngIndex = 0 Then
        Me.usrInfo.AddText mstrCollectable & " not found."
        Exit Sub
    End If
    ShowIcon
    With db.Material(lngIndex)
        Me.usrInfo.AddLink mstrCollectable, lseMaterial, mstrCollectable
        strSchool = GetSchoolName(.School)
        Me.usrInfo.AddText "Tier " & .Tier, 0
        Me.usrInfo.AddLink strSchool, lseSchool, strSchool, 0
        Me.usrInfo.AddText " (" & GetFrequencyText(.Frequency) & ")"
        If .Frequency = feRare Then
            Me.usrInfo.AddText "Value calculations do not apply to Rare collectables", 2
        Else
            Me.usrInfo.AddText "Value: " & .Value & " Essences", 0
            Me.usrInfo.AddLinkParentheses "Chart", lseCommand, "Chart", 2
            ShowFormula lngIndex
        End If
    End With
End Sub

Private Sub ShowFormula(plngIndex As Long)
    If db.Material(plngIndex).Override <> 0 Then
        ShowOverride plngIndex
        Exit Sub
    End If
    Me.usrInfo.AddText "Collectable values are determined using the formula:", 2
    Me.usrInfo.AddText "Supply * Demand * EssenceRate", 3
    ShowSupply plngIndex
    ShowDemand
    ShowEssenceRate
    ShowSummary plngIndex
    Me.usrInfo.AddText "Click", 0
    Me.usrInfo.AddLink "here", lseHelp, "Collectable_Value", 0
    Me.usrInfo.AddText " for context on what these values mean."
End Sub

Private Sub ShowOverride(plngIndex As Long)
    Me.usrInfo.AddText "This value isn't calculated by a formula, but instead has been manually set.", 2
    With db.Material(plngIndex)
        If .School = seCultural And .Tier < 4 Then
            Me.usrInfo.AddText "Cultural tiers 1 to 3 have been manually set to reflect that regardless of playstyle, you probably already have a ton of these.", 2
            Me.usrInfo.AddText "The discussion for how to value these can be found in the forum thread from", 0
            Me.usrInfo.AddLink "post 106", lseURL, "https://www.ddo.com/forums/showthread.php/485318-Collectable-Valuation-Project?p=5962467&viewfull=1#post5962467", 0
            Me.usrInfo.AddText " to", 0
            Me.usrInfo.AddLink "post 113", lseURL, "https://www.ddo.com/forums/showthread.php/485318-Collectable-Valuation-Project?p=5962846&viewfull=1#post5962846", 0
            Me.usrInfo.AddText ".", 0
        End If
    End With
End Sub

Private Sub ShowSupply(plngIndex As Long)
    Dim typFarm() As MaterialFarmType
    Dim lngFarms As Long
    Dim i As Long
    
    GetFarms db.Material(plngIndex), typFarm, lngFarms
    Me.usrInfo.AddTextFormatted "Supply", False, False, True, -1, 2
    Me.usrInfo.AddText "The number of seconds it takes to farm one specific collectable. Times are measured using an", 0
    Me.usrInfo.AddLink "Epic Challenge Runner", lseURL, "https://www.ddo.com/forums/showthread.php/482405-Epic-Challenge-Runner", 0
    Me.usrInfo.AddText " to keep all times on an even footing.", 2
    Me.usrInfo.AddText "Time to farm 1 " & mstrCollectable & ":", 2
    For i = 1 To lngFarms
        With typFarm(i)
            Me.usrInfo.AddText SecondsToTime(.Rate) & " " & .Farm.Farm & " (" & .Difficulty & ")"
        End With
    Next
    Me.usrInfo.AddText vbNullString
    Me.usrInfo.AddText mstrCollectable & " supply value: " & db.Material(plngIndex).Supply, 3
End Sub

Private Sub ShowDemand()
    Me.usrInfo.AddTextFormatted "Demand", False, False, True, -1, 2
    Me.usrInfo.AddText "A measure of how many bound shards a given collectable is used for, and how desired those shards are. An overview of the demand table can be found in ", 0
    Me.usrInfo.AddLink "this forum post", lseURL, "https://www.ddo.com/forums/showthread.php/485318-Collectable-Valuation-Project?p=5956820&viewfull=1#post5956820", 0
    Me.usrInfo.AddText ".", 2
    Me.usrInfo.AddText "Shard demand is divided into six categories, with each category being worth a number of points.", 0
    GatherDemand
    Select Case db.DemandStyle
        Case dseRaw: ShowRaw
        Case dseTop: ShowTop
        Case dseWeighted: ShowWeighted
    End Select
    Me.usrInfo.AddText vbNullString
End Sub

Private Sub ShowEssenceRate()
    Me.usrInfo.AddTextFormatted "EssenceRate", False, False, True, -1, 2
    Me.usrInfo.AddText "The average number of essences earned in one second of epic dailies. This number is somewhat arbitrary in that the formula needed a value roughly equivalent to 1/5th, and a test a run of solo epic dailies using the Epic Challenge Runner generated almost exactly 700 essences in an hour.", 2
    Me.usrInfo.AddText "700 / 3600 = 0.19444", 3
End Sub

Private Sub ShowSummary(plngIndex As Long)
    Me.usrInfo.AddTextFormatted "Summary", False, False, True, -1, 2
    Me.usrInfo.AddText "Plugging the above numbers into the formula and rounding to the nearest 25:", 2
    Me.usrInfo.AddText "Value = Supply * Demand * EssenceRate"
    With db.Material(plngIndex)
        Me.usrInfo.AddText "Value = " & .Supply & " * " & .Demand & " * " & Format(db.EssenceRate, "0.00000")
        Me.usrInfo.AddText "Value = " & Format(.Supply * .Demand * db.EssenceRate, "0.00000")
        Me.usrInfo.AddText "Value = " & .Value & " essences", 3
    End With
End Sub

Private Sub GatherDemand()
    Dim i As Long
    Dim s As Long
    
    Erase mtypDemand
    For s = 1 To db.Shards
        With db.Shard(s)
            For i = 1 To .Bound.Ingredients
                If .Bound.Ingredient(i).Material = mstrCollectable Then
                    AddDemand s, .Demand
                End If
            Next
        End With
    Next
End Sub

Private Sub AddDemand(plngShard As Long, penDemand As DemandEnum)
    With mtypDemand(penDemand)
        .Shards = .Shards + 1
        ReDim Preserve .Shard(1 To .Shards)
        .Shard(.Shards) = plngShard
    End With
End Sub

Private Sub ShowRaw()
    Dim lngTotal As Long
    Dim lngSubtotal As Long
    Dim i As Long
    Dim s As Long
    
    Me.usrInfo.AddText "The point values for all shards are then added up to reach a demand total.", 2
    For i = 1 To 6
        With mtypDemand(i)
            If .Shards Then
                Me.usrInfo.AddTextFormatted GetDemandName(i), True, False, False, -1, 0
                ShowPoints i
                Me.usrInfo.AddText ")"
                lngSubtotal = 0
                For s = 1 To .Shards
                    Me.usrInfo.AddText db.DemandValue(i) & " " & db.Shard(.Shard(s)).ShardName
                    lngSubtotal = lngSubtotal + db.DemandValue(i)
                    lngTotal = lngTotal + db.DemandValue(i)
                Next
                Me.usrInfo.AddText "Total for " & GetDemandName(i) & " shards: " & lngSubtotal, 2
            End If
        End With
    Next
    Me.usrInfo.AddText mstrCollectable & " total demand: " & lngTotal, 2
End Sub

Private Sub ShowTop()
    Dim lngTotal As Long
    Dim lngSubtotal As Long
    Dim lngCount As Long
    Dim i As Long
    Dim s As Long
    
    Me.usrInfo.AddText "The top " & db.DemandTop & " point values are then added up to reach a demand total.", 2
    lngCount = 1
    For i = 1 To 6
        With mtypDemand(i)
            If .Shards Then
                Me.usrInfo.AddTextFormatted GetDemandName(i), True, False, False, -1, 0
                ShowPoints i
                Me.usrInfo.AddText ")"
                lngSubtotal = 0
                For s = 1 To .Shards
                    If lngCount > db.DemandTop Then
                        Me.usrInfo.AddText "0 " & db.Shard(.Shard(s)).ShardName
                    Else
                        Me.usrInfo.AddText db.DemandValue(i) & " " & db.Shard(.Shard(s)).ShardName
                        lngSubtotal = lngSubtotal + db.DemandValue(i)
                        lngTotal = lngTotal + db.DemandValue(i)
                    End If
                    lngCount = lngCount + 1
                Next
                Me.usrInfo.AddText "Total for " & GetDemandName(i) & " shards: " & lngSubtotal, 2
            End If
        End With
    Next
    Me.usrInfo.AddText mstrCollectable & " total demand: " & lngTotal, 2
End Sub

Private Sub ShowWeighted()
    Dim lngTotal As Long
    Dim lngSubtotal As Long
    Dim lngCount As Long
    Dim i As Long
    Dim s As Long
    
    Me.usrInfo.AddText "The point values are then added up to reach a demand total.", 2
    Me.usrInfo.AddText "The contribution for each category is capped to a certain number of shards, which helps limit overvaluing the demand for collectables used in many low-value shards.", 2
    For i = 1 To 6
        With mtypDemand(i)
            If .Shards Then
                Me.usrInfo.AddTextFormatted GetDemandName(i), True, False, False, -1, 0
                ShowPoints i
                If db.DemandWeight(i) = 99 Then
                    Me.usrInfo.AddText ", unlimited contributions)"
                Else
                    ShowCap db.DemandWeight(i)
                End If
                lngSubtotal = 0
                lngCount = 1
                For s = 1 To .Shards
                    If lngCount > db.DemandWeight(i) Then
                        Me.usrInfo.AddText "0 " & db.Shard(.Shard(s)).ShardName
                    Else
                        Me.usrInfo.AddText db.DemandValue(i) & " " & db.Shard(.Shard(s)).ShardName
                        lngSubtotal = lngSubtotal + db.DemandValue(i)
                        lngTotal = lngTotal + db.DemandValue(i)
                    End If
                    lngCount = lngCount + 1
                Next
                Me.usrInfo.AddText "Total for " & GetDemandName(i) & " shards: " & lngSubtotal, 2
            End If
        End With
    Next
    Me.usrInfo.AddText mstrCollectable & " total demand: " & lngTotal, 2
End Sub

Private Sub ShowPoints(penDemand As DemandEnum)
    Dim lngPoints As Long
    
    lngPoints = db.DemandValue(penDemand)
    If lngPoints = 1 Then
        Me.usrInfo.AddText "(Worth 1 point", 0
    Else
        Me.usrInfo.AddText "(Worth " & lngPoints & " points", 0
    End If
    Me.usrInfo.BackupOneSpace
End Sub

Private Sub ShowCap(plngShards As Long)
    If plngShards = 1 Then
        Me.usrInfo.AddText ", capped at 1 shard)"
    Else
        Me.usrInfo.AddText ", capped at " & plngShards & " shards)"
    End If
End Sub
