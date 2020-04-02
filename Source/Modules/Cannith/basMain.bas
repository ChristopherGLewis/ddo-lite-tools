Attribute VB_Name = "basMain"
Option Explicit

Public glngActiveColor As Boolean

Public Const ErrorIgnore As Long = 37001

Public PixelX As Long
Public PixelY As Long

Public xp As clsWindowsXP
Public cfg As clsConfig

Sub Main()
    Load frmMessages
    Randomize
    StopwatchInit
    Set xp = New clsWindowsXP
    Set cfg = New clsConfig
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    InitData
    Select Case LCase(Command)
        Case "scale", "scales", "scaling": OpenForm "frmScaling"
        Case "shard", "shards", "effect", "effects": OpenForm "frmShard"
        Case "material", "materials", "ingredient", "ingredients", "collect", "collectable", "collectables", "collectible", "collectibles": OpenMaterialType meCollectable
        Case "soul", "soul gem", "soul gems", "soulgem", "soulgems": OpenMaterialType meSoulGem
        Case "icon", "debug": frmIcon.Show
        Case "school", "schools": OpenSchool seArcane
        Case "gear set", "gearset", "gear sets", "gearsets": OpenForm "frmGearset"
        Case "augment", "augments": OpenForm "frmAugments"
        Case "eldritch", "ritual", "rituals", "eldritchritual", "eldritch ritual", "eldritchrituals", "eldritch rituals": OpenForm "frmEldritch"
        Case Else
            If LCase(Command) = "output" Then
                Output
            ElseIf LCase(Right(Command, 8)) = ".gearset" Then
                If OpenGearsetFile(Command) Then Exit Sub
            End If
            OpenForm "frmItem"
    End Select
End Sub

Public Sub CloseApp()
    Dim lngForms As Long
    Dim frm As Form
    
    For Each frm In Forms
        lngForms = lngForms + 1
    Next
    If lngForms < 3 Then
        Unload frmMessages
        Set cfg = Nothing
        Set xp = Nothing
    End If
End Sub

Public Function FooterClick(pstrCaption As String) As Boolean
    FooterClick = True
    Select Case pstrCaption
        Case "New Item": OpenForm "frmItem"
        Case "New", "New Gearset": OpenForm "frmGearset"
        Case "Load", "Load Gearset", "Gearset": FooterClick = OpenGearset()
        Case "Effects", "Shards": OpenForm "frmShard"
        Case "Materials": OpenForm "frmMaterial"
        Case "Augments": OpenForm "frmAugments"
        Case "Schools": OpenSchool seArcane
        Case "Scaling": OpenForm "frmScaling"
        Case Else: UnderConstruction
    End Select
End Function

Public Sub OpenForm(pstrForm As String)
    Dim frm As Form
    
    Select Case pstrForm
        Case "frmItem"
            Set frm = New frmItem
        Case "frmGearset"
            Set frm = New frmGearset
        Case "frmShard"
            Set frm = New frmShard
        Case "frmMaterial"
            Set frm = New frmMaterial
        Case "frmAugments"
            Set frm = New frmAugments
        Case "frmScaling"
            frmScaling.Show
            Exit Sub
        Case "frmHelp"
            frmHelp.Show
            Exit Sub
        Case "frmAbout"
            frmAbout.Show
            Exit Sub
        Case "frmEldritch"
            frmEldritch.Show
            Exit Sub
        Case Else
            UnderConstruction
            Exit Sub
    End Select
    frm.Show
    Set frm = Nothing
End Sub

Public Sub OpenShard(pstrShard As String)
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmShard" Then
            If frm.ShardName = pstrShard Then
                frm.Show
                Exit Sub
            End If
        End If
    Next
    Set frm = New frmShard
    frm.ShardName = pstrShard
    frm.Show
    Set frm = Nothing
End Sub

Public Sub OpenMaterial(pstrMaterial As String)
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmMaterial" Then
            If frm.Material = pstrMaterial Then
                frm.Show
                Exit Sub
            End If
        End If
    Next
    Set frm = New frmMaterial
    Load frm
    frm.Material = pstrMaterial
    frm.Show
    Set frm = Nothing
End Sub

Public Sub OpenMaterialType(penType As MaterialEnum)
    Load frmMaterial
    frmMaterial.MaterialType = penType
    frmMaterial.Show
End Sub

Public Sub OpenValue(pstrCollectable As String)
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmValueDetail" Then
            If frm.Collectable = pstrCollectable Then
                frm.Show
                Exit Sub
            End If
        End If
    Next
    Set frm = New frmValueDetail
    Load frm
    frm.Collectable = pstrCollectable
    frm.Show
    Set frm = Nothing
End Sub

Public Sub OpenSchool(penSchool As SchoolEnum)
    Load frmSchool
    frmSchool.usrHeader.SetTab penSchool - 1
    frmSchool.Show
End Sub

Public Sub OpenAugment(plngAugmentID As Long, plngVariant As Long, plngScale As Long)
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmAugments" Then
            If frm.IsMatch(plngAugmentID, plngVariant, plngScale) Then
                frm.Show
                Exit Sub
            End If
        End If
    Next
    Set frm = New frmAugments
    frm.SetAugment plngAugmentID, plngVariant, plngScale
    frm.Show
    Set frm = Nothing
End Sub

Private Function OpenGearsetFile(pstrFile As String) As Boolean
    Dim strFile As String
    Dim strPath As String
    Dim frm As Form
    
    If Len(pstrFile) < 9 Then Exit Function
    strFile = pstrFile
    If InStr(strFile, ":") = 0 And Left$(strFile, 1) <> "\" Then
        strPath = cfg.CraftingPath & "\"
        If Not xp.File.Exists(strPath & strFile) Then
            strPath = App.Path & "\"
            If Not xp.File.Exists(strPath & strFile) Then strPath = strPath & "Save\"
        End If
        strFile = strPath & strFile
    End If
    If Not xp.File.Exists(strFile) Then Exit Function
    OpenGearsetFile = True
    Set frm = New frmGearset
    frm.OpenFile strFile
    frm.Show
    Set frm = Nothing
End Function

Public Function OpenGearset() As Boolean
    Dim strFile As String
    Dim frm As Form
    
    strFile = xp.ShowOpenDialog(cfg.CraftingPath, "Gearsets (*.gearset)|*.gearset", "*.gearset")
    If Len(strFile) = 0 Then Exit Function
    OpenGearset = True
    Set frm = New frmGearset
    frm.OpenFile strFile
    frm.Show
    Set frm = Nothing
End Function

Public Sub CheckImages()
    Dim strPath As String
    Dim strFile As String
    Dim i As Long
    
    strPath = App.Path & "\Images\Collectables\"
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType <> meSoulGem Then
                strFile = strPath & Replace(.Material, "Tome:", "Tome-") & ".bmp"
                If Not xp.File.Exists(strFile) Then Debug.Print .Material
            End If
        End With
    Next
    Debug.Print "Finished"
End Sub

Public Sub WidestGridName()
    Dim frm As Form
    Dim strWidest As String
    Dim lngWidest As Long
    Dim i As Long
    
    For Each frm In Forms
        Exit For
    Next
    For i = 1 To db.Shards
        If lngWidest < frm.TextWidth(db.Shard(i).GridName) Then
            lngWidest = frm.TextWidth(db.Shard(i).GridName)
            strWidest = db.Shard(i).GridName
        End If
    Next
    Set frm = Nothing
    Debug.Print strWidest & " (" & lngWidest & ")"
End Sub

Public Function CollectableCount() As Long
    Dim lngCount As Long
    Dim i As Long
    
    For i = 1 To db.Materials
        If db.Material(i).MatType = meCollectable Then
            lngCount = lngCount + 1
            Debug.Print lngCount & ". " & db.Material(i).Material
            If lngCount Mod 5 = 0 Then Debug.Print
        End If
    Next
    CollectableCount = lngCount
End Function

Public Sub GridColors(ColorOn As Long, ColorOff As Long, SelectOn As Long, SelectOff As Long, ColorGray)
    If cfg.DarkColors Then
        ColorOn = 255
        ColorOff = 215
        SelectOn = 220
        SelectOff = 185
        ColorGray = 215
    Else
        ColorOn = 255
        ColorOff = 235
        SelectOn = 235
        SelectOff = 215
        ColorGray = 230
    End If
End Sub

Public Sub ShortNames()
'    Dim i As Long
'
'    ReDim strLine(1 To db.Shards) As String
'    For i = 1 To db.Shards
'        With db.Shard(i)
'            strLine(i) = .ShortName & "," & Me.TextWidth(.ShortName)
'        End With
'    Next
'    xp.File.SaveStringAs App.Path & "\ShortNames.csv", Join(strLine, vbNewLine)
'    Exit Sub
End Sub
