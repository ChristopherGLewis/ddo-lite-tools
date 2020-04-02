Attribute VB_Name = "basResources"
Option Explicit

Public Enum ArrowEnum
    aeDelete
    aeUp
    aeDown
End Enum

Public Enum ArrowStateEnum
    aseEnabled
    aseDisabled
    asePressed
End Enum

Public Enum GeneralIconEnum
    gieCollectables
    gieIngredients
    gieSoulGems
    gieAugments
    gieML
    gieShard ' gieShard needs to be the last enum value before gieGeneralCount
    gieGeneralCount
End Enum

Public Enum ResourceLocationEnum
    rleResourceFile
    rleDisk
    rleImport
End Enum

Public Enum GearStyleEnum
    gseIcon
    gseBitmap
    gseBitmapSelected
End Enum

Public Enum SlotStyleEnum
    sseGear
    ssePaperDoll
End Enum

Public Enum MaterialLocationEnum
    mleResourceFile
    mleDiskBitmap
    mleDiskIcon
    mleImport
End Enum

Private Function BasePath() As String
    BasePath = App.Path & "\Source\Graphics\"
End Function

Private Function IconPath() As String
    IconPath = BasePath() & "Icons\"
End Function

Public Function BitmapPath() As String
    BitmapPath = BasePath() & "Bitmaps\"
End Function

Private Function ImportPath() As String
    ImportPath = BasePath & "Import\"
End Function

Public Function GetArrowResource(penArrow As ArrowEnum, penState As ArrowStateEnum, Optional penLocation As ResourceLocationEnum) As String
    Dim strPath As String
    Dim strFile As String
    
    Select Case penLocation
        Case rleResourceFile
            GetArrowResource = GetResourceArrowID(penArrow, penState)
            Exit Function
        Case rleImport
            strPath = ImportPath()
            strFile = GetResourceArrowID(penArrow, penState)
        Case rleDisk
            strPath = BitmapPath() & "Arrows\"
            Select Case penArrow
                Case aeDelete: strFile = "Delete"
                Case aeUp: strFile = "Up"
                Case aeDown: strFile = "Down"
            End Select
            Select Case penState
                Case aseDisabled: strFile = strFile & "-Disabled"
                Case asePressed: strFile = strFile & "-Pressed"
            End Select
    End Select
    GetArrowResource = strPath & strFile & ".bmp"
End Function

Public Function GetResourceArrowID(penArrow As ArrowEnum, penState As ArrowStateEnum) As String
    Dim strArrow As String
    Dim strModifier As String
    
    Select Case penArrow
        Case aeDelete: strArrow = "DEL"
        Case aeUp: strArrow = "UP"
        Case aeDown: strArrow = "DOWN"
    End Select
    Select Case penState
        Case aseEnabled: strModifier = "E"
        Case aseDisabled: strModifier = "D"
        Case asePressed: strModifier = "P"
    End Select
    GetResourceArrowID = "ARW" & strArrow & strModifier
End Function
'
'Public Function GetGearResource(penGear As GearEnum, penStyle As GearStyleEnum, Optional penLocation As ResourceLocationEnum = rleResourceFile) As String
'    Dim strPath As String
'    Dim strFile As String
'    Dim strExt As String
'
'    Select Case penLocation
'        Case rleResourceFile
'            GetGearResource = GetResourceGearID(penGear, penStyle)
'            Exit Function
'        Case rleImport
'            strFile = GetResourceGearID(penGear, penStyle)
'            strPath = ImportPath()
'        Case rleDisk
'            strFile = GetGearName(penGear)
'            Select Case penStyle
'                Case gseIcon: strPath = IconPath() & "Gear\"
'                Case gseBitmap: strPath = BitmapPath() & "Gear\34x34\"
'                Case gseBitmapSelected: strPath = BitmapPath() & "Gear\34x34 Selected\"
'            End Select
'    End Select
'    If penStyle = gseIcon Then strExt = ".ico" Else strExt = ".bmp"
'    GetGearResource = strPath & strFile & strExt
'End Function
'
'Private Function GetResourceGearID(penGear As GearEnum, penStyle As GearStyleEnum) As String
'    Dim strGear As String
'
'    GetResourceGearID = "UNKNOWN"
'    strGear = GetGearName(penGear)
'    If InStr(strGear, " Armor") Then strGear = Replace(strGear, " Armor", vbNullString)
'    strGear = UCase$(Replace(strGear, " ", vbNullString))
'    Select Case penStyle
'        Case gseIcon: strGear = "GRI" & strGear
'        Case gseBitmap: strGear = "GRB" & strGear
'        Case gseBitmapSelected: strGear = "GRS" & strGear
'    End Select
'    GetResourceGearID = strGear
'End Function

Public Function GetSlotResource(penSlot As SlotEnum, Optional penLocation As ResourceLocationEnum = rleResourceFile) As String
    Dim strPath As String
    Dim strFile As String
    
    Select Case penLocation
        Case rleResourceFile
            GetSlotResource = "SLT" & UCase$(GetSlotName(penSlot))
            Exit Function
        Case rleImport
            strFile = "SLT" & UCase$(GetSlotName(penSlot))
            strPath = ImportPath()
        Case rleDisk
            strFile = GetSlotName(penSlot)
            strPath = BitmapPath() & "Gear\Slots\"
    End Select
    GetSlotResource = strPath & strFile & ".bmp"
End Function

Public Function GetItemResource(pstrItemName As String, Optional penLocation As MaterialLocationEnum = mleResourceFile) As String
    Dim lngIndex As Long
    Dim strPath As String
    Dim strFile As String
    
    lngIndex = SeekItem(pstrItemName)
    If penLocation = mleResourceFile Then
        If lngIndex = 0 Then strFile = "UNKNOWN" Else strFile = db.Item(lngIndex).ResourceID
        GetItemResource = strFile
    Else
        If lngIndex = 0 Then
            strPath = BitmapPath()
            strFile = "Unknown.bmp"
        Else
            Select Case db.Item(lngIndex).ItemType
                Case iteArmor: strPath = BitmapPath() & "Gear\Armor\"
                Case iteShield, iteOrb, iteRunearm: strPath = BitmapPath() & "Gear\Shields\"
                Case iteWeapon: strPath = BitmapPath() & "Gear\Weapons\"
                Case iteAccessory: strPath = BitmapPath() & "Gear\Accessories\"
            End Select
            strFile = db.Item(lngIndex).Image & ".bmp"
        End If
        GetItemResource = strPath & strFile
    End If
End Function

Public Function GetMaterialResource(plngMaterialID As Long, Optional penLocation As MaterialLocationEnum = mleResourceFile) As String
    Dim strPath As String
    Dim strFile As String
    Dim strExt As String
    
    If penLocation = mleResourceFile Then
        GetMaterialResource = GetResourceMaterialID(plngMaterialID)
        Exit Function
    ElseIf penLocation = mleImport Then
        strFile = GetResourceMaterialID(plngMaterialID)
        strPath = ImportPath()
        strExt = ".ico"
    Else
        strFile = db.Material(plngMaterialID).Material
        RemoveText strFile, "Soul Gem: Strong "
        RemoveText strFile, "Soul Gem: "
        RemoveText strFile, ":", "-"
        If penLocation = mleDiskIcon Then
            strPath = IconPath()
            strExt = ".ico"
        Else
            strPath = BitmapPath()
            strExt = ".bmp"
        End If
        Select Case db.Material(plngMaterialID).MatType
            Case meCollectable: strPath = strPath & "Collectables\"
            Case meMisc: strPath = strPath & "Ingredients\"
            Case meSoulGem: strPath = strPath & "Soul Gems\"
        End Select
    End If
    GetMaterialResource = strPath & strFile & strExt
End Function

Private Function RemoveText(pstrText As String, pstrRemove As String, Optional pstrReplace As String) As String
    If InStr(pstrText, pstrRemove) Then pstrText = Replace(pstrText, pstrRemove, pstrReplace)
End Function

Private Function GetResourceMaterialID(plngIndex As Long) As String
    Dim strText As String
    Dim strWord() As String
    Dim lngWords As Long
    Dim strPrefix As String
    Dim lngLen As Long
    Dim i As Long
    
    GetResourceMaterialID = "UNKNOWN"
    Select Case db.Material(plngIndex).MatType
        Case meCollectable: strPrefix = "COL"
        Case meSoulGem: strPrefix = "SG"
        Case meMisc: strPrefix = "MSC"
    End Select
    strText = UCase$(db.Material(plngIndex).Material)
    RemoveText strText, "'"
    RemoveText strText, ","
    RemoveText strText, " OF ", " "
    RemoveText strText, " A ", " "
    RemoveText strText, " AND ", " "
    RemoveText strText, " THE ", " "
    RemoveText strText, "SOUL GEM: STRONG "
    RemoveText strText, "SOUL GEM: "
    RemoveText strText, ":"
    strWord = Split(strText, " ")
    lngWords = UBound(strWord)
    If lngWords > 3 Then
        lngWords = 3
        ReDim Preserve strWord(lngWords)
    End If
    Select Case lngWords
        Case 0: lngLen = 9
        Case 1: lngLen = 4
        Case 2: lngLen = 3
        Case 3: lngLen = 2
    End Select
    For i = 0 To UBound(strWord)
        strWord(i) = Left$(strWord(i), lngLen)
    Next
    GetResourceMaterialID = strPrefix & Join(strWord, vbNullString)
End Function

Public Function GetGeneralResource(penGeneral As GeneralIconEnum, Optional plngLevel As Long = 1, Optional penLocation As MaterialLocationEnum = mleResourceFile) As String
    Dim strPath As String
    Dim strFile As String
    Dim strExt As String
    
    If penLocation = mleResourceFile Then
        GetGeneralResource = GetResourceGeneralID(penGeneral, plngLevel, True)
        Exit Function
    ElseIf penLocation = mleImport Then
        strFile = GetResourceGeneralID(penGeneral, plngLevel, True)
        strPath = ImportPath()
        strExt = ".ico"
    Else
        strFile = GetResourceGeneralID(penGeneral, plngLevel, False)
        If penLocation = mleDiskIcon Then
            strPath = IconPath()
            strExt = ".ico"
        Else
            strPath = BitmapPath()
            strExt = ".bmp"
        End If
        Select Case penGeneral
            Case gieAugments, gieCollectables, gieIngredients, gieSoulGems: strPath = strPath & "Bags\"
            Case gieML, gieShard: strPath = strPath & "Ingredients\"
        End Select
    End If
    GetGeneralResource = strPath & strFile & strExt
End Function

Private Function GetResourceGeneralID(penGeneral As GeneralIconEnum, plngLevel, pblnResourceID As Boolean) As String
    Dim strReturn As String
    
    If pblnResourceID Then
        Select Case penGeneral
            Case gieCollectables: strReturn = "BAGCOLLECT"
            Case gieIngredients: strReturn = "BAGINGRED"
            Case gieSoulGems: strReturn = "BAGSOULGEM"
            Case gieAugments: strReturn = "BAGAUGMENT"
            Case gieML: strReturn = "SHARDML"
            Case gieShard: strReturn = "SHARD" & Format(plngLevel, "000")
        End Select
    Else
        Select Case penGeneral
            Case gieCollectables: strReturn = "Collectables"
            Case gieIngredients: strReturn = "Ingredients"
            Case gieSoulGems: strReturn = "Soul Gems"
            Case gieAugments: strReturn = "Augments"
            Case gieML: strReturn = "ShardML"
            Case gieShard: strReturn = "Shard" & Format(plngLevel, "000")
        End Select
    End If
    GetResourceGeneralID = strReturn
End Function

