Attribute VB_Name = "basDataDefs"
' Definitions for all enumerations and data structures for crafting and gearsets.
' I was getting "circular reference between modules" errors so I just dumped
' them all in here.
' NOTE: QuickMatch is self-contained so its definitions are in basQuickMatch.bas
Option Explicit

Public db As DatabaseType

Public gtypOpenItem As OpenItemType

Public Const GearMax As Long = 20
Public Const SpotKeyMax As Long = 50
Public Const ItemScales As Long = 4


' ************* ENUMERATIONS *************


Public Enum GearEnum
    geUnknown = -1
    geHelmet = 0
    geGoggles
    geNecklace
    geCloak
    geBracers
    geGloves
    geBelt
    geBoots
    geRing
    geTrinket
    ge2hMelee
    ge1hMelee
    geHandwraps
    geRange
    geShield
    geOrb
    geRunearm
    geMetalArmor
    geLeatherArmor
    geClothArmor
    geDocent
    geGearCount
End Enum

Public Enum GroupEnum
    greUnknown
    greStat
    greDPS
    greTactic
    greWeapon
    greBane
    greStatDamage
    greDefense
    greSave
    greResist
    greAbsorb
    greGuard
    greDC
    greSpellpower
    greLore
    greMisc
    greNonScaling
    greSkill
End Enum

Public Enum MaterialEnum
    meUnknown
    meCollectable
    meSoulGem
    meMisc
End Enum

Public Enum SchoolEnum
    seAny
    seArcane
    seCultural
    seLore
    seNatural
End Enum

Public Enum FrequencyEnum
    feCommon = 1
    feUncommon = 2
    feRare = 3
End Enum

Public Enum DemandEnum
    deUniversal = 1
    deCommon
    deUncommon
    deNiche
    deObsolete
    deRarelyUsed
End Enum

Public Enum DemandStyleEnum
    dseRaw
    dseTop
    dseWeighted
End Enum

Public Enum WorldEnum
    weEberron
    weRealms
End Enum

Public Enum ItemTypeEnum
    iteUnknown
    iteWeapon
    iteArmor
    iteShield
    iteOrb
    iteRunearm
    iteAccessory
    iteEmpty
    iteItemTypes
End Enum

Public Enum ItemStyleEnum
    iseUnknown
    iseMelee2H
    iseMelee1H
    iseRange
    iseThrower
    iseMetalArmor
    iseLeatherArmor
    iseClothArmor
    iseDocent
    iseShield
    iseOrb
    iseRunearm
    iseClothing
    iseJewelry
    iseEmpty
    iseItemStyles
End Enum

Public Enum ItemSlotStyleEnum
    isseNone
    isseRed
    isseBlue
    isseGreen
    isseYellow
    isseColorless
    isseDual
End Enum

' Gearset slots
Public Enum SlotEnum
    seUnknown = -1
    seHelmet = 0
    seGoggles
    seNecklace
    seCloak
    seBracers
    seGloves
    seBelt
    seBoots
    seRing1
    seRing2
    seTrinket
    seArmor
    seMainHand
    seOffHand
    seSlotCount
End Enum

' Mapping Key, 17 gear * 3 spots per gear = 51 possible spots
Public Enum SpotEnum
    spotUnused = -1
    spotHelmet = 0
    spotGoggles = 3
    spotNecklace = 6
    spotCloak = 9
    spotBracers = 12
    spotGloves = 15
    spotBelt = 18
    spotBoots = 21
    spotRing = 24
    spotTrinket = 27
    spotArmor = 30
    spotMelee2H = 33
    spotMelee1H = 36
    spotShield = 39
    spotRange = 42
    spotRunearm = 45
    spotOrb = 48
End Enum

Public Enum AffixEnum
    aePrefix
    aeSuffix
    aeExtra
    aeEldritch
End Enum

Public Enum MainHandEnum
    mheMelee
    mheRange
End Enum

Public Enum OffHandEnum
    oheMelee
    oheShield
    oheOrb
    oheRunearm
    oheEmpty
End Enum

Public Enum ArmorMaterialEnum
    ameMetal
    ameLeather
    ameCloth
    ameDocent
End Enum

Public Enum ValueListEnum
    vleActive
    vleAll
    vleFiltered
End Enum


' ************* AUGMENTS *************


Public Enum AugmentColorEnum
    aceAny
    aceRed
    aceOrange
    acePurple
    aceBlue
    aceGreen
    aceYellow
    aceColorless
End Enum

Public Enum AugmentVendorEnum
    aveUnknown
    aveCollector
    aveGianthold
    aveLahar
    aveLahar5
End Enum

Public Type AugmentItemType
    Augment As String
    Items As Long
    Item() As String
End Type

Public Type AugmentVendorType
    Vendor As String
    Style As AugmentVendorEnum
    Color As AugmentColorEnum
    Fast As String
    AnyLevel As Boolean
    ML As Long
    Cost As String
    Location As String
End Type

Public Type AugmentScaleType
    ML As Long
    Prefix() As String
    Value As String
    Store As Long
    Remnants As Long
    RemnantsBonusDays As Boolean
    Vendors As Long
    Vendor() As AugmentVendorEnum
End Type

Public Type AugmentType
    AugmentName As String
    Color As AugmentColorEnum
    Static As Boolean
    Named As Boolean
    ResourceID As String
    Descrip As String
    Notes As String
    PrefixNotValue As Boolean
    Variations As Long
    Variation() As String
    StoreMissing() As Boolean
    Wiki() As String
    Scalings As Long
    Scaling() As AugmentScaleType
End Type


' ************* CRAFTING *************


Public Type FarmType
    Farm As String
    Wiki As String
    Any As Long
    Arcane As Long
    Lore As Long
    Natural As Long
    Seconds As Long
    Realms As Boolean
    TreasureBag As Boolean
    Video As String
    Need As String
    Fight As String
    Notes As String
End Type

Public Type TierFarmType
    Farm As String
    Difficulty As String
End Type

Public Type WorldType
    Pool As Long
    Fastest As Long
End Type

Public Type PoolType
    World(1) As WorldType
End Type

Public Type TierType
    Freq(1 To 3) As PoolType
    TierFarms As Long
    TierFarm() As TierFarmType
End Type

Public Type SchoolType
    Tier(1 To 6) As TierType
End Type

Public Type MaterialType
    MatType As MaterialEnum
    School As SchoolEnum
    Tier As Long
    Frequency As FrequencyEnum
    Supply As Long
    Demand As Long
    DemandMatrix(1 To 6) As Long
    Value As Long
    Override As Long
    Eberron As Boolean
    Material As String
    Plural As String
    Descrip As String
    Notes As String
End Type

Public Type IngredientType
    Count As Long
    Material As String
    Frequency As Long
End Type

Public Type RecipeType
    Level As Long
    Essences As Long
    Ingredients As Long
    Ingredient() As IngredientType
End Type

Public Type ShardType
    ShardName As String
    Abbreviation As String
    ShortName As String ' VERY short; used for first tab of gearset screen (eg: iPRR)
    GridName As String
    Group As String
    Bound As RecipeType
    Unbound As RecipeType
    Prefix(GearMax) As Boolean
    Suffix(GearMax) As Boolean
    Extra(GearMax) As Boolean
    ScaleName As String
    ML As Long
    Demand As DemandEnum
    Notes As String
    Descrip As String
    Warning As String
End Type

Public Type RitualType
    RitualName As String
    Recipe As RecipeType
    Descrip As String
    ItemType(iteItemTypes - 1) As Boolean
    ItemStyle(iseItemStyles - 1) As Boolean
End Type

Public Type ScalingType
    Order As Long
    Group As String
    ScaleName As String
    Table() As String
End Type

Public Type ItemType
    ItemName As String
    ItemType As ItemTypeEnum
    ItemStyle As ItemStyleEnum
    TwoHand As Boolean
    Image As String ' Filename of image
    ResourceID As String ' ID of icon in resource file
    Scales As Boolean ' Is this a scaling item? (armor/shield)
    Scalings As Long
    Scaling() As String
    SlotStyles As Long
    SlotStyle() As ItemSlotStyleEnum
End Type

Public Type ChoicePair
    Choice As String
    OffhandPair As String
End Type

Public Type ChoiceType
    Count As Long
    List() As ChoicePair
    Default As ChoicePair
End Type

Public Type DatabaseType
    Materials As Long
    Material() As MaterialType
    Farms As Long
    Farm() As FarmType
    School(4) As SchoolType
    Scales As Long
    Scaling() As ScalingType
    Groups As Long
    Group() As String
    Items As Long
    Item() As ItemType
    Melee1H As ChoiceType
    Melee2H As ChoiceType
    Range As ChoiceType
    Augments As Long
    Augment() As AugmentType
    AugmentItems As Long
    AugmentItem() As AugmentItemType
    AugmentVendors As Long
    AugmentVendor() As AugmentVendorType
    Rituals As Long
    Ritual() As RitualType
    Shards As Long
    Shard() As ShardType
    ShardIndex() As Long
    Frequency() As Double ' % chance a dispenser will drop common / uncommon / rare
    Backpack() As Double ' % chance an Any displenser will drop arcane / culture / lore / natural
    EssenceRate As Double ' Average essences per second from solo epic dailies
    DemandStyle As DemandStyleEnum
    DemandValue() As Long ' Value of each demand category: (1) = Universal, (2) = Common, ..., (6) = Rarely Used
    DemandWeight() As Long ' How many values from each category can be used when using Weighted style
    DemandTop As Long ' How many values to use when using Top style
End Type


' ************* GRID DEFS *************


' The slot icons at the top of the grid
Public Type SlotType
    GearsetSlot As SlotEnum ' Index in gs.Item()
    Gear As GearEnum
'    Left As Long
    IconLeft As Long
End Type

' Used by both rows and columns, this array holds a list of all valid choices
Public Type ValueList
    Count As Long
    Value() As Long
End Type

' Gear slots
Public Type ColType
    Slot As Long ' Index in grid.Slot()
    Item As Long ' Index in gs.Item()
    Affix As AffixEnum
    Header As String
    LeftThick As Boolean
    RightThick As Boolean
    Effect(2) As ValueList ' (0) = Actually used, (1) = All, (2) = Filtered
    RowSelected As Long ' Effect chosen for this slot, or 0 if none chosen
End Type

' Effects
Public Type RowType
    Effect As Long ' Index in gs.Effect()
    Shard As Long ' Index in db.Shard()
    Caption As String ' db.Shard().GridName
    TopThick As Boolean
    BottomThick As Boolean
    ScaleOrder As Long ' db.Scaling().Order
    ScaleGroup As String ' db.Scaling().Group
    Spot(2) As ValueList ' (0) = Actually used, (1) = All, (2) = Filtered; Copy either 1 or 2 to 0 before using
    ColSelected As Long ' Slot chosen for this effect, or 0 if none chosen
End Type

Public Type CellType
    Active As Boolean ' TRUE if can be selected
    Selected As Boolean ' TRUE if actually selected
End Type

' These all used to be separate module-level variables in frmGearset.frm
' Now frmGearset.fm just declares Private grid As GridType
Public Type GridType
    RowHeight As Long
    TextOffsetY As Long
    SlotWidth As Long
    ColWidth As Long
    EffectWidth As Long
    HeaderHeight As Long
    Slots As Long
    Slot() As SlotType
    Rows As Long
    Row() As RowType
    Cols As Long
    Col() As ColType
    Cell() As CellType
    CurrentRow As Long
    CurrentCol As Long
    CurrentSlot As Long
    IconTop As Long
    IconWidth As Long
    IconHeight As Long
    Initialized As Boolean
End Type


' ************* GEARSET DEFS *************


Public Type AugmentSlotType
    Exists As Boolean
    Augment As Long
    Variation As Long
    Scaling As Long ' Calculated on the fly instead of stored
    Done As Boolean
End Type

Public Type ItemSlotType
    Crafted As Boolean
    ItemStyle As String
    Named As String
    ML As Long
    MLDone As Boolean
    Effect(2) As Long ' 0 = Prefix, 1 = Suffix, 2 = Extra
    EffectDone(2) As Boolean
    Augment(1 To 7) As AugmentSlotType
    EldritchRitual As Long
    EldritchDone As Boolean
    Gear As GearEnum
    SpotKey As SpotEnum ' Points to this item's base location is ItemSpot() and ItemSpotKey()
End Type                         ' Map(SpotKey) = Prefix, Map(SpotKey+1) = Suffix, Map(SpotKey+2) = Extra

Public Type GearsetType
    Item() As ItemSlotType ' 0 to seSlotCount - 1
    BaseLevel As Long
    Armor As ArmorMaterialEnum
    Mainhand As MainHandEnum
    Offhand As OffHandEnum
    TwoHanded As Boolean
    ArmorName As String
    MainhandName As String
    OffhandName As String
    Effects As Long
    Effect() As Long
    Analyzed As Boolean
    Notes As String
End Type

Public Type OpenItemType
    hwnd As Long ' Gearset form that called, or 0 if not connected (item swaps are connected)
    ML As Long
    Crafted As Boolean
    Named As String
    Gear As GearEnum
    Slot As SlotEnum
    Prefix As Long
    Suffix As Long
    Extra As Long
    Augment(1 To 7) As AugmentSlotType
    EldritchRitual As Long
    Mainhand As String
    Offhand As String
    Armor As String
End Type


' ************* ANALYSIS DEFS *************


Public Type EffectSpotType
    Current As Long ' The current spot being analyzed
    Spots As Long
    Spot() As SpotEnum
    ValidSpot() As Boolean ' All the spots this effect can go
End Type

Public Type AnalysisType
    ItemSpotKey() As Byte ' Main loop starts by copying ItemSpotKey() over ItemSpot()
    ItemSpot() As Byte      ' Use Bytes to speed this up (less memory copied = faster copy)
    Effect() As EffectSpotType
End Type

