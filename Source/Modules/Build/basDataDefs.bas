Attribute VB_Name = "basDataDefs"
' Written by Ellis Dee
' How do you simulate a database in native VB6? User-Defined Type structures. Lots and lots of UDTs.
Option Explicit

Public db As DatabaseType

'This is kinda screwy, used for GrantedFeats and Enh/Des Reqs.
'TODO make Enh/Des Reqs a separate TYPE
Public Type PointerType
    Raw As String
    Style As PointerEnum  'NOTE This is NOT tree type
    Feat As Long
    Tree As Long
    Tier As Long ' This is where Level goes for granted feats
    Ability As Long
    Selector As Long
    Rank As Long ' 0=Standard rules, otherwise the specific rank required
    Destiny As Long
End Type

Public Type RaceType
    RaceID As Long
    RaceName As String
    Abbreviation As String
    Type As RaceTypeEnum
    IconicClass As ClassEnum
    SubRace As RaceEnum
    ListFirst As Boolean
    BonusFeat As Boolean
    Stats(6) As Long
    SkillPoints As Long
    GrantedFeat() As PointerType
    GrantedFeats As Long
    Tree() As String
    Trees As Long
End Type

' SpellType is a separate, unrelated, purely lookup reference table holding spell info
Public Type SpellType
    SpellName As String
    Wiki As String
    Descrip As String
    Rare As Boolean
    Cost As String
    Cooldown As String
    Crawled As Boolean
End Type

' SpellListType is the actual spell list for a given class. Spells are stored by name, which is the primary key for SpellType for lookups
Public Type SpellListType
    Spell() As String
    Spells As Byte
End Type

Public Type PactType
    PactName As String
    Spells() As String
End Type

Public Type ClassType
    ClassID As Long
    ClassName As String
    Abbreviation As String
    Initial() As String ' Initial used at top of skill chart; several choices to ensure each is unique
    Color As ColorValueEnum ' Color for Leveling Order display (only used for splash classes)
    Alignment(6) As Boolean
    SkillPoints As Long
    BAB As Long ' BABEnum
    BonusFeat(20) As Byte ' 0 = none, 1 = bonus, 2 = class
    NativeSkills(1 To 21) As Boolean
    MaxSpellLevel As Long
    SpellSlots() As Long ' 2-dimensional array: (ClassLevel, SpellLevel)
    SpellList() As SpellListType ' Index = spell level
    FreeSpell() As String
    FreeSpells As Long
    MandatorySpell() As String
    MandatorySpells As Long
    CanCastSpell(9) As Long ' 0 = healing, otherwise index = spell level
    Tree() As String
    Trees As Long
    GrantedFeat() As PointerType
    GrantedFeats As Long
    Pact() As PactType
    Pacts As Long
    DisallowedClasses As Long
    DisallowedClass() As String
    DisallowedClassID() As Long
    Trapper As Boolean
End Type

Public Type ReqListType
    Req() As PointerType
    Reqs As Long
End Type

Public Type RankType
    Rank As Long
    Req() As ReqListType ' 1 = All, 2 = One, 3 = None
    Class() As Boolean
    ClassLevel() As Long
    Cost As Long  'Cost for this rank - typically defaults to ability cost
End Type

Public Type ClassLevelType
    Class As ClassEnum
    ClassLevels As Long
End Type

'This is a Selector in an Ability
Public Type SelectorType
    SelectorName As String
    Cost As Long
    ClassBonus() As Boolean
    RaceBonus() As Boolean
    Stat As StatEnum
    StatValue As Long
    Skill As SkillsEnum
    SkillValue As Long
    Descrip As String
    Wiki As String
    All As Boolean ' Flag denoting that this selector qualifies as having taken all selectors
    Class() As Boolean
    ClassLevel() As Long
    NotClass As ClassEnum ' Only added to prevent pacts and domains that can't be lawful from taking monk levels
    Rank() As RankType
    RankReqs As Boolean
    Req() As ReqListType ' 1 = All, 2 = One, 3 = None
    Alignment(6) As Boolean
    ' Race(0) is a RaceReqEnum value which determines the meaning of Race() array:
    ' rreAny: All races allowed (no filter)
    ' rreRequired: Race() is a list of races marked 1 if allowed
    ' rreNotAllowed: Race() is a list of races marked 1 if NOT allowed
    ' rreStandard: Race() is a list of standard (not iconic) races marked 1 if EXCLUDED
    ' rreIconic: Race() is a list of iconic marked 1 if EXCLUDED
    Race() As Long
    ' Disable invalid
    Hide As Boolean
End Type

Public Type FeatType
    FeatIndex As Long ' Useful for functions that receive a type structure instead of an index
    FeatName As String
    Abbreviation As String
    SortName As String
    Group() As Boolean ' Index = FilterEnum
    SelectorStyle As SelectorStyleEnum
    Parent As PointerType ' Parent selector
    Selector() As SelectorType
    Selectors As Long
    SelectorOnly As Boolean
    RaceBonus() As Boolean
    ClassBonus() As Boolean
    ClassBonusLevel As ClassLevelType ' Only used for Spring Attack (Monk 6)
    Descrip As String
    Wiki As String
    Channel As FeatChannelEnum
    ' Flags
    Deity As Boolean
    Times As Long
    PastLife As Boolean
    Legend As Boolean
    Selectable As Boolean
    Pact As Boolean
    Domain As Boolean
    ' Prereqs
    BAB As Long
    Level As Long
    ' Race(0) is a RaceReqEnum value which determines the meaning of Race() array:
    ' rreAny: All races allowed (no filter)
    ' rreRequired: Race() is a list of races marked 1 if allowed
    ' rreNotAllowed: Race() is a list of races marked 1 if NOT allowed
    ' rreStandard: Race() is a list of standard (not iconic) races marked 1 if EXCLUDED
    ' rreIconic: Race() is a list of iconic marked 1 if EXCLUDED
    Race() As Long
    Class() As Boolean
    ClassLevel() As Long
    GrantedBy As ClassLevelType ' Only used by Child Of (Favored Soul 3)
    Alignment() As Boolean
    Skill As SkillsEnum
    SkillValue As Long
    SkillTome As Boolean
    Stat As StatEnum
    StatValue As Long
    CanCastSpell As Boolean
    CanCastSpellLevel As Long
    RaceOnly As Boolean ' This feat can only be taken as a racial bonus feat
    ClassOnly As Boolean ' This feat can only be taken as a class or class bonus feat
    ClassOnlyClasses() As Boolean ' If this feat is ClassOnly, these are the classes that can take it (Index = ClassEnum)
    ClassOnlyLevels(20) As Boolean ' If this feat is ClassOnly, these are the class levels when it can be taken
    Req() As ReqListType ' 1 = All, 2 = One, 3 = None
End Type

Public Type FeatIndexType
    FeatIndex As Long
    FeatName As String
End Type

Public Type FeatTakenType
    Times As Long
    Selector() As Boolean
    SelectorsTaken As Long
End Type

'This is an actual enhancement.
Public Type AbilityType
    AbilityName As String  'AbilityName:
    Abbreviation As String
    Ranks As Long
    Cost As Long   'This is the cost per.  See Ranks
'    Group() As Boolean ' Index = FilterEnum
    SelectorStyle As SelectorStyleEnum
    Selector() As SelectorType
    Selectors As Long
    Parent As PointerType
    Sibling() As PointerType
    Siblings As Long
    SelectorOnly As Boolean ' If TRUE, display name as [Selector]. If false, use [AbilityName]: [Selector]
    Descrip As String 'Descrip:
    ' Prereqs
    Req() As ReqListType ' 1 = All, 2 = One, 3 = None
    Rank() As RankType
    RankReqs As Boolean
'    Class() As Boolean
'    ClassLevel() As Long
End Type

Public Type TierType
    Ability() As AbilityType ' This is the Column in a row
    Abilities As Long  ' Count of abilities
End Type

'Tree - either Enhancement or Destiny
Public Type TreeType
    TreeID As Long
    TreeName As String
    Abbreviation As String
    Initial() As String ' Used for Leveling Guide
    Color As ColorValueEnum ' Color for Leveling Guide
    TreeType As TreeStyleEnum   'THIS SHOULD BE STYLE
    Stats() As Boolean
    Tier() As TierType ' This is the ROW in a tree - Tier(0) = Cores.  Array of AbilityTypes
    Tiers As Long
    Lockout As String
    Wiki As String
End Type

Public Type FeatMapType
    Lite As String
    Ron As String
    Builder As String
    FeatIndex As Long
    Selector As Long
End Type

Public Type NameChangeType
    Type As String
    OldName As String
    NewName As String
    AssignSelector As String
    Selector As Long
End Type

Public Type TemplateType
    Class As ClassEnum
    Trapping As Boolean
    Always As Boolean
    Divine As Boolean
    Caption As String
    Descrip As String
    Warning As String
    Levelups As StatEnum
    StatPoints(4, 6) As Long
End Type

Public Type EpicGrantedFeatType
    FeatName As String
    Level() As Long
End Type

Public Type DatabaseType
    Race() As RaceType
    Class() As ClassType
    Feats As Long
    Feat() As FeatType
    FeatLookup() As FeatIndexType
    FeatDisplay() As FeatIndexType
    FeatMaps As Long
    FeatMap() As FeatMapType
    FeatMapIndex As PlannerEnum
    Trees As Long
    Tree() As TreeType
    Destinies As Long
    Destiny() As TreeType  'Array of Destiny starting at 1
    Spells As Long
    Spell() As SpellType
    NameChanges As Long
    NameChange() As NameChangeType
    Templates As Long
    Template() As TemplateType
    '2021.11.10 - Epic Granted - TODO
    EpicGrantedFeats As Long
    EpicGrantedFeat() As EpicGrantedFeatType
    Loaded As Boolean
End Type

'Full Requirement type
Public Type ReqAbilityType
    FeatID As Long   'Req is either feat based or Tree/Ability based
    FeatName As String
    TreeID As Long  'Both Destiny and Enhancement
    TreeName As String
    TreeStype As TreeStyleEnum
    Tier As String
    TierID As Long
    AbilityID As Long
    AbilityName As String
    Rank As Long
    SelectorID As Long
    SelectorName As String
End Type

