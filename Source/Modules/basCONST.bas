Attribute VB_Name = "basCONST"
Option Explicit

'Updated 2024.08.14 for U69
'** LEVEL CHANGE
'Also change
' basBuild->InitStandardFeats()
' basBuild->InitLegendFeats()
' basFormat->IdentifyChannels()
' basFormatLite->GetBuildFeatSlot()

Public Const MAX_LEVEL As Long = 34

Public Const MAX_LEVELUPS As Integer = 8

Public Const MAX_STATS As Integer = 6

'Longest destiny Name
Public Const MAX_DESTINY_NAME As String = "Divine Crusader"

'TODO New Class/Archetype, change these, then make the changes in basEnum
Public Const MAX_CLASSES As Long = 25

'Used by basOutput to align the class names
Public Const MAX_CLASS_NAME As String = "Arcane Trickster"

'General Debug flags for diagnosing load issues with trees
Public Const DEBUG_FLAG = False
Public Const DEBUG_TREE = "Duergar Mindcleaver"
Public Const DEBUG_TIER = 2
Public Const DEBUG_ABILITY = "Dwarven Weapon Training I"


