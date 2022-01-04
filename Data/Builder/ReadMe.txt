Data file structure notes

This file is best viewed with word wrap turned OFF.

All Files
---------

Each line starts with a "field" name, followed by a ":", following by a space. Both the 
colon and the space are required. Leaving either out results in the line being ignored and
an error being logged.

Blank lines must be completely blank. If there's a tab or a space on an otherwise empty 
line, a somewhat confusing error is logged. It'll say "Invalid Line" but then it will
appear to not display a line with an error. That means a blank line isn't really blank.

All lists are separated by commas. (,) The space after the comma is optional. Note that this
means no list items can contain commas, which requires re-formatting some spell names.
(eg: "Aid, Mass" becomes "Mass Aid".)

Most fields are optional. The data file structure philosophy is "less is more", where the 
less data we need to put in the file the better.

Capitalization matters. "Hello" and "hello" are considered different, so be mindful of 
captialization inconsistencies.

Name:
Required. Full name. This is used when pointing to it as a prereq for something else.

Abbreviation:
Optional. Generally used for display purposes only. If not included, abbreviation is set to 
the full name.

Flags:
Flags is a generic list of boolean values. Any potential flag not included defaults to 
False, and any flags listed will be True. Different files have different valid flags.


Races
-----

Stats: #, #, #, #, #, #
Base stat list in standard order: Str, Dex, Con, Int, Wis, Cha
Legal values are 6, 8, 10

Type: 
Legal values are Free, Premium, Iconic
Defines the section of the race dropdown this race appears.

IconicClass:
Class name of the iconic's default class.

Trees:
List of racial class trees available.

Flags: BonusFeat, BonusSkill, ListFirst
Include the first two flags where appropriate for human, PDK, Half-Elf, Dragonborn, etc...
ListFirst moves the race to the top of its section in the race dropdown. Can be applied to
multiple races of a single type. (They'll be listed first, alphabetically.)

Classes
-------

Initial:
List of four abbreviations, in order:
1: Single letter initial used at the top of skill charts.
2: Two-letter abbreviation used for skill charts if two classes have the same first letter.
3: Abbreviation used for feat output. Can be up to 7 characters. (More than 7 is truncated.)
4: Three-letter abbreviation used for leveling guide if you can pick a tree from two classes
   on a single build. eg: Vanguard (Ftr) and Vanguard (Pal).

Color:
Legal values are Red, Green, Blue, Yellow, Orange, Purple
The name of the color used for the class's output in the Class Levels and Skills sections. 
Users can customize the actual color values for those six colors in Tools => Options.

Alignment:
List of alignments this class can take.

BAB:
Legal values are 1, 0.75, 0.5

SkillPoints:
Base skill points per level.

Skills:
List of skills native to this class.

GrantedFeats#:
List of feats granted to this class at class level #. List these in ascending order, but 
levels don't have to be contiguous. (Just including a single line for level 3 would be 
fine, levels 1 and 2 not needed.)

BonusFeat:
List of levels when this class gets a bonus feat. Note that bonus feats are not special to 
any particular class. For example, anyone can take toughness, plus monks can take toughness 
as a bonus feat.

ClassFeat:
List of levels when this class gets a ClassOnly feat. ClassOnly feats can only be taken as 
class feats or class bonus feats. Examples include deity feats, monk fists of 
light/darkness, druid wild shapes, rogue bonus feats, and ranger favored enemy.

Trees:
List of all enhancement trees available to this class. Do not Include global trees like 
Harper Agent. Global trees are handled separately.

HealingSpell:
The level this class is allowed to take the Empower Healing Spell feat.

MaxSpellLevel:
The highest spell level this class can cast at level 20. This must come before SpellSlots#: 
and SpellList#: in the datafile to ensure that the arrays are initialized properly before
use.

SpellSlots#:
This is basically a transposed version of the ddowiki spell slots table for each class, 
horizontal instead of vertical. Note that SpellSlots0: is ignored; it's a simple list from 
1 to 20 to make data entry/verification easier.

SpellList#:
List of all spells of the specified spell level.

MandatorySpells:
List of all mandatory spells for all spell levels in a single list. (Cleric cure spells.)

FreeSpells:
List of all free spells for all spell levels in a single list. (Artificer Repair/Inflict, 
and Druid Summon Nature's Ally.)

PactSpells: PactName, [spell list]
List where the first list item is the pact name followed by all the spells (in level order)
granted by that pact. You can include as many PactSpells: lines as you need.


Feats, Enhancements, Destinies, Spells
--------------------------------------
NOTE - for the most part, if the feats/enh/destinies/spells are NOT in correct alphabetical order
BAD things will happen.

Descrip:
Descriptions are pulled all at once from ddowiki before releasing each update. Special codes
supported include:
}  New line
{  Indent paragraph
{- Indent paragraph using - as a bullet point
Whatever character immediately following the { is used as the bullet point character. Right
now, the indent logic is to simplistically add " " + bullet character + " " to the beginning
of the first line, then each wrapped line is indented with three spaces until the paragraph 
is terminated with }.

WikiName: 
The human-readable version of the wiki link to this ability. Strip out the ddowiki.com/page/
prefix, and use spaces instead of underscores. For example, the Acid Blast spell's actual URL
and what's listed in the data file as its WikiName:
   http://ddowiki.com/page/Acid_Blast_(spell)
   Acid Blast (spell)
Only supply a WikiName if it's different than the feat or spell name.

Class:
Specify list of classes and number of class levels for each. eg: "Cleric 1, Paladin 4". 

All:
One:
None:
Optional prereq lists. Each is a comma separated list. Each list can include a mix of feats 
and abilities. (Though feats can only point to other feats.)
 - Feats take the prefix "Feat: " in the enhancements and destinies files. Do not include 
   the "Feat: " prefix in the feats file.
 - Abilities take one of the following two prefixes:
    - "Tier #: "
    - "TreeName Tier #: "
       If no treename is specified, the current tree is assumed. For example, Lay Waste:
       All: Feat: Cleave, Tier 2: Momentum Swing
 - Abilities can also take a " Rank #" suffix. For example:
   - None: Feat: Magical Training, Arcane Archer Tier 1: Energy of the Wild Rank 3, Spellsinger Tier 1: Studies: Magical Rank 3

Note: Tier 0 prereqs are not required, but this requires the Dest/Enh.txt to be in the proper order


Rank#All:
Rank#One:
Rank#None:
  - Rank-specific All/One/None requirements; all the same rules and options apply
  - Only Rank2... and Rank3... are valid options (Rank1 reqs are the ability reqs)


Selectors
---------

Feats, enhancements and destinies can have a list of subtypes, called selectors. As an
example, Improved Critical. In some cases, feats are designated as selector lists even 
though the game considers them separate feats. (eg: Least Dragonmark, Deity feats.)

Selectors use the same structure in the feats, enhancements and destinies files.

Selector:
A list of selectors for this ability. This is optional if the selector list is shared or 
exclusive. If included on a shared/exclusive selector, the names are taken as overrides. 
Swashbuckler Style is an example, where taking the II version requires the same choice as 
I, but both the I and II selector names are different and meaningful.
NOTE: Selector: should be the last line of an entry, only followed by SelectorName: lists.

SharedSelector:
Points to the ability that has the root selector list. Shared means all subsequent 
abilities can only choose selectors that have already been chosen. Magister Spell School 
bonuses is a good example: 
Spell Focus (feat): Any school
[School] Specialist: Only schools where you've taken the spell focus feat
[School] Augmentation: Must be same school as Specialist
[School] Mastery: Must be same school as Augmentation
Always point to the highest level. In the above example, none of the Magister abilities 
need their own selector lists, but instead can keep pointing down the line, re-using the 
original selector list (from the feat) for all of them. 

Parent:
Exclusive selector. Points to the Ability that has the root selector list. The opposite of 
Shared, abilities can only choose selectors that have NOT already been chosen, like Shintao
Elemental Curatives.

Siblings:
List of abilities that have a mutually exclusive selector list. Consider Shintao Curatives.
Tier 1 is the root list, and won't mention anything about shared/parent/whatever. 
Tier 2 will have Parent listed as Tier 1, with no siblings. Tier 3 will point to 1 as 
parent, and 2 as sibling. And finally, Tier 4 will have 1 as parent, and the siblings list 
will point to 2 and 3.

Flags: SelectorOnly
When this ability is chosen, the actual ability name isn't displayed. Only the selector 
name.

SelectorName:
Optional, special structure to support selectors having their own individual requirements 
list. The first SelectorName: line officially ends the data input for the core ability, so 
make sure the SelectorName list is always last. These don't define the selector names, but 
rather point to them. After each selector name line, specify selector-specific requirements 
that differ from the core ability requirements. Selector-specific requirements aren't 
additive; they fully overwrite wherever specified. As an example, consider the Exotic 
Weapon feat:

FeatName: Exotic Weapon
Group: Heroic, Melee, Range
BAB: 1
RaceBonus: Human, Purple Dragon Knight
ClassBonus: Artificer, Fighter
Selector: Bastard Sword, Dwarven Axe, Kama, Khopesh, Great Crossbow, Rpt Lt Crossbow, Rpt Hvy Crossbow, Shuriken, Handwraps
SelectorName: Bastard Sword
Stat: Strength 13
SelectorName: Dwarven Axe
Stat: Strength 13
SelectorName: Kama
ClassBonus: Fighter
SelectorName: Khopesh
ClassBonus: Fighter
SelectorName: Great Crossbow
ClassBonus: Fighter
SelectorName: Rpt Lt Crossbow
ClassBonus: Fighter
SelectorName: Rpt Hvy Crossbow
ClassBonus: Fighter
SelectorName: Shuriken
ClassBonus: Fighter

The core feat is listed as a bonus feat for fighters and artificers, and then that list 
gets overwritten to remove artificers from all except bastard sword and dwarven axe. Those 
two also get the Strength 13 requirement, which don't apply to the other types. Note that 
you don't have to supply a SelectorName line for all selectors, and the order doesn't 
matter. (The actual Exotic Weapon entry in Feats.txt is much more detailed than this example.)

Selector-specific requirements can include:
 - Stats (eg: Exotic Weapon)
 - Skills (eg: Epic Skill Focus)
 - Race (eg: Deity feats)
 - Class (eg: Skill Focus: Disable Device requires Rogue or Artificer levels)
 - ClassBonus (eg: Improved Critical)
 - All/One/None requirements lists (eg: Swashbuckling Style I)
 - Rank#All/Rank#One/Rank#None requirements lists (eg: Swashbuckling Style II)
 - Alignments (eg: Warlock Pacts and Cleric Domains)

You can also specify other selector-specific data, including:
  - WikiName
  - Descrip



Feats
-----

Group: 
Required. Multiple groups allowed, separated by commas. This is purely a courtesy field for 
letting the user filter abilities by type to make it easier to find things. Supported 
groups are hardcoded:
...Heroic, Epic, Destiny, Legend, Melee, Range, Spellcasting, Defense, Misc
Any given feat/ability can be flagged for multiple groups. Brief description of the concept 
for each:
 - Heroic, Epic, Destiny and Legend are self-explanatory.
 - Melee: Direct improvement to your personal melee weapon damage per swing. Guard effects 
   and summon bonuses don't count as melee, nor do effects that boost allies only.
 - Range: Same as melee, but for range attacks.
 - Spellcasting: Anything that improves your spellcasting (eg: metamagics), gives you more 
   spell points, makes your spells cost less, or SLAs.
 - Defense: Increase hit points, AC, PRR/MRR, saves, healing amp, or anything like that.
 - Misc: Increase to stats or skills, guard effects, summons/pets stuff, plus anything not 
   easily categorized elsewhere. (eg: Turn Undead stuff, songs, ki, etc...) 
 - Alchemy is a new grouping for Alchemists 
 - Bard Song is for bard songs


Stat: 
Ability score required. Use the format "Ability #"

Skill:
Trained ranks required. Use the format "Skill #". (Basically, this is only in there for the 
SWF feat line.)

BAB: 
BAB required.

Level:
Character level required. NOT class level. This is essentially for epic and destiny feats, 
but is also useful for the past life feat. (Which is implemented as a selector.)

Alignment:
List of compatible alignments. Do not include if there are no alignment restrictions.
(Used for Warlock Pacts and Cleric Domains.)

Race: [ListStyle], [Race], [Race], [Race], ..., [Race]
Custom list where the first list entry determines the logic applied to the list,
followed by a list of race names. The first list entry can be one of the following values:
  Required: Listed races allowed
  NotAllowed: Listed races are not allowed
  Not Iconic: All non-iconic races allowed except the listed non-iconic races 
              (Don't include any iconics in list)
  Iconic: All iconic races allowed except the listed iconic races
          (Don't include any non-iconics in the list)
While unusually complex, this structure offers a some clear advantages compared to a simple 
list of allowed races:
1) Most deity feats won't need to be updated when new races are added
2) Construct Essence feats won't need to be updated when new races are added
3) Details on the feats screen are much more human-readable. For example, Instead of a list 
   of a dozen allowed races, the Construct Essences feat will say: Races Not Allowed: 
   Warforged, Bladeforged. And Follower of Onatar will say: Not Iconic.

RaceBonus:
List of races that can take this feat as a racial bonus feat. eg: Human, Purple Dragon Knight

ClassBonus:
List of classes that can take this feat as a bonus feat.

ClassBonusLevel: [Class] [Levels]
Currently only used for Spring Attack. (Monk 6)

CanCastSpell: [SpellLevel]
This feat requires the ability to cast spells of a certain level. Most metamagics just require
the ability to cast any spell, so this value will be 1. (Can cast level 1 spells.) Heighten
requires being able to cast level 2 spells. Set the value to 0 if the requirement is to be
able to cast healing spells, which is required by Empower Healing Spell.

Repeat: #
Legal values: 1, 3 or 99
The maximum number of times this feat can be taken. Only rarely needed, most feats obey the
standard rule of selector feats can be repeated until you run out of selectors, non-selector
feats can't be repeated. Set the # to 99 for feats that have no limit on repeats. Examples
include Toughness (99), Great Ability (99), Natural Fighting (3), Least Dragonmark (1).
Note that for selector feats (eg: Great Ability), if a Repeat value is set, you can take 
the same selector multiple times. (Up to but not exceeding the Times: # specified.)

Flags: RaceOnly, ClassOnly, PastLife, Legend, Unselectable, SkillTome
 - RaceOnly: Can only be taken as a racial bonus feat. (eg: Dilettante)
 - ClassOnly: Can only be taken as a class feat. (eg: Fists of Light)
 - PastLife: Requires Hero or Legend build (eg: Past Life)
 - Legend: Requires Legend build (eg: Completionist)
 - Unselectable: Can never be selected. Allows including key granted feats of interest,
   particularly those that are prereqs for other abilities. For example, the Turn Undead 
   feat is granted to clerics and paladins and is a prereq for the turn undead destiny 
   stuff in the divine sphere, but it can't ever be chosen as a selected feat by anyone.
 - SkillTome: Skill tomes count for the Skill requirement for this feat

ClassOnlyClass:
ClassOnly feats (marked in the Flags: field) can only be taken as class bonus feats. The 
ClassOnlyClass list specifies which classes can take it. (eg: Cleric & Paladin for Deity 
feats.) Note: Requires the Flags: ClassOnly

ClassOnlyLevel:
Levels when this ClassOnly feat can be taken. This field was created almost entirely for 
Fists of Light/Darkness to differentiate them from standard monk bonus feats, but it also 
came in handy for druid Wild Shapes.

NotClass: [Class]
Special flag that tells the builder this selector can't be taken if you have any class
levels for the specified class. Only a single class can be specified. This is only used
for Warlock Pacts and Cleric Domains that can't be lawful: You can't then take monk levels.
This was added in 3.0 to properly show an error state when you have monk levels but don't
choose an alignment, and then take feats that require you not be lawful.

GrantedBy: [Class] [Levels]
Special flag that tells the builder you get this feat for free from a given class and level.
This was added in 3.0 to prevent a build with at least 3 cleric or paladin levels from 
choosing a "Child Of" deity feat if they also have 3 favored soul levels. (The classes
structure can't handle conditional granted feats, so it had to be done on the feat level.)


Trees (both enhancements and destinies)
-----
Trees have Headers then abilities
# Header:
TreeName: <name> 
Abbreviation: <Abbr>
Type: <Race, Class, RaceClass, Global, and Destiny>
Initial: Bom

Note that due to the way the program works, all abilities on the same tier for a given tree 
must have unique names. This can be an issue with core abilities, so slap a "I", "II", 
etc... onto core abilities with duplicate names.

Type:
Required. Valid options are Race, Class, RaceClass, Global, and Destiny. Even though its 
existence is superfluous in the destiny file, it is still required as a valid tree 
identifier when the program loads data.

Initial:
Required. Initials used when identifying the tree in the Leveling Guide. Include additional 
options as a comma-separated list if the first one is the same as other trees. For example:
Warchanter initials: War, WC 
Warpriest initials: War, WP
Warforged initials: War, WF

Stats:
List of stats available to the tree. This means we don't have to track stats as separate 
abilities on each tier, streamlining the data file.

Lockout: 
The one tree that's an anti-req for this tree. Used for Savants as well as AA / Elf-AA.

Tree Abilities
--------------

Tier: #
Required. Must be 0 through 6. Tier 0 refers to core abilities. Core prereqs pointing to 
the previous core are added automatically; don't include them in the datafile. 

Ranks: #
Optional. Defaults to 1 rank if not included.

Cost: #
Optional. Defaults to 1 AP (per rank) if not included.

Flags:
Selector:
SharedSelector
Parent:
Siblings:
All:
One:
None:
Rank2All:
Rank3All:
Rank3None:
SelectorName:

Spells
------
MUST BE SORTED!

Flags: Rare
Include the rare flag for rare scroll wizard and artificer spells.
