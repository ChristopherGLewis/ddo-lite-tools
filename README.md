# DDO Character Planner lite
Originally created by EllisDee37, with input from DDOApps.
Updates by Chris Lewis (Farog on Khyber, ChicagoChris in the forums)

## Background
These tools were created by EllisDee37, but stopped being updated in 2018-05-25.  
I've taken the original VB6 code and added the Alchemist class and will be making minor updates 
as SSG releases them.

## GitHub Repo
The code and releases are available at https://github.com/ChristopherGLewis/ddo-lite-tools
Any issues and input should be filed there.

## Release Notes

### 4.3.0 - U56-U58 Archtype updates
Changes:
    Finalized updated for L32
    Updates for Archtypes in U56 - Dark Apostate, Sacred Fist, Stormsinger
    Updates for Archtypes in U58 - Blightcaster, Dark Hunter, Acolyte of the Skin
    Updates for U57 - Imbue updates in Feats, Enhancements, and Destinies
    Updated for quests in U56-U58 (thanks @SardaofChaos)
        
Fixes:
    Feats/Enhancements/Destines:
        Thanks to @yhelm123

### 4.2.0 - U55 Isle of Dread
Changes:
    Update to Level 32
    Added L31 feats    
    
Fixes:
    Feats:
        * Deity feats fixed for Eberron based Iconics
        * Epic Arcane Blast renamed to Epic Pact Dice
    Destinies:
        * Updates to Destinies from @yhelm123

### 4.1.0 - U54 Tabaxi
Changes:
    
    Race:
        * Added Tabaxi and Tabaxi Trailblazer
        * Tabaxi issues:
            * Corner is MultiSelector but only offers dodge
            * Lucky Cat may require 9 lives per release notes.

    Feats:
        * Added Tabaxi Feline Agility
        * Removed Mass Frog (Thanks Amastris)

    Spells:
        * Added Arcane Tempest to Warlock (Thanks Amastris)

    Enhancements:
        * Fixed issue with Grace in Elf
        * Fixed issue with Favored Enemy not recognizing selector
        * fixed issue with requirements using Ranks (AA->Soul Magic)

    Destinies:
        * Updates from U53.1 & U53.2
    
    Other:
        * Fixed Perm Destiny Pts not saving if no destiny points spent 

    Compendium Updates (1.8.0):
        * Updates for U53 from SardaOfChaos
        * Fixed footer disappearing on EP Sort
        * Added Eclipse Power PL
        * Added Tabaxi/Trailblazer PL
        * Added Universal AP/Destiny Tomes
        
    Issues:  
        * Destinies with Rank costs of 2/1/1 not working correctly
        * AM Enhancements SharedSelector can't pick 2/3/4/5 spells correctly - changed
          to use spell class & desc.
        * Siblings is not working correctly

### 4.0.2 - U51 Destiny Updates - Fixes
Changes:
    
    Fixed issue where Destiny at 22 & 25 were not being saved.

    Feats
        * Removed Dire Charge since its now a Destiny enhancement
    
    Enhancements
        * Fixed issue with SharedSelector in enhancements like Divine Disciple

### 4.0.1 - U51 Destiny Updates - Fixes
Changes:
    
    Fixed Destiny Save

    Max Permanent Destiny Points to 19 (update in Tomes.txt).
    SSG notes say 18, but I've personally got 19, and don't have all 
    my EPL's...
    
    Destinies
        * Fury - minor tweaks on selectors
    
    Enhancements
        * Fixed issue with requirements that use Feat:
        * Fixed issue with Racial Req: processing

### 4.0.0 - U51 Destiny Updates
Changes:
    Updates to support U51's Destiny Updates.
        * Reformatted Destiny.txt
        * Created a new Destiny page to work with U51 3 Destiny system
        * Updated outputs to match U51
        * Destinies now only track Permanent Destiny Points, not Fate Points.  
          This makes the use of Destinies simpler since FP's are just left around
          for legacy purposes.  
        * Revamped Requirements to make them work with selectors better. This update
          works for both Destinies and Enhancements, although Destinies currently
          have no requirements outside their own tree.

    Added new spells
        * Thanks to yhelm123

    Fixed Ravager T5 Critical Rage
        * Thanks to SouCarioca 

    Notes:
        * There is no leveling guide for Destinies.  I'll see if I can implement it
          in the future - should be similar in functionality to Enhancements LG.
        * There is no tree crawler for the Destinies yet.  
        * Destiny names and descriptions are still being updated by SSG.  Expect
          issues with saved files until the SSG updates settle down.
        * There may be lots of corner case bugs.  Please enter them in the issues 
          tracker

    Compendium
        * Added Dread Sea Scrolls


### 3.5.1 - Updates to Paladin and compendium fixes
Changes:
    Updates to Paladin spells & Enhancements

    Compendium 1.7.0
        * Changed favor points to a permanent footer
        * Added a font checkbox to show all saga's on small screens
        * Added The Dryad and the Demigod Raid

### 3.5.0 - U50 Updates and addition of Granted Feats
Changes:
    Updated Enhancements, Feats and Races for u50
        * Added Horizon Walker
        * Updated Shadar-Kai and Radian Servant Enhancements
        * Added Shadar-Kai Spiked Chain Attack feat.
        * Falconry: Meticulous Weaponry now has antireq of Item Defense.
            * May need to look at other Item Defenses to add AntiReq's
    Minor updates to Enhancement tracking of Base, Racial and Universals
    Added Granted Feats
        * Added all Granted Feats to the Feats.txt and Classes.txt files
        * Added a toggle in settings to show granted feats on display
        * Changed max feats per build to 128
        * Added ability to take feat more than once
    Compendium 1.6.0
        * Resized Patron window

### 3.4.2
Changes:
    Updated how Enhancement point calcs work to allow for separate pools for RacialPL and Universal bonuses
    Recrawled Enhancement tree and fixed order of Arcanotechnician T5's 

### 3.4.1
Changes:
    Added more 48.4 enhancement changes (thanks @LrdSlvrhnd, @SardaofChaos & @Grace_ana)
    Added Destiny Tome UI element

### 3.4.0
Changes:
    Updated Builder to support Universal Tome & Destiny tome points.  
    Rev'd save version to 5
    Added 48.4 enhancement changes (thanks @LrdSlvrhnd, @SardaofChaos & @Grace_ana)
    Fixed an issue with SpellSinger T1 studies that @Grace_ana found
    Updated Compendium - easier none/6 selection on challenges
      - click on the 1st star to toggle one/none
      - click on the 5th star to toggle 5/6
      - This is in addition to clicking left/right of the stars

### 3.3.4
Changes:
    Added 4th Epic past live per circle to Compendium
    Compendium version is now 1.5.0, other versions unchanged


### 3.3.3
Changes:
    Added Alchemist, Shifter and Shifter iconic to compendium
    Updated Compendium with updates from SardaofChaos 
    Compendium version is now 1.4.0, other versions unchanged

### 3.3.2
Fixes: 
   Alchemist missing Bonus feat at 12.
   Alchemist had an extra L3 spell at L15
   Swords to Plowshares had a tab at the end of line breaking save/restore
Changes:
    The data load should now trim off tabs
Recrawled all trees.

### 3.3.1
Fix for Alchemical Studies. Alchemical Studies - X can be taken at as a Class Feat, but only 2 times per Reaction. 
Note that is required a feat rename (':' is a special character in parsing the input files) so if you reload a saved Alchemist you will have to indicate the appropriate new feat name. 
Recrawled all trees.  Updated quest info per tremlas (Thanks!)

### 3.3.0
Added Shifter race, Razorclaw Shifter iconic, and Feydark Illusionist tree.  Recrawled all trees.  Updated quest info per SardaofChaos (Thanks!)

### 3.2.4
Updated Fatesinger (U42P4).  Recrawled Destinies.

### 3.2.3
Added the new Warlock feats from U46p2.  Fix to Inquisitive "What Later?"

### 3.2.2
Updated Knight of the Chalice, Sacred Defender and Stalward Defense per U45. Pale Master and Swords to Plowshares feat per U42 patch 4. General Wiki crawl of enhancements resulting in fixes to Bladeforged and Wood Elf.

### 3.2.1
Updated Epic Destinies with changes in U42 Patch 4 

### 3.2.0
Updates for Alchemist and other Update 45 changes
