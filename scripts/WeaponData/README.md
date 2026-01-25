# Weapon Data Consolidation Script

## Overview

`consolidate_all_updates.py` is a single, unified Python script that handles all weapon data updates. It replaces the need to run multiple separate scripts in sequence.

## What It Does

The script performs 5 phases of data extraction and XML file updates:

### Phase 1: Extract Weapon Variant Stats
- Reads variant stat names from row 22 of each weapon sheet in Excel
- Extracts Tier 1 Variant Values (18 stats × 6 levels)
- Extracts Tier 2 Variant Values (19 stats × 6 levels)

### Phase 2: Extract Weapon Damage Data
- Reads level damage values from all weapon sheets
- Captures minimum and maximum damage for each level

### Phase 3: Build Stat Bonus Mapping
- Creates comprehensive stat name mapping (Excel names → XML stat names)
- Handles composite stats (e.g., "Health;Health Regen")
- Supports case-insensitive matching

### Phase 4: Update Weapon Stat Bonuses
- Updates BaseStatBonuses in weapon XML files:
  - `ItemDefinitions_Weapons`
  - `ItemDefinitions_DLC1`
  - `ItemDefinitions_DLC2`
- Variants 0-1: No bonuses (except WarShield always gets -20 Dodge)
- Variants 2-3: Use Tier 1 values
- Variants 4-5: Use Tier 2 values

### Phase 5: Update Scroll Item Damage
- Updates BaseDamage values for scroll items in `ItemDefinitions_Usables`
- Maps each scroll to its source weapon
- Removes damage from TeleportationScroll (special case)

## Usage

### Prerequisites

1. Place `tls_weapon_docs.xlsx` in the same directory as the script
2. Ensure the following modded files are in the correct locations:
   - `../../modded_files/ItemDefinitions_Weapons`
   - `../../modded_files/ItemDefinitions_DLC1`
   - `../../modded_files/ItemDefinitions_DLC2`
   - `../../modded_files/ItemDefinitions_Usables`

### Running the Script

From the script directory:

```bash
python consolidate_all_updates.py
```

### Output

The script prints a detailed report showing:
- Weapons and variants extracted from Excel
- Stat name mappings created
- Number of updates applied to each XML file
- Final summary of all updates

## File Dependencies

**Input Files:**
- `tls_weapon_docs.xlsx` - Excel file with weapon data and variant stats

**Modified Files:**
- `modded_files/ItemDefinitions_Weapons` - Weapon definitions
- `modded_files/ItemDefinitions_DLC1` - DLC1 weapon definitions
- `modded_files/ItemDefinitions_DLC2` - DLC2 weapon definitions
- `modded_files/ItemDefinitions_Usables` - Scroll item definitions

## Previous Scripts

This consolidated script replaces the functionality of:
- `extract_weapon_data.py` - Weapon damage extraction
- `extract_weapon_variants.py` - Variant stat extraction
- `build_stat_bonus_mapping_v2.py` - Stat name mapping
- `update_stat_bonuses_final.py` - Weapon stat bonus updates
- `update_scroll_items_corrected.py` - Scroll damage updates

These individual scripts can still be used for isolated updates if needed, but the consolidated script is recommended for complete weapon data synchronization.

## Special Cases

### WarShield
- Always receives -20 Dodge modifier at all variants (levels 0-5)
- Additional tier-specific bonuses are applied on top of this

### OffHand Weapons
- Excluded from stat bonus updates:
  - BattleMageMagicWand, BattleMageSword, DuelingPistol
  - MysticHammer, ParryingDagger, PreciseHandCrossbow
  - ReliableMagicScepter, SwiftAxe, TransferMagicOrb, WarpCrystal

### TeleportationScroll
- Has its BaseDamage removed (special case - doesn't do damage)

## Troubleshooting

If the script doesn't find expected files:
1. Verify `tls_weapon_docs.xlsx` is in the script directory
2. Check that the path to `modded_files/` is correct (should be `../../modded_files/` from script location)
3. Ensure Python 3.7+ is installed with required packages: `openpyxl`

## Notes

- The script preserves all existing XML structure except for BaseStatBonuses elements
- Each run updates the files in-place; backups are recommended before running
- Unicode characters in output may display differently depending on terminal encoding
