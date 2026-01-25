import json
import xml.etree.ElementTree as ET

# Load weapon data
with open('weapon_data.json', 'r') as f:
    weapon_data = json.load(f)

# CORRECTED mapping of scroll item IDs to their source weapons
scroll_mapping = {
    'AxeBoomerangScroll': ('1h Axe', 'Axe'),
    'ThrowingDaggersScroll': ('Dagger', 'Dagger'),
    'ChargeScroll': ('2h sword', '2HSword'),              # FIXED: was Spear
    'SwordBlastScroll': ('2h sword', '2HSword'),          # FIXED: was sword
    'SuperSpinScroll': ('2H AXE', '2HAxe'),               # FIXED: was 2h sword
    'GroundSmashScroll': ('2H Hammer', '2HHammer'),
    'TripleSwipeScroll': ('Spear', 'Spear'),
    'GrapeshotScroll': ('Pistol', 'Pistol'),
    'RainOfArrowsScroll': ('Shortbow', 'Shortbow'),
    'ExplosiveBoltScroll': ('Crossbow', 'Crossbow'),
    'AssassinateScroll': ('Rifle', 'Rifle'),              # FIXED: was Dagger
    'MagicMissilesScroll': ('Wand', 'Wand'),
    'HammerOfFaithScroll': ('Scepter', 'Scepter'),
    'DeathRayScroll': ('Magic orb', 'Magic orb'),         # FIXED: was Tome of Secrets
    'ScorchingWaveScroll': ('power staff', 'power staff'), # FIXED: was Wand
    'FireThrowerScroll': ('power staff', 'power staff'),   # FIXED: was Rifle
    'FireballScroll': ('Tome of Secrets', 'Tome of Secrets'),  # FIXED: was Wand
    'LightningStrikeScroll': ('Tome of Secrets', 'Tome of Secrets'),  # FIXED: was Wand
    'BeeStingScroll': ('druid staff', 'druid staff'),      # FIXED: was Wand
    'TeleportationScroll': None,  # Should not have damage values
}

tree = ET.parse('modded_files/ItemDefinitions_Usables')
root = tree.getroot()

changes = []
total_updated = 0
total_removed = 0

# Find all scroll items and update them
for item_def in root.findall('ItemDefinition'):
    item_id = item_def.get('Id')
    
    # Check if this is a scroll we want to update
    if item_id not in scroll_mapping:
        continue
    
    mapping = scroll_mapping[item_id]
    
    # Special case: TeleportationScroll should not have BaseDamage elements
    if mapping is None:
        print(f"\nProcessing {item_id} (removing damage values):")
        level_variations = item_def.find('LevelVariations')
        if level_variations is not None:
            for level_elem in level_variations.findall('Level'):
                base_damage = level_elem.find('BaseDamage')
                if base_damage is not None:
                    level_id = level_elem.get('Id')
                    old_min = base_damage.get('Min')
                    old_max = base_damage.get('Max')
                    level_elem.remove(base_damage)
                    total_removed += 1
                    print(f"  Level {level_id}: Removed {old_min}-{old_max}")
        continue
    
    excel_sheet, weapon_prefix = mapping
    
    if excel_sheet not in weapon_data:
        print(f"Warning: {excel_sheet} not in weapon data")
        continue
    
    levels = weapon_data[excel_sheet]['levels']
    
    print(f"\nProcessing {item_id} (from {excel_sheet}):")

    # Scrolls use the same level mapping as base weapon variants (levels 0 to 4)
    level_mapping = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5}
    
    level_variations = item_def.find('LevelVariations')
    if level_variations is None:
        print(f"  No LevelVariations found")
        continue
    
    updated_count = 0
    for level_elem in level_variations.findall('Level'):
        level_id = level_elem.get('Id')
        try:
            level_id_int = int(level_id)
        except (ValueError, TypeError):
            continue
        
        if level_id_int not in level_mapping:
            continue
        
        excel_level = level_mapping[level_id_int]
        excel_level_str = str(excel_level)
        
        if excel_level_str not in levels:
            print(f"  Level {level_id}: Excel level {excel_level} not found")
            continue
        
        base_damage = level_elem.find('BaseDamage')
        if base_damage is None:
            continue
        
        old_min = base_damage.get('Min')
        old_max = base_damage.get('Max')
        new_min = levels[excel_level_str]['min']
        new_max = levels[excel_level_str]['max']
        
        if old_min != str(new_min) or old_max != str(new_max):
            base_damage.set('Min', str(new_min))
            base_damage.set('Max', str(new_max))
            changes.append({
                'scroll': item_id,
                'level': level_id,
                'excel_level': excel_level,
                'old': f"{old_min}-{old_max}",
                'new': f"{new_min}-{new_max}"
            })
            updated_count += 1
            total_updated += 1
            print(f"  Level {level_id} (Excel {excel_level}): {old_min}-{old_max} -> {new_min}-{new_max}")
    
    if updated_count == 0:
        print(f"  No updates needed")

tree.write('modded_files/ItemDefinitions_Usables', encoding='utf-8', xml_declaration=True)
print(f"\n✓ Updated {total_updated} scroll damage values")
print(f"✓ Removed {total_removed} damage values from TeleportationScroll")
print(f"✓ File saved")

with open('scroll_item_changes_corrected.json', 'w') as f:
    json.dump(changes, f, indent=2)
print(f"✓ Change log saved to scroll_item_changes_corrected.json")
