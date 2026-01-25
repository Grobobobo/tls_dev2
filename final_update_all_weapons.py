import xml.etree.ElementTree as ET
import json

# Load Excel data
with open('weapon_data.json', 'r') as f:
    weapon_data = json.load(f)

# Mapping from Excel sheet names to weapon id prefixes
sheet_to_prefixes = {
    'sword': ['Sword'],
    'Hammer': ['Hammer'],
    '1h Axe': ['Axe'],
    'Dagger': ['Dagger', 'ParryingDagger'],
    '2h sword': ['2HSword'],
    '2H Hammer': ['2HHammer'],
    '2H AXE': ['2HAxe'],
    'Spear': ['Spear'],
    'Hand crossbow': ['HandCrossbow', 'PreciseHandCrossbow'],
    'Crossbow': ['Crossbow'],
    'Pistol': ['Pistol', 'DuelingPistol'],
    'Shortbow': ['Shortbow'],
    'Longbow': ['Longbow'],
    'Rifle': ['Rifle'],
    'Wand': ['MagicWand', 'BattleMageMagicWand'],
    'Scepter': ['MagicScepter', 'ReliableMagicScepter'],
    'Tome of Secrets': ['TomeOfMagic'],
    'Magic orb': ['MagicOrb', 'TransferMagicOrb'],
    'power staff': ['MagicStaff'],
    'druid staff': ['DruidicStaff'],
}

tree = ET.parse('modded_files/ItemDefinitions_Weapons')
root = tree.getroot()

changes = []
total_updated = 0

for sheet_name, prefixes in sheet_to_prefixes.items():
    if sheet_name not in weapon_data:
        continue
    
    levels = weapon_data[sheet_name]['levels']
    
    for item_def in root.findall('ItemDefinition'):
        item_id = item_def.get('Id')
        if not item_id:
            continue
        
        # Check if matches
        matches = False
        for prefix in prefixes:
            if item_id.startswith(prefix):
                matches = True
                break
        
        if not matches:
            continue
        
        # Skip offhand
        hands_elem = item_def.find('Hands')
        if hands_elem is not None and 'Offhand' in hands_elem.text:
            continue
        
        level_variations = item_def.find('LevelVariations')
        if level_variations is None:
            continue
        
        is_base = item_id[-1] == '0'
        level_mapping = {0: -1, 1: 0, 2: 1, 3: 2, 4: 3, 5: 4} if is_base else {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5}
        
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
                    'weapon': item_id,
                    'level': level_id,
                    'old': f"{old_min}-{old_max}",
                    'new': f"{new_min}-{new_max}"
                })
                updated_count += 1
                total_updated += 1
        
        if updated_count > 0:
            print(f"{item_id}: {updated_count} levels updated")

tree.write('modded_files/ItemDefinitions_Weapons', encoding='utf-8', xml_declaration=True)
print(f"\nTotal updates: {total_updated}")
print(f"File saved")

with open('weapon_changes_final.json', 'w') as f:
    json.dump(changes, f, indent=2)
