import json
import xml.etree.ElementTree as ET
from pathlib import Path

# Load the mapping data
with open('weapon_stat_bonuses.json', 'r') as f:
    mapping_data = json.load(f)

weapon_variants_mapping = mapping_data['weapon_variants_mapping']
stat_name_mapping = mapping_data['stat_name_mapping']
tier1_bonuses = mapping_data['tier1_bonuses']
tier2_bonuses = mapping_data['tier2_bonuses']

# Convert string keys back to ints for tier bonuses
tier1_bonuses = {int(k): v for k, v in tier1_bonuses.items()}
tier2_bonuses = {int(k): v for k, v in tier2_bonuses.items()}

# OffHand weapons to exclude
OFFHAND_WEAPONS = {'BattleMageMagicWand', 'BattleMageSword', 'DuelingPistol', 'MysticHammer', 
                   'ParryingDagger', 'PreciseHandCrossbow', 'ReliableMagicScepter', 'SwiftAxe', 
                   'TransferMagicOrb', 'WarpCrystal'}

# Manual mapping for tricky weapon names
WEAPON_NAME_MAPPING = {
    'Axe': '1h Axe',
    'MagicWand': 'Wand',
    'MagicScepter': 'Scepter',
    'MagicStaff': 'power staff',
    'TomeOfMagic': 'Tome of Secrets',
    'DruidicStaff': 'druid staff',
    'WarShield': 'War Shield',
}

# Build a case-insensitive mapping from Excel weapon names
excel_weapons_lower = {k.lower(): k for k in weapon_variants_mapping.keys()}

def find_excel_weapon_name(xml_base):
    """Find the Excel weapon name for a given XML weapon base (case-insensitive)"""
    # Check manual mapping first
    if xml_base in WEAPON_NAME_MAPPING:
        return WEAPON_NAME_MAPPING[xml_base]
    
    # Normalize the XML base name for comparison
    normalized = xml_base.lower()
    
    # Try direct lookup
    if normalized in excel_weapons_lower:
        return excel_weapons_lower[normalized]
    
    # Try removing hyphens and spaces from Excel names for comparison
    for excel_name, excel_name_lower in excel_weapons_lower.items():
        if excel_name_lower.replace(' ', '').replace('-', '') == normalized.replace(' ', '').replace('-', ''):
            return excel_name
    
    # Try checking if any excel weapon name contains the normalized string
    for excel_name, excel_name_lower in excel_weapons_lower.items():
        # Remove all non-alphanumeric for fuzzy matching
        excel_clean = ''.join(c for c in excel_name_lower if c.isalnum())
        xml_clean = ''.join(c for c in normalized if c.isalnum())
        if excel_clean == xml_clean:
            return excel_name
    
    return None

def parse_composite_value(value_str):
    """Parse composite values like '7;2' or '40;8' into a list"""
    if isinstance(value_str, str) and ';' in value_str:
        try:
            return [int(float(v.strip())) for v in value_str.split(';')]
        except:
            return []
    elif isinstance(value_str, (int, float)):
        return [int(value_str)]
    return []

def find_stat_value_in_bonuses(stat_name, bonuses_dict):
    """Find a stat value in bonuses dict, handling both simple and composite keys"""
    # First, try exact match
    if stat_name in bonuses_dict:
        return bonuses_dict[stat_name]
    
    # Try case-insensitive match
    for key, value in bonuses_dict.items():
        if key.lower() == stat_name.lower():
            return value
    
    # Try to find in composite keys (e.g., "Mana;Mana Regen")
    for key, value in bonuses_dict.items():
        if ';' in key:
            parts = [p.strip() for p in key.split(';')]
            for i, part in enumerate(parts):
                if part.lower() == stat_name.lower():
                    # Return the corresponding part of the value
                    values = parse_composite_value(value)
                    if i < len(values):
                        return values[i]
    
    return None

def map_excel_stat_to_xml(excel_stat_name):
    """Map a single Excel stat name to XML stat name (case-insensitive)"""
    if excel_stat_name in stat_name_mapping:
        return stat_name_mapping[excel_stat_name]
    
    # Try case-insensitive match
    for key, val in stat_name_mapping.items():
        if key.lower() == excel_stat_name.lower():
            return val
    
    return None

def create_base_stat_bonuses(weapon_id, variant_id, level_id, excel_weapon_name):
    """Create BaseStatBonuses element for a weapon variant"""
    
    # Weapons ending in 0-1 should not have bonuses (except WarShield which always gets -20 Dodge)
    if variant_id == 0 or variant_id == 1:
        if excel_weapon_name.lower() == 'war shield':
            # WarShield always has -20 Dodge
            base_stat_bonuses = ET.Element('BaseStatBonuses')
            stat_elem = ET.SubElement(base_stat_bonuses, 'BaseStat')
            stat_elem.set('Stat', 'Dodge')
            stat_elem.set('Value', '-20')
            return base_stat_bonuses
        else:
            return None
    
    # Get the stat names for this variant
    variant_mapping = weapon_variants_mapping.get(excel_weapon_name)
    if not variant_mapping:
        return None
    
    variant_id_str = str(variant_id)
    if variant_id_str not in variant_mapping:
        return None
    
    stat_names = variant_mapping[variant_id_str]
    
    # Normalize stat_names to always be a list
    if isinstance(stat_names, str):
        stat_names = [stat_names]
    
    # Determine which tier to use
    if variant_id in [2, 3]:
        bonuses_dict = tier1_bonuses.get(level_id, {})
    elif variant_id in [4, 5]:
        bonuses_dict = tier2_bonuses.get(level_id, {})
    else:
        return None
    
    # Create BaseStatBonuses element
    base_stat_bonuses = ET.Element('BaseStatBonuses')
    
    for excel_stat_name in stat_names:
        excel_stat_name = excel_stat_name.strip()
        
        # Map to XML stat name
        xml_stat_name = map_excel_stat_to_xml(excel_stat_name)
        if not xml_stat_name:
            continue
        
        # Find the value in bonuses dict
        value = find_stat_value_in_bonuses(excel_stat_name, bonuses_dict)
        if value is None:
            continue
        
        # Handle composite values and map to correct XML stat
        if isinstance(xml_stat_name, list):
            # Multi-stat mapping (e.g., ["MovePointsTotal", "Dodge"])
            values = parse_composite_value(value)
            for i, stat in enumerate(xml_stat_name):
                if i < len(values):
                    stat_elem = ET.SubElement(base_stat_bonuses, 'BaseStat')
                    stat_elem.set('Stat', stat)
                    stat_elem.set('Value', str(values[i]))
        else:
            # Single stat mapping
            values = parse_composite_value(value)
            if values:
                # If we have composite values but single stat mapping, take first value
                stat_elem = ET.SubElement(base_stat_bonuses, 'BaseStat')
                stat_elem.set('Stat', xml_stat_name)
                stat_elem.set('Value', str(values[0]))
    
    # For WarShield, always add -20 Dodge
    if excel_weapon_name.lower() == 'war shield':
        # Check if Dodge already exists
        dodge_exists = any(elem.get('Stat') == 'Dodge' for elem in base_stat_bonuses.findall('BaseStat'))
        if not dodge_exists:
            stat_elem = ET.SubElement(base_stat_bonuses, 'BaseStat')
            stat_elem.set('Stat', 'Dodge')
            stat_elem.set('Value', '-20')
        else:
            # Update existing Dodge to be -20
            for elem in base_stat_bonuses.findall('BaseStat'):
                if elem.get('Stat') == 'Dodge':
                    elem.set('Value', '-20')
    
    return base_stat_bonuses if len(base_stat_bonuses) > 0 else None

def process_xml_file(file_path, output_path):
    """Process an ItemDefinitions XML file and update stat bonuses"""
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    changes = []
    update_count = 0
    
    for item in root.findall('.//ItemDefinition'):
        item_id = item.get('Id')
        if not item_id:
            continue
        
        # Extract variant ID from the weapon ID
        variant_id = int(item_id[-1]) if item_id[-1].isdigit() else None
        if variant_id is None:
            continue
        
        # Get the weapon base (strip off the trailing digit)
        weapon_base = item_id.rstrip('0123456789')
        
        # Skip OffHand weapons
        if weapon_base in OFFHAND_WEAPONS:
            continue
        
        # Find the Excel weapon name (case-insensitive)
        excel_weapon_name = find_excel_weapon_name(weapon_base)
        if not excel_weapon_name:
            continue
        
        # Process all level variants
        level_variations = item.find('LevelVariations')
        if level_variations is None:
            continue
        
        for level_elem in level_variations.findall('Level'):
            level_id_str = level_elem.get('Id')
            if level_id_str is None:
                continue
            
            try:
                level_id = int(level_id_str)
            except (ValueError, TypeError):
                continue
            
            # Create new BaseStatBonuses
            new_base_stat_bonuses = create_base_stat_bonuses(item_id, variant_id, level_id, excel_weapon_name)
            
            # Remove old BaseStatBonuses if it exists
            old_bsb = level_elem.find('BaseStatBonuses')
            if old_bsb is not None:
                level_elem.remove(old_bsb)
            
            # Add new BaseStatBonuses
            if new_base_stat_bonuses is not None:
                level_elem.append(new_base_stat_bonuses)
                
                # Log the change
                bonus_str = ', '.join([
                    f"{stat.get('Stat')}={stat.get('Value')}"
                    for stat in new_base_stat_bonuses.findall('BaseStat')
                ])
                print(f"  {item_id} Level {level_id}: {bonus_str}")
                changes.append({
                    'weapon_id': item_id,
                    'variant_id': variant_id,
                    'level': level_id,
                    'bonuses': bonus_str
                })
                update_count += 1
    
    # Write the updated XML
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    return update_count, changes

# Process all files
all_changes = []
total_updates = 0

files_to_process = [
    ('modded_files/ItemDefinitions_Weapons', 'ItemDefinitions_Weapons'),
    ('modded_files/ItemDefinitions_DLC1', 'ItemDefinitions_DLC1'),
    ('modded_files/ItemDefinitions_DLC2', 'ItemDefinitions_DLC2'),
]

for file_path, display_name in files_to_process:
    print(f"\n{'='*80}")
    print(f"Processing: {file_path}")
    print(f"{'='*80}")
    update_count, changes = process_xml_file(file_path, file_path)
    all_changes.extend(changes)
    total_updates += update_count
    print(f"\n[+] Updated {update_count} BaseStatBonuses in {file_path}")

# Save changes log
with open('stat_bonus_changes.json', 'w') as f:
    json.dump(all_changes, f, indent=2)

print(f"\n{'='*80}")
print(f"[+] TOTAL UPDATES: {total_updates} BaseStatBonuses across all files")
print(f"{'='*80}")
print(f"[+] Changes log saved to stat_bonus_changes.json")
