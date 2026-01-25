import json
import xml.etree.ElementTree as ET
from collections import defaultdict

# Load variant data
with open('weapon_variants.json', 'r') as f:
    data = json.load(f)

variant_stats = data['variant_stats']
tier1_headers = data['tier1_headers']
tier1_data = data['tier1_data']
tier2_headers = data['tier2_headers']
tier2_data = data['tier2_data']

# CORRECTED mapping from Excel stat names to internal XML stat names
# Based on UnitLevelUpStatDefinitions
stat_name_mapping = {
    # Main Stats
    'Momentum': 'MomentumAttacks',
    'Opportunism': 'OpportunisticAttacks',
    'Isolation': 'IsolatedAttacks',
    'Physical Damage': 'PhysicalDamage',
    'Ranged Damage': 'RangedDamage',
    'Magic Damage': 'MagicalDamage',
    'Skill Range': 'SkillRangeModifier',
    'Move Points': 'MovePointsTotal',
    'Dodge': 'Dodge',
    'Stun Chance': 'StunChanceModifier',
    'XP gain': 'ExperienceGainMultiplier',
    'Block': 'Block',
    'Health': 'HealthTotal',
    'Health Regen': 'HealthRegen',
    'Mana': 'ManaTotal',
    'Mana Regen': 'ManaRegen',
    'Reliability': 'Reliability',
    'Critical Power': 'CriticalPower',
    'Critical': 'Critical',
    'Poison Damage': 'PoisonDamageModifier',
    'Accuracy': 'Accuracy',
    'Armor': 'ArmorTotal',
    'Resistance': 'Resistance',
    'Resistance Reduction': 'ResistanceReduction',
    'Resistance reduction': 'ResistanceReduction',
    'Propagation Bounces': 'PropagationBouncesModifier',
    'Propagation Damage': 'PropagationDamage',
    
    # Composite stat keys - map exact variant combinations to XML stat lists
    'Health;Health Regen': ['HealthTotal', 'HealthRegen'],
    'Move Points;Dodge': ['MovePointsTotal', 'Dodge'],
    'Armor;Resistance': ['ArmorTotal', 'Resistance'],
    'Mana;Mana Regen': ['ManaTotal', 'ManaRegen'],
    
    # Case-insensitive variants
    'momentum': 'MomentumAttacks',
    'opportunism': 'OpportunisticAttacks',
    'isolation': 'IsolatedAttacks',
    'physical damage': 'PhysicalDamage',
    'ranged damage': 'RangedDamage',
    'magic damage': 'MagicalDamage',
    'skill range': 'SkillRangeModifier',
    'move points': 'MovePointsTotal',
    'dodge': 'Dodge',
    'stun chance': 'StunChanceModifier',
    'xp gain': 'ExperienceGainMultiplier',
    'block': 'Block',
    'health': 'HealthTotal',
    'health regen': 'HealthRegen',
    'mana': 'ManaTotal',
    'mana regen': 'ManaRegen',
    'reliability': 'Reliability',
    'critical power': 'CriticalPower',
    'critical': 'Critical',
    'poison damage': 'PoisonDamageModifier',
    'accuracy': 'Accuracy',
    'armor': 'ArmorTotal',
    'resistance': 'Resistance',
    'resistance reduction': 'ResistanceReduction',
    'propagation bounces': 'PropagationBouncesModifier',
    'propagation damage': 'PropagationDamage',
}

# Build Tier 1 bonuses dictionary
# Note: Excel has rows 1-6, which map to weapon levels 0-5
tier1_bonuses = {}
for level in range(6):
    level_str = str(level + 1)  # Excel row is level + 1
    tier1_bonuses[level] = {}
    if level_str in tier1_data:
        for header_idx, header in enumerate(tier1_headers):
            if header_idx < len(tier1_data[level_str]):
                try:
                    value = tier1_data[level_str][header_idx]
                    if isinstance(value, str):
                        try:
                            value = float(value)
                        except:
                            pass  # Keep as string if it's composite like "7;2"
                    else:
                        value = float(value)
                    tier1_bonuses[level][header] = value
                except (ValueError, TypeError):
                    pass

# Build Tier 2 bonuses dictionary
# Note: Excel has rows 1-6, which map to weapon levels 0-5
tier2_bonuses = {}
for level in range(6):
    level_str = str(level + 1)  # Excel row is level + 1
    tier2_bonuses[level] = {}
    if level_str in tier2_data:
        for header_idx, header in enumerate(tier2_headers):
            if header_idx < len(tier2_data[level_str]):
                try:
                    value = tier2_data[level_str][header_idx]
                    if isinstance(value, str):
                        try:
                            value = float(value)
                        except:
                            pass  # Keep as string if it's composite like "12;4"
                    else:
                        value = float(value)
                    tier2_bonuses[level][header] = value
                except (ValueError, TypeError):
                    pass

# Build weapon variants mapping
weapon_variants_mapping = {}
for weapon, variants in variant_stats.items():
    weapon_variants_mapping[weapon] = variants

# Output
output = {
    'weapon_variants_mapping': weapon_variants_mapping,
    'stat_name_mapping': stat_name_mapping,
    'tier1_bonuses': tier1_bonuses,
    'tier2_bonuses': tier2_bonuses,
}

with open('weapon_stat_bonuses.json', 'w') as f:
    json.dump(output, f, indent=2)

print("[+] Built weapon_stat_bonuses.json")
print(f"[+] Mapped {len(weapon_variants_mapping)} weapons")
print(f"[+] Created {len(stat_name_mapping)} stat name mappings")
print(f"[+] Tier 1: {len(tier1_headers)} stats, 6 levels")
print(f"[+] Tier 2: {len(tier2_headers)} stats, 6 levels")

# Debug: Show resistance mappings
print("\n[DEBUG] Resistance-related mappings:")
for k, v in stat_name_mapping.items():
    if 'resist' in k.lower():
        print(f"  {k} -> {v}")

print("\n[DEBUG] Tier 1 Resistance values:")
for level in range(6):
    res = tier1_bonuses[level].get('Resistance', None)
    print(f"  Level {level}: {res}")

print("\n[DEBUG] Tier 2 Resistance Reduction values:")
for level in range(6):
    res_red = tier2_bonuses[level].get('Resistance Reduction', None)
    print(f"  Level {level}: {res_red}")
