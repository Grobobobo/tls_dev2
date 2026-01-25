import json

# Load weapon data from Excel
with open('weapon_data.json', 'r') as f:
    weapon_data = json.load(f)

# Create a mapping of scroll names to their base weapon types
scroll_to_weapon = {
    'AxeBoomerangScroll': '1h Axe',
    'ThrowingDaggersScroll': 'Dagger',
    'ChargeScroll': '2h sword',
    'SwordBlastScroll': '2h sword',
    'SuperSpinScroll': '2H AXE',
    'GroundSmashScroll': '2H Hammer',
    'TripleSwipeScroll': 'Spear',
    'GrapeshotScroll': 'Pistol',
    'RainOfArrowsScroll': 'Shortbow',
    'ExplosiveBoltScroll': 'Crossbow',
    'AssassinateScroll': 'Rifle',
    'MagicMissilesScroll': 'Wand',
    'HammerOfFaithScroll': 'Scepter',
    'DeathRayScroll': 'Magic orb',
    'ScorchingWaveScroll': 'power staff',
    'FireThrowerScroll': 'power staff',
    'FireballScroll': 'Tome of Secrets',
    'LightningStrikeScroll': 'Tome of Secrets',
    'BeeStingScroll': 'druid staff',
    'TeleportationScroll': None,  # No base damage values
}

# For each scroll, print the expected damage values from the weapon
print("Scroll to Weapon Mapping (with expected damage values):\n")
for scroll_name, weapon_sheet in sorted(scroll_to_weapon.items()):
    if weapon_sheet in weapon_data:
        levels = weapon_data[weapon_sheet]['levels']
        print(f"{scroll_name} -> {weapon_sheet}:")
        for level in sorted([int(k) for k in levels.keys()]):
            min_dmg = levels[str(level)]['min']
            max_dmg = levels[str(level)]['max']
            print(f"  Level {level}: {min_dmg}-{max_dmg}")
    else:
        print(f"{scroll_name} -> {weapon_sheet}: (NOT FOUND)")

# Note: Scrolls in the Usables file have 6 levels (0-5)
# We need to determine if they use the base variant (0 ending) or regular variant mapping
print("\n\nScrolls use variant mapping (like non-0 weapons): level 0-5 = weapon level 0-5")
