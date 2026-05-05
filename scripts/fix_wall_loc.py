import csv
import os
import io

base_path = os.path.join(os.path.dirname(__file__), '..', 'base_files', 'Loc_TLS')
mod_path = os.path.join(os.path.dirname(__file__), '..', "Grobo's Tactical Overhaul", 'Localization', 'Loc.txt')

# Read base loc file to extract the two wall upgrade rows
with open(base_path, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    rows = list(reader)

stone_wall_row = None
reinforced_row = None
for row in rows:
    if row[0] == 'BuildingUpgradeTooltipDescription_UpgradeWoodenWallReinforcedToStoneWall0':
        stone_wall_row = row[:]
    elif row[0] == 'BuildingUpgradeTooltipDescription_UpgradeStoneWallToStoneWallReinforced0':
        reinforced_row = row[:]

if not stone_wall_row:
    print("ERROR: Stone wall row not found")
    exit(1)
if not reinforced_row:
    print("ERROR: Reinforced stone wall row not found")
    exit(1)

# Stone Wall upgrade: WoodenWallReinforced (150->180) to StoneWall (300->320)
# Replace "150" with "180" and "300" with "320" in all language columns
print("=== Stone Wall upgrade (before) ===")
for i, col in enumerate(stone_wall_row):
    if '150' in col or '300' in col:
        print(f"  Col {i}: has 150={col.count('150')}, 300={col.count('300')}")

for i in range(1, len(stone_wall_row)):
    stone_wall_row[i] = stone_wall_row[i].replace('150', '180').replace('300', '320')

print("=== Stone Wall upgrade (after) ===")
for i, col in enumerate(stone_wall_row):
    if '180' in col or '320' in col:
        print(f"  Col {i}: has 180={col.count('180')}, 320={col.count('320')}")

# Reinforced Stone Wall upgrade: StoneWall (300->320) to StoneWallReinforced (450->480)
# Replace "300" with "320" and "450" with "480" in all language columns
print("\n=== Reinforced Stone Wall upgrade (before) ===")
for i, col in enumerate(reinforced_row):
    if '300' in col or '450' in col:
        print(f"  Col {i}: has 300={col.count('300')}, 450={col.count('450')}")

for i in range(1, len(reinforced_row)):
    reinforced_row[i] = reinforced_row[i].replace('300', '320').replace('450', '480')

print("=== Reinforced Stone Wall upgrade (after) ===")
for i, col in enumerate(reinforced_row):
    if '320' in col or '480' in col:
        print(f"  Col {i}: has 320={col.count('320')}, 480={col.count('480')}")

# Encode rows to CSV format
def encode_row(row):
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator='')
    writer.writerow(row)
    return buf.getvalue()

stone_line = encode_row(stone_wall_row)
reinforced_line = encode_row(reinforced_row)

# Read mod Loc.txt and append
with open(mod_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Check if keys already exist
if 'BuildingUpgradeTooltipDescription_UpgradeWoodenWallReinforcedToStoneWall0' in content:
    print("\nWARNING: Stone wall key already exists in mod Loc.txt, skipping")
else:
    content = content.rstrip('\n') + '\n' + stone_line + '\n'
    print("\nAdded Stone Wall upgrade entry")

if 'BuildingUpgradeTooltipDescription_UpgradeStoneWallToStoneWallReinforced0' in content:
    print("WARNING: Reinforced stone wall key already exists in mod Loc.txt, skipping")
else:
    content = content.rstrip('\n') + '\n' + reinforced_line + '\n'
    print("Added Reinforced Stone Wall upgrade entry")

with open(mod_path, 'w', encoding='utf-8', newline='') as f:
    f.write(content)

print("\nDone!")
