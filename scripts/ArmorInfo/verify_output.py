import xml.etree.ElementTree as ET, sys
sys.stdout.reconfigure(encoding='utf-8')

# Check ClothArmor0 (Worn Cloth Armor) - New values: Dodge 5/7/9/11/13/15, Price 10/15/24/38/58/85, MovePointsTotal 1/1/1/1/2/2
tree = ET.parse(r'c:\Users\Marcin\Documents\GitHub\tls_dev2\modded_files\ItemDefinitions_BodyArmors.xml')
root = tree.getroot()

print("=== ClothArmor0 (Worn Cloth Armor) ===")
item = root.findall('ItemDefinition')[0]
print(f'Item: {item.get("Id")}')
for lv in item.findall('.//Level'):
    lv_id = lv.get('Id')
    msb = lv.find('MainStatBonus')
    bp = lv.find('BasePrice')
    bsbs = [(b.get('Stat'), b.text) for b in lv.findall('.//BaseStatBonus')]
    sk = [s.text for s in lv.findall('.//Skill')]
    print(f'  Lv{lv_id}: {msb.get("Stat")}={msb.text}, Price={bp.text}, Bonuses={bsbs}, Skills={sk}')

print()
print("=== FallenLordArmor (Stag Armor) ===")
# Item index 24 = FallenLordArmor, should match New row: Stag Armor
# New: Dodge 4/6/8/10/12/14, MovePointsTotal 1/1/1/1/2/2, OpportunisticAttacks 5/7/9/11/13/15, Accuracy 4/5/7/8/10/11, Resistance 4/5/7/8/10/11
item = root.findall('ItemDefinition')[24]
print(f'Item: {item.get("Id")}')
for lv in item.findall('.//Level'):
    lv_id = lv.get('Id')
    msb = lv.find('MainStatBonus')
    bp = lv.find('BasePrice')
    bsbs = [(b.get('Stat'), b.text) for b in lv.findall('.//BaseStatBonus')]
    sk = [s.text for s in lv.findall('.//Skill')]
    print(f'  Lv{lv_id}: {msb.get("Stat")}={msb.text}, Price={bp.text}, Bonuses={bsbs}, Skills={sk}')

print()
# Check Helmet0 (Leather Cap) - New: Dodge 6/8/11/13/16/18, Resistance 2/2/3/3/4/4
tree2 = ET.parse(r'c:\Users\Marcin\Documents\GitHub\tls_dev2\modded_files\ItemDefinitions_Helmets.xml')
root2 = tree2.getroot()
print("=== Helmet0 (Leather Cap) ===")
item = root2.findall('ItemDefinition')[0]
print(f'Item: {item.get("Id")}')
for lv in item.findall('.//Level'):
    lv_id = lv.get('Id')
    msb = lv.find('MainStatBonus')
    bp = lv.find('BasePrice')
    bsbs = [(b.get('Stat'), b.text) for b in lv.findall('.//BaseStatBonus')]
    print(f'  Lv{lv_id}: {msb.get("Stat")}={msb.text}, Price={bp.text}, Bonuses={bsbs}')

print()
# Check ClothPants0 (Worn Pants) - New: Dodge 4/5/7/8/10/11, ManaTotal 4/6/8/10/12/14
tree3 = ET.parse(r'c:\Users\Marcin\Documents\GitHub\tls_dev2\modded_files\ItemDefinitions_Pants.xml')
root3 = tree3.getroot()
print("=== ClothPants0 (Worn Pants) ===")
item = root3.findall('ItemDefinition')[0]
print(f'Item: {item.get("Id")}')
for lv in item.findall('.//Level'):
    lv_id = lv.get('Id')
    msb = lv.find('MainStatBonus')
    bp = lv.find('BasePrice')
    bsbs = [(b.get('Stat'), b.text) for b in lv.findall('.//BaseStatBonus')]
    print(f'  Lv{lv_id}: {msb.get("Stat")}={msb.text}, Price={bp.text}, Bonuses={bsbs}')
