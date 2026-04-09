# -*- coding: utf-8 -*-
filepath = r"c:\Users\Marcin\Documents\GitHub\tls_dev2\Grobo's Tactical Overhaul\Localization\Loc.txt"

with open(filepath, 'rb') as f:
    raw = f.read()

has_bom = raw[:3] == b'\xef\xbb\xbf'
content = raw.decode('utf-8-sig')
nl = '\r\n' if '\r\n' in content else '\n'

count_ok = 0
count_fail = 0

def rep(c, old, new, label):
    global count_ok, count_fail
    old_nl = old.replace('\n', nl)
    new_nl = new.replace('\n', nl)
    if old_nl in c:
        c = c.replace(old_nl, new_nl, 1)
        print(f"  OK: {label}")
        count_ok += 1
    else:
        print(f"  FAIL: {label}")
        count_fail += 1
    return c

c = content

# === PerkDescription: add {5} (crit power per step) and Over {7}: {8} (AP at threshold) ===
print("=== PerkDescription_PanicMovePoint ===")
c = rep(c, 'filled, get:\n{1}\n\n', 'filled, get:\n{1}\n{5}\nOver {7}: {8}\n\n', 'Desc EN')
c = rep(c, 'obtenez :\n{1}\n\n', 'obtenez :\n{1}\n{5}\nAu-dessus de {7} : {8}\n\n', 'Desc FR')
c = rep(c, '获得：\n{1}\n\n', '获得：\n{1}\n{5}\n超过{7}时：{8}\n\n', 'Desc CN')
c = rep(c, '獲得：\n{1}\n\n', '獲得：\n{1}\n{5}\n超過{7}時：{8}\n\n', 'Desc TW')
c = rep(c, 'を獲得:\n{1}\n\n', 'を獲得:\n{1}\n{5}\n{7}以上で: {8}\n\n', 'Desc JP')
c = rep(c, 'получает:\n{1}\n\n', 'получает:\n{1}\n{5}\nСвыше {7}: {8}\n\n', 'Desc RU')
c = rep(c, 'надається:\n{1}\n\n', 'надається:\n{1}\n{5}\nПонад {7}: {8}\n\n', 'Desc UA')
c = rep(c, 'erhältst du:\n{1}\n\n', 'erhältst du:\n{1}\n{5}\nÜber {7}: {8}\n\n', 'Desc DE')
c = rep(c, 'obtienes:\n{1}\n\n', 'obtienes:\n{1}\n{5}\nMás de {7}: {8}\n\n', 'Desc ES')
c = rep(c, 'receba:\n{1}.\n\n', 'receba:\n{1}.\n{5}\nAcima de {7}: {8}\n\n', 'Desc PT')
c = rep(c, '획득:\n{1}\n\n', '획득:\n{1}\n{5}\n{7} 초과 시: {8}\n\n', 'Desc KR')

# === PerkEffectInformations: add {6} (crit total) and Over {7}: {9} (AP total) ===
print("\n=== PerkEffectInformations_PanicMovePoint ===")
c = rep(c, 'Current Bonus: {3}\n"', 'Current Bonus: {3} {6}\nOver {7}: {9}\n"', 'Effect EN')
c = rep(c, 'Bonus actuel<nbsp>: {3}\n"', 'Bonus actuel<nbsp>: {3} {6}\nAu-dessus de {7}<nbsp>: {9}\n"', 'Effect FR')
c = rep(c, '当前加成：{3}\n"', '当前加成：{3} {6}\n超过{7}时：{9}\n"', 'Effect CN')
c = rep(c, '目前加成：{3}\n"', '目前加成：{3} {6}\n超過{7}時：{9}\n"', 'Effect TW')
c = rep(c, '現在のボーナス: {3}\n"', '現在のボーナス: {3} {6}\n{7}以上で: {9}\n"', 'Effect JP')
c = rep(c, 'текущий бонус: {3}\n."', 'текущий бонус: {3} {6}\nСвыше {7}: {9}\n."', 'Effect RU')
c = rep(c, 'Поточний бонус: {3}\n"', 'Поточний бонус: {3} {6}\nПонад {7}: {9}\n"', 'Effect UA')
c = rep(c, 'Aktueller Bonus: {3}\n"', 'Aktueller Bonus: {3} {6}\nÜber {7}: {9}\n"', 'Effect DE')
c = rep(c, 'Bonificación actual: {3}\n"', 'Bonificación actual: {3} {6}\nMás de {7}: {9}\n"', 'Effect ES')
c = rep(c, 'Bônus Atual: {3}\n"', 'Bônus Atual: {3} {6}\nAcima de {7}: {9}\n"', 'Effect PT')
c = rep(c, '현재 보너스: {3}\n"', '현재 보너스: {3} {6}\n{7} 초과 시: {9}\n"', 'Effect KR')

# Write back preserving BOM and line endings
encoded = c.encode('utf-8')
if has_bom:
    encoded = b'\xef\xbb\xbf' + encoded
with open(filepath, 'wb') as f:
    f.write(encoded)

print(f"\nDone: {count_ok} OK, {count_fail} FAIL")
