#!/usr/bin/env python3
"""生成《坊市定向寻物表.xlsx》
复刻游戏定向寻物筛选逻辑（从游戏本体逆向），对每个坊市可寻物品计算六维条件信号：
分类 / 元素 / Buff / 机制 / 属性 / 特征，并生成可直接在 Excel 中组合筛选的工作簿。

用法：
  python gen_search_xlsx.py                 # 默认 Demo 版本
  python gen_search_xlsx.py --edition playtest
  python gen_search_xlsx.py --edition demo --out "D:/somewhere/表.xlsx"
"""
import argparse
import json, os, sys, re, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 复用 server.py 的版本注册表与自动发现逻辑，避免写死路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import server as srv

PROJECT_ROOT = srv.ROOT


def resolve_sdk_path(edition_id):
    ctx = srv.get_edition(edition_id)
    root = ctx["gameRoot"]
    if not root:
        raise SystemExit(f"未找到版本 {edition_id} 的游戏目录；请确认游戏已安装或创建 game-assets.txt。")
    cand = [
        os.path.join(root, "ModSDK", "reference", "domains", "items", "base-item-public-catalog.json"),
        os.path.join(root, "ModSDK", "reference", "domains", "items", "base-item-index.json"),
    ]
    for c in cand:
        if os.path.isfile(c):
            return c
    raise SystemExit(f"未找到版本 {edition_id} 的 ModSDK 物品目录：{root}")


# ---------- 游戏本体词典（从游戏代码逆向） ----------
CAT_CN = {  # eu
    "scripture": "功法", "recipe": "配方", "pill": "丹药", "implement": "法器",
    "divination": "问签", "chest": "宝箱", "material": "灵材", "formation": "阵法",
    "talisman": "符箓", "throwable": "投掷物", "spell": "招式", "bag": "储物袋",
}
ELEM_CN = {  # Xd
    "metal": "金", "wood": "木", "water": "水", "fire": "火", "earth": "土",
    "thunder": "雷", "poison": "毒", "ice": "冰", "wind": "风", "none": "无",
}
BUFF_CN = {  # te
    "heat": "炎热", "cold": "寒冷", "heavy": "沉重", "agile": "敏捷", "dew": "露水",
    "condensedMetal": "凝金", "lightningCharge": "引雷", "armor": "护甲", "sturdy": "坚硬",
    "poison": "中毒", "dodge": "闪避", "regeneration": "再生", "burning": "灼烧",
    "frost": "冰霜", "paralysis": "麻痹", "windPressure": "风压",
}
STAT_CN = {  # qi
    "health": "气血", "stamina": "体力", "mana": "法力", "bloodEssence": "精血", "spiritSense": "神识",
}
MECH_CN = {  # w_ + Cc signals
    "start": "开局", "rotation": "轮转", "transmutation": "转化", "adjacent": "相邻",
    "exhaust": "耗尽", "cultivation": "修为", "lucky": "机缘", "spiritStones": "灵石",
    "charge": "充能", "throw": "投掷", "cooldown": "冷却", "formationSwitch": "变阵",
    "consumption": "消解", "scriptureBreakthrough": "功法突破",
}
FEAT_CN = {  # au
    "artifact": "法宝", "alchemyFurnace": "炼丹炉", "weapon": "武器", "equipment": "装备",
    "accessory": "饰品", "spiritPlant": "灵植", "beastMaterial": "妖兽材料", "ore": "矿石",
    "forgingSand": "灵砂", "forged": "炼制", "task": "任务", "brush": "笔", "dagger": "匕首",
    "sword": "剑", "orb": "宝珠", "ruler": "尺", "armor": "防具", "spear": "枪", "box": "匣",
    "staff": "法杖",
}
GRADE_CN = {"common": "凡品", "spirit": "灵品", "mystic": "玄品", "earth": "地品", "heaven": "天品", "immortal": "仙品"}

SHOPS = [  # d1 config from game data
    {"id": "commissioned_item_search_qi", "name": "寻珍行", "merchant": "万里寻", "grade": "common",
     "grade_cn": "凡品", "loc": "map_01 河市镇", "stone": "押金5灵石·3选1", "lucky": "72机缘·5选2"},
    {"id": "commissioned_item_search_foundation", "name": "探玄楼", "merchant": "闻千机", "grade": "spirit",
     "grade_cn": "灵品", "loc": "map_102 筑基坊市", "stone": "押金50灵石·3选1", "lucky": "160机缘·5选2"},
]


def effects_iter(effects):
    def walk(o):
        if isinstance(o, dict):
            yield o
            for v in o.values():
                yield from walk(v)
        elif isinstance(o, list):
            for x in o:
                yield from walk(x)
    return walk(effects)


def compute_signals(it):
    """复刻游戏 yc/KA/Cc 的六维信号提取。"""
    sig = {
        "category": it.get("category"),
        "elements": set(it.get("baseElements") or [it.get("element")]),
        "buffs": set(), "stats": set(), "mech": set(),
        "features": set(it.get("features") or []),
    }
    el = it.get("effectList") or []

    def walk(o, fn):
        if isinstance(o, dict):
            fn(o)
            for v in o.values():
                walk(v, fn)
        elif isinstance(o, list):
            for x in o:
                walk(x, fn)

    # 逐效果显式处理（上面闭包写法易错，这里直接遍历）
    sig["mech"] = set()
    has_charge = False

    def scan(o):
        nonlocal has_charge
        if isinstance(o, dict):
            for k, v in o.items():
                if k in ("buffId", "sourceBuffId", "targetBuffId") and isinstance(v, str):
                    sig["buffs"].add(v)
                if k == "stat" and isinstance(v, str) and v in STAT_CN:
                    sig["stats"].add(v)
                if k == "damageHealth" and v is not None:
                    sig["stats"].add("health")
                if k == "startOnly" and v is True:
                    sig["mech"].add("start")
                if k in ("neighborCounts", "directionalNeighbors", "alignments", "lineChains"):
                    sig["mech"].add("adjacent")
                if k == "source" and v == "adjacentItems":
                    sig["mech"].add("adjacent")
                if k == "type" and v in ("adjacent", "directional-adjacent"):
                    sig["mech"].add("adjacent")
                if k in ("cultivationGain", "cultivationCost", "cultivationGainPerStack", "cultivationGainPerBuffStack") and v is not None:
                    sig["mech"].add("cultivation")
                if k in ("spiritStoneGain", "spiritStoneCost", "stoneCost") and v is not None:
                    sig["mech"].add("spiritStones")
                if k in ("luckyGain", "luckyCost", "luckyGainFromPrimaryStat") and v is not None:
                    sig["mech"].add("lucky")
                if k == "kind" and v == "scriptureProgression":
                    sig["mech"].add("scriptureBreakthrough")
                if k == "trigger" and v == "onConsumableExhausted":
                    sig["mech"].add("exhaust")
                if k == "charge" and v is not None:
                    has_charge = True
                scan(v)
        elif isinstance(o, list):
            for x in o:
                scan(x)

    scan(el)
    for e in effects_iter(el):
        if e.get("kind") == "consumable" or e.get("consumable") is not None:
            sig["mech"].add("exhaust")
        if e.get("kind") == "periodicPulse":
            pp = e.get("periodicPulse") or {}
            if pp.get("intervalSec") is not None and pp.get("startOnly") is not True:
                sig["mech"].add("rotation")
            costs = any(pp.get(k) for k in ("cultivationCost", "spiritStoneCost", "luckyCost", "lifespanCost")) or pp.get("primaryStatCosts")
            gains = any(pp.get(k) for k in ("cultivationGain", "spiritStoneGain", "luckyGain", "lifespanGain")) or pp.get("primaryStatGains")
            if costs and gains:
                sig["mech"].add("transmutation")
    if sig["features"] and "formationSwitch" in str(el):
        pass
    if it.get("category") == "throwable":
        sig["mech"].add("throw")
    if it.get("category") in ("pill", "talisman"):
        sig["mech"].add("exhaust")
    if it.get("category") == "scripture" or any(e.get("kind") == "scriptureProgression" for e in effects_iter(el)):
        pass  # scriptureProgression 已加
    if has_charge:
        sig["mech"].add("charge")
    # spell activeCast -> cooldown
    def find_key(o, key):
        if isinstance(o, dict):
            if key in o and o[key]:
                return True
            return any(find_key(v, key) for v in o.values())
        elif isinstance(o, list):
            return any(find_key(x, key) for x in o)
        return False
    if it.get("category") == "spell" and find_key(el, "activeCast"):
        sig["mech"].add("cooldown")
    # formationSwitch trigger
    if "formationSwitch" in json.dumps(el):
        sig["mech"].add("formationSwitch")
    # onBuffConsumed trigger -> consumption
    if "onBuffConsumed" in json.dumps(el):
        sig["mech"].add("consumption")
    return sig


def bazaar_eligible(it):
    if not it.get("isPublished"):
        return False
    ch = (it.get("distribution") or {}).get("channels") or []
    if "bazaar" not in ch:
        return False
    feats = it.get("features") or []
    if "task" in feats or "beastMaterial" in feats or "forged" in feats:
        return False
    if (it.get("storageBag") or {}).get("sourceEnemyIds"):
        return False
    return True


def load_items(sdk_path):
    raw = json.load(open(sdk_path, encoding="utf-8"))
    items = raw["items"] if isinstance(raw, dict) and "items" in raw else raw
    out = []
    for it in items:
        if not bazaar_eligible(it):
            continue
        s = compute_signals(it)
        shops = [sh for sh in SHOPS if sh["grade"] == it.get("grade")]
        out.append({
            "id": it["numericId"], "name": it.get("name"), "grade": it.get("grade"),
            "grade_cn": GRADE_CN.get(it.get("grade"), it.get("grade")),
            "cat": it.get("category"), "cat_cn": CAT_CN.get(it.get("category"), it.get("category")),
            "elements": s["elements"], "buffs": s["buffs"], "stats": s["stats"],
            "mech": s["mech"], "features": s["features"],
            "price": it.get("price"), "desc": (it.get("description") or "")[:100],
            "shops": shops,
        })
    out.sort(key=lambda x: x["id"])
    return out


# ---------- Excel 样式 ----------
HDR_FILL = PatternFill("solid", fgColor="6E5220")
HDR_FONT = Font(color="FFF4DC", bold=True, size=10)
TITLE_FONT = Font(color="F0C868", bold=True, size=13)
NOTE_FONT = Font(color="A6987F", size=9)
GRADE_FILL = {
    "凡品": PatternFill("solid", fgColor="4a4a46"),
    "灵品": PatternFill("solid", fgColor="2e5e46"),
    "玄品": PatternFill("solid", fgColor="3a5a80"),
    "地品": PatternFill("solid", fgColor="5a4470"),
    "天品": PatternFill("solid", fgColor="6e5220"),
    "仙品": PatternFill("solid", fgColor="7a3a34"),
}
THIN = Border(*[Side(style="thin", color="3a3226")] * 4)
CENTER = Alignment(horizontal="center", vertical="center")




HDR_FILL = PatternFill("solid", fgColor="6E5220")

def run(edition_id="demo", out_path=None):
    sdk_path = resolve_sdk_path(edition_id)
    edition = srv.EDITION_BY_ID.get(edition_id, {})
    if not out_path:
        out_path = os.path.join(PROJECT_ROOT, "output", "坊市定向寻物表.xlsx")
    items = load_items(sdk_path)
    print(f"版本: {edition.get('label', edition_id)}  物品目录: {sdk_path}")
    print(f"可寻物品: {len(items)}")
    wb = Workbook()

    # ============ Sheet1 总表 ============
    ws = wb.active
    ws.title = "总表"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:N1")
    ws["A1"] = f"坊市定向寻物 · {edition.get('label', edition_id)} 版本 · 全部可寻物品总表（逻辑复刻自游戏本体）"
    ws["A1"].font = TITLE_FONT
    headers = ["编号", "名称", "品级", "分类", "元素", "Buff", "机制", "属性", "特征", "价格(灵石)", "寻珍行", "探玄楼", "描述"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=2, column=i, value=h)
    style_header(ws, 2, len(headers))
    r = 3
    for it in items:
        shop_ids = {s["id"] for s in SHOPS if s["grade"] == it["grade"]}
        row = [
            it["id"], it["name"], it["grade_cn"], it["cat_cn"],
            "、".join(ELEM_CN.get(e, e) for e in sorted(it["elements"])),
            "、".join(BUFF_CN.get(b, b) for b in sorted(it["buffs"])) or "—",
            "、".join(MECH_CN.get(m, m) for m in sorted(it["mech"])) or "—",
            "、".join(STAT_CN.get(s, s) for s in sorted(it["stats"])) or "—",
            "、".join(FEAT_CN.get(f, f) for f in sorted(it["features"])) or "—",
            it["price"] if it.get("price") is not None else "—",
            "✓" if "commissioned_item_search_qi" in shop_ids else "",
            "✓" if "commissioned_item_search_foundation" in shop_ids else "",
            it["desc"],
        ]
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", horizontal="center" if c in (1, 3, 11, 12) else "left")
        gcell = ws.cell(row=r, column=3)
        gcell.fill = GRADE_FILL.get(it["grade_cn"], GRADE_FILL["凡品"])
        gcell.font = Font(color="FFFFFF", size=10, bold=True)
        r += 1
    widths = [9, 18, 7, 9, 12, 22, 22, 14, 20, 11, 8, 8, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:M{r-1}"

    # ============ Sheet2 条件矩阵（筛选式查询） ============
    # 列 = 每个维度的每个条件值；值 = 1；用户用筛选器组合
    ws2 = wb.create_sheet("条件矩阵")
    dims = [
        ("分类", CAT_CN, lambda it: {it["cat_cn"]}),
        ("元素", ELEM_CN, lambda it: {ELEM_CN.get(e, e) for e in it["elements"]}),
        ("Buff", BUFF_CN, lambda it: {BUFF_CN.get(b, b) for b in it["buffs"]}),
        ("机制", MECH_CN, lambda it: {MECH_CN.get(m, m) for m in it["mech"]}),
        ("属性", STAT_CN, lambda it: {STAT_CN.get(s, s) for s in it["stats"]}),
        ("特征", FEAT_CN, lambda it: {FEAT_CN.get(f, f) for f in it["features"]}),
        ("品级", GRADE_CN, lambda it: {it["grade_cn"]}),
    ]
    col_headers = ["编号", "名称", "寻珍行", "探玄楼"]
    dim_cols = []
    for dname, mapping, getter in dims:
        vals = sorted({v for it in items for v in getter(it)} ) if False else None
    # build columns
    col_defs = []
    for dname, mapping, getter in dims:
        vals = sorted({v for it in items for v in getter(it) if v})
        for v in vals:
            col_defs.append((dname, v, mapping.get(v, v)))
    header = col_headers + [cd[2] for cd in col_defs]
    for c, h in enumerate(header, 1):
        ws2.cell(row=1, column=c, value=h)
    style_header(ws2, 1, len(header))
    r = 2
    for it in items:
        shop_ids = {s["id"] for s in SHOPS if s["grade"] == it["grade"]}
        ws2.cell(row=r, column=1, value=it["id"]).border = THIN
        ws2.cell(row=r, column=2, value=it["name"]).border = THIN
        ws2.cell(row=r, column=3, value="✓" if "commissioned_item_search_qi" in shop_ids else "").border = THIN
        ws2.cell(row=r, column=4, value="✓" if "commissioned_item_search_foundation" in shop_ids else "").border = THIN
        signals_flat = set()
        getter_map = {
            "分类": lambda it: {it["cat_cn"]},
            "元素": lambda it: {ELEM_CN.get(e, e) for e in it["elements"]},
            "Buff": lambda it: {BUFF_CN.get(b, b) for b in it["buffs"]},
            "机制": lambda it: {MECH_CN.get(m, m) for m in it["mech"]},
            "属性": lambda it: {STAT_CN.get(s, s) for s in it["stats"]},
            "特征": lambda it: {FEAT_CN.get(f, f) for f in it["features"]},
            "品级": lambda it: {it["grade_cn"]},
        }
        cellvals = set()
        for dname, _, _ in dims:
            cellvals |= getter_map[dname](it)
        for ci, (dname, raw_v, cn_v) in enumerate(col_defs, 5):
            cell = ws2.cell(row=r, column=ci, value=1 if cn_v in cellvals else None)
            if cn_v in cellvals:
                cell.fill = PatternFill("solid", fgColor="4a3f2e")
            cell.border = THIN
            cell.alignment = CENTER
        r += 1
    # 隐藏辅助列：命中(依定向寻物所选条件) + 累计命中序号
    hit_col = len(header) + 1
    cum_col = hit_col + 1
    hcl, ccl = get_column_letter(hit_col), get_column_letter(cum_col)
    dname_list_m = [d for d, _, _ in dims if d != "品级"]
    parts_all = []
    for j, dname in enumerate(dname_list_m):
        cell_ref = "'定向寻物'!$B$" + str(4 + j)
        dcols = [get_column_letter(5 + k) for k, (dn, _, _) in enumerate(col_defs) if dn == dname]
        col_expr = ",".join(cl + "{r}" for cl in dcols)
        parts_all.append((cell_ref, col_expr))
    for i in range(len(items)):
        r = i + 2
        parts = []
        for cell_ref, col_expr in parts_all:
            ce = col_expr.replace("{r}", str(r))
            parts.append('IF({cr}="(不限)",1,MAX(0,{ce}))'.format(cr=cell_ref, ce=ce))
        # 商铺品阶过滤：寻珍行只搜凡品(C列✓)，探玄楼只搜灵品(D列✓)
        parts.append("IF('定向寻物'!$B$11=\"寻珍行\",IF(C{r}=\"✓\",1,0),IF(D{r}=\"✓\",1,0))".format(r=r))
        formula = "=IF(AND(" + ",".join(parts) + '),1,0)'
        c = ws2.cell(row=r, column=hit_col, value=formula)
        c.font = Font(size=8, color="6f6350")
        prev = ccl + str(r - 1) if i > 0 else "0"
        c2 = ws2.cell(row=r, column=cum_col, value="=" + prev + "+" + hcl + str(r))
        c2.font = Font(size=8, color="6f6350")
    ws2.column_dimensions[hcl].hidden = True
    ws2.column_dimensions[ccl].hidden = True

    # freeze + widths
    ws2.freeze_panes = "E2"
    ws2.column_dimensions["A"].width = 9
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 7
    ws2.column_dimensions["D"].width = 7
    for c in range(5, len(header) + 1):
        ws2.column_dimensions[get_column_letter(c)].width = 6
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(header))}{r-1}"

    # ============ Sheet「定向寻物」：仿游戏交互，选条件即出结果 ============
    wsr = wb.create_sheet("定向寻物")
    wsr.sheet_view.showGridLines = False
    # 深色主题：整片区域铺深色底
    DARK = PatternFill("solid", fgColor="141210")
    for r in range(1, 150):
        for c in range(1, 11):
            wsr.cell(row=r, column=c).fill = DARK
    # 列宽
    wsr.column_dimensions["A"].width = 16
    for col in "BCDEFG":
        wsr.column_dimensions[col].width = 18
    wsr.column_dimensions["A"].width = 14

    wsr["A1"] = "定向寻物 · 交互查询（与游戏坊市同规则）"
    wsr["A1"].font = TITLE_FONT
    wsr["A2"] = "在下方每个条件格用下拉框选择（可留空=不限），选择后自动计算定金，下方列出该条件下两间商铺能搜到的全部物品。"
    wsr["A2"].font = NOTE_FONT

    # 条件选择区（每个维度一个下拉，值域用隐藏的选项列）
    criteria_rows = []
    labels = ["分类", "元素", "Buff", "机制", "属性", "特征"]
    # 为每个维度收集中文值域，写入隐藏列 J 以后
    # 下拉选项 = 游戏本体完整选项表（与游戏内寻物下拉一致），而非仅当前池中出现的值
    dim_values = {
        "分类": sorted(set(CAT_CN.values())),
        "元素": sorted(set(ELEM_CN.values())),
        "Buff": sorted(set(BUFF_CN.values())),
        "机制": sorted(set(MECH_CN.values())),
        "属性": sorted(set(STAT_CN.values())),
        "特征": sorted(set(FEAT_CN.values())),
    }
    # 隐藏选项区：J1 起每维度一列
    opt_col = 10  # J
    for dname, mapping, getter in dims:
        if dname == "品级":
            continue
        vals = dim_values[dname]
        for i, v in enumerate(vals):
            c = wsr.cell(row=1 + i, column=opt_col, value=v)
            c.font = Font(size=8, color="6f6350")
        wsr.column_dimensions[get_column_letter(opt_col)].hidden = True
        opt_col += 1
    # 数据有效性下拉（含"(不限)"选项）
    from openpyxl.worksheet.datavalidation import DataValidation
    r0 = 4
    dname_list = [d for d, _, _ in dims if d != "品级"]
    for i, dname in enumerate(dname_list):
        r = r0 + i
        wsr.cell(row=r, column=1, value=dname).font = Font(color="F0C868", bold=True, size=10)
        wsr.cell(row=r, column=1).fill = PatternFill("solid", fgColor="262019")
        cell = wsr.cell(row=r, column=2, value="(不限)")
        cell.fill = PatternFill("solid", fgColor="3a2f1c")
        cell.font = Font(color="FFF4DC", size=10)
        cell.border = THIN
        cell.alignment = CENTER
        criteria_rows.append((dname, f"B{r}"))
    # 每个维度加"(不限)"选项
    for i, dname in enumerate(dname_list):
        col_letter = get_column_letter(10 + i)
        last = 1 + len(dim_values[dname])
        wsr.cell(row=last + 1, column=10 + i, value="(不限)")
    # 重新挂验证（含不限）
    for i, dname in enumerate(dname_list):
        col_letter = get_column_letter(10 + i)
        last = 1 + len(dim_values[dname]) + 1
        dv = DataValidation(type="list", formula1=f"${col_letter}$1:${col_letter}${last}", allow_blank=True)
        wsr.add_data_validation(dv)
        dv.add(wsr.cell(row=r0 + i, column=2))

    # 商铺选择 + 定金区
    rr = r0 + len(dname_list) + 1
    wsr.cell(row=rr, column=1, value="商铺").font = Font(color="F0C868", bold=True, size=10)
    shop_cell = wsr.cell(row=rr, column=2, value="寻珍行")
    shop_cell.fill = PatternFill("solid", fgColor="3a2f1c")
    shop_cell.font = Font(color="FFF4DC", size=10)
    shop_cell.border = THIN
    shop_cell.alignment = CENTER
    dv_shop = DataValidation(type="list", formula1='"寻珍行,探玄楼"', allow_blank=False)
    wsr.add_data_validation(dv_shop)
    dv_shop.add(shop_cell)
    wsr.cell(row=rr, column=4, value="条件数：").font = NOTE_FONT
    wsr.cell(row=rr, column=5, value='=SUMPRODUCT((B4:B9<>"(不限)")*1)').font = Font(color="F0C868", bold=True)
    wsr.cell(row=rr + 1, column=4, value="定金：").font = NOTE_FONT
    wsr.cell(row=rr + 1, column=5,
             value='=IF(B11="寻珍行",ROUND(5*1.3^MAX(0,E11-1),0),ROUND(50*1.3^MAX(0,E11-1),0))').font = Font(color="F0C868", bold=True)
    wsr.cell(row=rr + 1, column=6, value="灵石").font = NOTE_FONT

    # 结果区标题
    hr = rr + 3
    wsr.cell(row=hr, column=1, value="搜索结果（该条件下两间商铺都能搜到的物品）").font = Font(color="F0C868", bold=True, size=11)
    res_headers = ["编号", "名称", "品级", "分类", "元素", "Buff", "机制", "属性", "特征", "所在商铺"]
    for c, h in enumerate(res_headers, 1):
        wsr.cell(row=hr + 1, column=c, value=h)
    style_header(wsr, hr + 1, len(res_headers))

    # 紧凑结果：结果行 k 通过 MATCH(k, 累计命中列) 找到第 k 个命中物品在条件矩阵中的行
    last_row = len(items) + 1
    total_hits = f"COUNTIF('条件矩阵'!${hcl}$2:${hcl}${last_row},1)"
    wsr.cell(row=hr, column=6, value="命中件数：").font = NOTE_FONT
    wsr.cell(row=hr, column=7, value=f"={total_hits}").font = Font(color="F0C868", bold=True, size=12)
    MAX_ROWS = 120  # 结果区预留行数（更多命中请用『条件矩阵』筛选）
    for k in range(1, MAX_ROWS + 1):
        r = hr + 1 + k
        mrow_f = f"MATCH({k},'条件矩阵'!${ccl}$2:${ccl}${last_row},0)"
        for col, letter in [(1, "A"), (2, "B"), (3, "C"), (4, "D")]:
            formula = ('=IF(ROW()-{hr1}<={n},IFERROR(INDEX(\'条件矩阵\'!${letter}$2:${letter}${lr},{m}),""))'
                       .format(hr1=hr + 1, n=MAX_ROWS, m=mrow_f, letter=letter, lr=last_row))
            cell = wsr.cell(row=r, column=col, value=formula)
            cell.fill = PatternFill("solid", fgColor="1E1A16")
            cell.border = THIN
        for c in range(1, 11):
            wsr.cell(row=r, column=c).font = Font(size=10, color="FFFFFF")
    wsr.column_dimensions["K"].hidden = True

    # ============ Sheet3 使用说明 ============
    ws3 = wb.create_sheet("使用说明")
    ws3.column_dimensions["A"].width = 110
    notes = [
        "《坊市定向寻物表》使用说明",
        "",
        "本表复刻了游戏『坊市 → 定向寻物』商铺的搜索逻辑（从游戏本体逆向验证）。",
        "",
        "■ 两间商铺",
        "  · 寻珍行（商人：万里寻）—— 位于地图 map_01 河市镇，只能寻『凡品』物品",
        "    灵石档：押金 5 灵石，归来后 3 件候选选 1 件；机缘档：押下 72 点机缘，5 件候选选 2 件",
        "  · 探玄楼（商人：闻千机）—— 位于地图 map_102（元婴篇坊市），只能寻『灵品』物品",
        "    灵石档：押金 50 灵石，3 件候选选 1 件；机缘档：押下 160 点机缘，5 件候选选 2 件",
        "  · 条件每多指定一项，定金按 1.3 倍递增（定金 = 基础押金 × 1.3^(条件数-1)，向上取整）",
        "",
        "■ 方法一（推荐）：『定向寻物』工作表 —— 与游戏坊市操作一致",
        "  1. 在条件区用下拉框逐个选择条件（分类/元素/Buff/机制/属性/特征），不选=不限。",
        "  2. 选择商铺（寻珍行=凡品池 / 探玄楼=灵品池），上方自动显示条件数与所需定金。",
        "  3. 下方『搜索结果』区域实时列出该条件下能搜到的全部物品（含所在商铺）。",
        "  4. 改条件即改结果，无需手动筛选。",
        "",
        "■ 方法二：『条件矩阵』工作表 —— 传统筛选",
        "  1. 在列头筛选箭头中勾选 1 组合多维度（如 Buff=护甲 + 品级=凡品）。",
        "  2. 适合一次性浏览某个信号的完整物品集。『总表』同理可按文本列筛选。",
        "",
        "■ 游戏规则备注",
        "  · 委托期间不能立第二份契约；可终止寻访（终止会退定金）。",
        "  · 委托池排除：未发布/不可获得物品、任务物品、妖兽材料、炼制成品。",
        "  · 灵石档与机缘档的差异：机缘档候选更多（5件）、可选件数更多（2件），但消耗机缘资源。",
        "  · 候选排序按编号稳定随机（同一条件+同一时间戳结果可复现）。",
        "",
        f"游戏版本：{edition.get('label', edition_id)}（appId {edition.get('appId', '?')}）",
        f"生成时间：{__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M')}",
    ]
    for i, line in enumerate(notes, 1):
        cell = ws3.cell(row=i, column=1, value=line)
        if i == 1:
            cell.font = TITLE_FONT
        elif line and not line.startswith((" ", "《", "生成")):
            cell.font = Font(color="F0C868", bold=True, size=10)
        else:
            cell.font = NOTE_FONT

    wb.save(out_path)
    print("saved:", out_path)
    return out_path


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor="6E5220")
        cell.font = Font(color="FFF4DC", bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成《坊市定向寻物表.xlsx》")
    ap.add_argument("--edition", default=srv.DEFAULT_EDITION_ID,
                    choices=[e["id"] for e in srv.EDITIONS], help="游戏版本（默认 Demo）")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径（默认 output/坊市定向寻物表.xlsx）")
    args = ap.parse_args()
    run(args.edition, args.out)