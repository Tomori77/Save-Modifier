#!/usr/bin/env python3
"""口袋修仙 存档工坊 - 单文件本地服务
双击运行后自动打开浏览器。存档位置自动检测（Steam / SteamLibrary 各库盘）。
备份保存在 exe（或脚本）所在目录的 backups/ 下。
"""
import json, hashlib, os, re, shutil, sys, time, base64, threading, webbrowser, glob
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def project_root():
    """打包运行 = exe 同目录；脚本运行 = src/ 的上一级（工程根）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


ROOT = project_root()
BACKUP_DIR = os.path.join(ROOT, "backups")
WEB_DIR = os.path.join(ROOT, "web")

# ---------- 版本注册表：Demo / Playtest / 正式版 ----------
# 游戏基底（类别集合）不变即可自动适配；此处只登记各版本的 Steam appId、
# 游戏目录名与 %APPDATA% 存档目录名，用户可在前端自由切换版本。
EDITIONS = [
    {"id": "demo", "label": "Demo", "appId": "4777710", "isDefault": True,
     "gameDirNames": ["口袋修仙 Demo", "Pocket Cultivation Demo"],
     "appdataNames": ["Pocket Cultivation Demo", "口袋修仙 Demo"],
     "savesTokens": ["demo"], "savesExcludes": ["playtest"]},
    {"id": "playtest", "label": "Playtest", "appId": "5106860",
     "gameDirNames": ["口袋修仙 Playtest", "Pocket Cultivation Playtest"],
     "appdataNames": ["Pocket Cultivation Playtest", "口袋修仙 Playtest"],
     "savesTokens": ["playtest"], "savesExcludes": ["demo"]},
    {"id": "formal", "label": "正式版", "appId": "4707190",
     "gameDirNames": ["口袋修仙", "Pocket Cultivation"],
     "appdataNames": ["口袋修仙", "Pocket Cultivation"],
     "savesTokens": [], "savesExcludes": ["demo", "playtest"]},
]
EDITION_BY_ID = {e["id"]: e for e in EDITIONS}
DEFAULT_EDITION_ID = next((e["id"] for e in EDITIONS if e.get("isDefault")), EDITIONS[0]["id"])


# ---------- 境界表：从游戏打包 JS 实时提取（随版本自适应） ----------
def _find_realm_array_text(src):
    m = re.search(r'\[?\s*\{\s*id:\s*"mortal"', src)
    if not m:
        return None
    start = src.rfind("[", 0, m.start() + 2)
    if start < 0:
        return None
    depth = 0
    in_str = False
    esc = False
    quote = ""
    i = start
    while i < len(src):
        c = src[i]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == quote:
                in_str = False
        else:
            if c in "\"'`":
                in_str = True
                quote = c
            elif c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
                if depth == 0:
                    return src[start:i + 1]
        i += 1
    return None


def _parse_realms(arr_text):
    realms = []
    i, n = 0, len(arr_text)
    while i < n:
        if arr_text[i] == "{":
            depth = 0
            in_str = False
            esc = False
            quote = ""
            j = i
            while j < n:
                c = arr_text[j]
                if in_str:
                    if esc:
                        esc = False
                    elif c == "\\":
                        esc = True
                    elif c == quote:
                        in_str = False
                else:
                    if c in "\"'`":
                        in_str = True
                        quote = c
                    elif c == "{":
                        depth += 1
                    elif c == "}":
                        depth -= 1
                        if depth == 0:
                            break
                j += 1
            obj = arr_text[i:j + 1]
            rid = re.search(r'\bid:\s*"([^"]+)"', obj)
            rname = re.search(r'\bname:\s*"([^"]*)"', obj)
            lc = re.search(r'layerConfigs:\s*\[', obj)
            layers = []
            if lc:
                layers = re.findall(r'displayName:\s*"([^"]*)"', obj[lc.end() - 1:])
            if rid:
                realms.append({"id": rid.group(1), "name": rname.group(1) if rname else rid.group(1), "layers": layers})
            i = j + 1
        else:
            i += 1
    return realms


def find_realms(game_root):
    """从游戏 assets 中提取境界表；找不到返回内置兜底。"""
    if game_root:
        assets = _assets_root_for(game_root)
        best = None
        if assets:
            js_dir = os.path.join(assets, "assets")
            search_dirs = [os.path.join(js_dir)] if os.path.isdir(js_dir) else []
            search_dirs.append(assets)
            for d in search_dirs:
                for f in glob.glob(os.path.join(d, "*.js")):
                    try:
                        s = open(f, encoding="utf-8", errors="replace").read()
                    except Exception:
                        continue
                    if 'id:"mortal"' not in s:
                        continue
                    txt = _find_realm_array_text(s)
                    if not txt:
                        continue
                    rl = _parse_realms(txt)
                    if len(rl) >= 2 and (best is None or len(rl) > len(best)):
                        best = rl
        if best:
            return best
    return FALLBACK_REALMS


FALLBACK_REALMS = [
    {"id": "mortal", "name": "凡人", "layers": ["凡人"]},
    {"id": "qiRefining", "name": "炼气期", "layers": ["炼气期一层", "炼气期二层", "炼气期三层", "炼气期四层", "炼气期五层", "炼气期六层", "炼气期七层", "炼气期八层", "炼气期九层", "炼气期圆满"]},
    {"id": "foundation", "name": "筑基期", "layers": ["筑基初期", "筑基中期", "筑基后期", "筑基期圆满"]},
    {"id": "coreFormation", "name": "结丹期", "layers": ["结丹初期", "结丹中期", "结丹后期", "结丹期圆满"]},
    {"id": "nascentSoul", "name": "元婴期", "layers": ["元婴初期", "元婴中期", "元婴后期", "元婴期圆满"]},
    {"id": "spiritTransformation", "name": "化神期", "layers": ["化神初期", "化神中期", "化神后期", "化神期圆满"]},
    {"id": "voidRefining", "name": "炼虚期", "layers": ["炼虚初期", "炼虚中期", "炼虚后期", "炼虚期圆满"]},
    {"id": "bodyIntegration", "name": "合体期", "layers": ["合体初期", "合体中期", "合体后期", "合体期圆满"]},
    {"id": "mahayana", "name": "大乘期", "layers": ["大乘初期", "大乘中期", "大乘后期", "大乘期圆满"]},
    {"id": "tribulation", "name": "渡劫期", "layers": ["渡劫期一重天劫", "渡劫期二重天劫", "渡劫期三重天劫", "渡劫期四重天劫", "渡劫期五重天劫"]},
]


def read_game_root_hint():
    p = os.path.join(ROOT, "game-assets.txt")
    if os.path.isfile(p):
        line = open(p, encoding="utf-8-sig").read().strip().strip('"')
        if line and os.path.isdir(line):
            return line
    return None


def steam_library_roots():
    roots = []
    cands = [
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("STEAM_PATH", ""),
        r"D:\Steam", r"E:\Steam", r"F:\Steam", r"G:\Steam",
        r"D:\SteamLibrary", r"E:\SteamLibrary", r"F:\SteamLibrary", r"G:\SteamLibrary",
    ]
    for c in filter(None, cands):
        vdf = os.path.join(c, "steamapps", "libraryfolders.vdf")
        if os.path.isfile(vdf):
            try:
                for m in re.finditer(r'"path"\s+"([^"]+)"', open(vdf, encoding="utf-8", errors="replace").read()):
                    roots.append(m.group(1).replace("\\\\", "\\"))
            except Exception:
                pass
            roots.append(c)
    # de-dup preserving order
    seen, out = set(), []
    for r in roots:
        k = r.lower().rstrip("\\/")
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def _assets_root_for(game_root):
    if not game_root:
        return None
    public = os.path.join(game_root, "resources", "public")
    if os.path.isdir(public):
        return public
    if os.path.isdir(os.path.join(game_root, "generated")) or os.path.isdir(os.path.join(game_root, "assets")):
        return game_root
    return None


def _installdir_from_appmanifest(lib, app_id):
    if not app_id:
        return None
    mf = os.path.join(lib, "steamapps", "appmanifest_%s.acf" % app_id)
    if not os.path.isfile(mf):
        return None
    try:
        text = open(mf, encoding="utf-8", errors="replace").read()
    except Exception:
        return None
    m = re.search(r'"installdir"\s+"([^"]+)"', text)
    if m:
        return os.path.join(lib, "steamapps", "common", m.group(1))
    return None


def find_game_root(edition):
    """定位某版本的游戏安装根（含 resources/public 或 generated/assets）。"""
    hint = read_game_root_hint()
    if hint:
        base = os.path.basename(os.path.normpath(hint))
        if edition.get("isDefault") or base in edition["gameDirNames"]:
            if _assets_root_for(hint):
                return hint
    for lib in steam_library_roots():
        common = os.path.join(lib, "steamapps", "common")
        if not os.path.isdir(common):
            continue
        installdir = _installdir_from_appmanifest(lib, edition.get("appId"))
        candidates = ([installdir] if installdir else []) + [os.path.join(common, n) for n in edition["gameDirNames"]]
        for gdir in candidates:
            if gdir and _assets_root_for(gdir):
                return gdir
    return None


def find_saves_root(edition):
    appdata = os.environ.get("APPDATA", "")
    for n in edition["appdataNames"]:
        p = os.path.join(appdata, n, "saves")
        if os.path.isdir(p):
            return p
    # 兜底：按版本特征词匹配，避免串版本
    tokens = edition.get("savesTokens") or []
    excludes = edition.get("savesExcludes") or []
    if os.path.isdir(appdata):
        for entry in os.listdir(appdata):
            low = entry.lower()
            if "pocket cultivation" not in low and "口袋修仙" not in entry:
                continue
            if not os.path.isdir(os.path.join(appdata, entry, "saves")):
                continue
            if any(x in low for x in excludes):
                continue
            if tokens and not any(t in low for t in tokens):
                continue
            return os.path.join(appdata, entry, "saves")
    return None


def read_game_version(game_root):
    """读取游戏版本号：优先 release-artifact-manifest.json 的 releaseVersion，退回 version 文件。"""
    if not game_root:
        return None
    out = {"releaseVersion": None, "buildVersion": None, "contentEdition": None}
    mf = os.path.join(game_root, "release-artifact-manifest.json")
    if os.path.isfile(mf):
        try:
            d = json.load(open(mf, encoding="utf-8"))
            out["releaseVersion"] = d.get("releaseVersion")
            out["contentEdition"] = d.get("contentEdition")
            out["appId"] = d.get("appId")
        except Exception:
            pass
    vf = os.path.join(game_root, "version")
    if os.path.isfile(vf):
        try:
            out["buildVersion"] = open(vf, encoding="utf-8-sig").read().strip()
        except Exception:
            pass
    return out


def resolve_edition(edition_id=None):
    edition = EDITION_BY_ID.get(edition_id) if edition_id else None
    if edition is None:
        edition = EDITION_BY_ID[DEFAULT_EDITION_ID]
    game_root = find_game_root(edition)
    saves_root = find_saves_root(edition)
    return {
        "id": edition["id"], "label": edition["label"], "appId": edition.get("appId"),
        "installed": bool(game_root or saves_root),
        "gameRoot": game_root,
        "assetsRoot": _assets_root_for(game_root),
        "savesRoot": saves_root,
        "modSdk": os.path.join(game_root, "ModSDK") if game_root else None,
        "version": read_game_version(game_root),
        "realms": find_realms(game_root),
    }


_EDITION_CACHE = {}


def get_edition(edition_id=None, refresh=False):
    key = edition_id or DEFAULT_EDITION_ID
    if refresh or key not in _EDITION_CACHE:
        _EDITION_CACHE[key] = resolve_edition(edition_id)
    return _EDITION_CACHE[key]


def installed_editions(refresh=False):
    if refresh:
        _EDITION_CACHE.clear()
    return [get_edition(e["id"]) for e in EDITIONS if get_edition(e["id"])["installed"]]


# 默认版本上下文（未指定版本时使用）
_DEFAULT_CTX = get_edition(None)


# ---------- 存档哈希 ----------
def norm(r):
    if isinstance(r, list):
        return [norm(x) for x in r]
    if isinstance(r, dict):
        return {k: norm(v) for k, v in sorted(r.items(), key=lambda kv: kv[0].lower())}
    return r


def compute_hash(payload):
    s = json.dumps(norm(payload), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


# ---------- 物品目录：实时从用户本机游戏/ModSDK 读取，不自带 ----------
def _collect_labels(obj, out):
    """递归收集效果结构里的人类可读 label（label 常嵌套在 periodicPulse/passiveEffect/consumable 等子对象里）。"""
    if isinstance(obj, dict):
        lb = obj.get("label")
        if isinstance(lb, str) and lb and lb not in out:
            out.append(lb)
        for k, v in obj.items():
            if k in ("label", "starOverrides", "executionScope", "visibilityScope", "id", "kind"):
                continue  # starOverrides/executors 的 label 与主体重复，跳过避免刷屏
            _collect_labels(v, out)
    elif isinstance(obj, list):
        for x in obj:
            _collect_labels(x, out)


EFFECT_KIND_CN = {
    "periodicPulse": "轮转效果", "passive": "被动效果", "consumable": "使用效果",
    "scriptureProgression": "功法效果", "triggered": "触发效果",
}


def effect_labels(effect):
    out = []
    _collect_labels(effect, out)
    return out


def find_sdk_catalog_path(edition=None):
    """在指定版本的游戏目录旁找 ModSDK 的公开物品目录。"""
    root = (edition or _DEFAULT_CTX)["gameRoot"]
    if not root:
        return None
    candidates = [
        os.path.join(root, "ModSDK", "reference", "domains", "items", "base-item-public-catalog.json"),
        # 游戏目录内嵌 Demo SDK 的情形
        os.path.join(root, "ModSDK", "reference", "domains", "items", "base-item-index.json"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _read_mod_enums(edition):
    """读取该版本 ModSDK 的 enums.json（grades/elements/addableCategories 等），随版本自适应。"""
    root = (edition or _DEFAULT_CTX)["gameRoot"]
    if not root:
        return {}
    p = os.path.join(root, "ModSDK", "reference", "domains", "items", "enums.json")
    if not os.path.isfile(p):
        return {}
    try:
        return json.load(open(p, encoding="utf-8"))
    except Exception:
        return {}


def _walk_effects(o):
    if isinstance(o, dict):
        yield o
        for v in o.values():
            yield from _walk_effects(v)
    elif isinstance(o, list):
        for x in o:
            yield from _walk_effects(x)


def item_signals(it):
    """六维信号：元素 / Buff / 机制 / 属性 / 特征（分类与品级单独提供）。供前端交叉筛选。"""
    el = it.get("effectList") or []
    buffs, stats, mech = set(), set(), set()
    has_charge = False
    for o in _walk_effects(el):
        for k, v in o.items():
            if k in ("buffId", "sourceBuffId", "targetBuffId") and isinstance(v, str):
                buffs.add(v)
            if k in ("stat", "statId") and isinstance(v, str) and v in STAT_KEYS:
                stats.add(v)
            if k == "damageHealth" and v is not None:
                stats.add("health")
            if k == "startOnly" and v is True:
                mech.add("start")
            if k in ("neighborCounts", "directionalNeighbors", "alignments", "lineChains"):
                mech.add("adjacent")
            if k == "source" and v == "adjacentItems":
                mech.add("adjacent")
            if k == "type" and v in ("adjacent", "directional-adjacent"):
                mech.add("adjacent")
            if k in ("cultivationGain", "cultivationCost", "cultivationGainPerStack", "cultivationGainPerBuffStack") and v is not None:
                mech.add("cultivation")
            if k in ("spiritStoneGain", "spiritStoneCost", "stoneCost") and v is not None:
                mech.add("spiritStones")
            if k in ("luckyGain", "luckyCost", "luckyGainFromPrimaryStat") and v is not None:
                mech.add("lucky")
            if k == "kind" and v == "scriptureProgression":
                mech.add("scriptureBreakthrough")
            if k == "trigger" and v == "onConsumableExhausted":
                mech.add("exhaust")
            if k == "charge" and v is not None:
                has_charge = True
    for e in _walk_effects(el):
        if e.get("kind") == "consumable" or e.get("consumable") is not None:
            mech.add("exhaust")
        if e.get("kind") == "periodicPulse":
            pp = e.get("periodicPulse") or {}
            if pp.get("intervalSec") is not None and pp.get("startOnly") is not True:
                mech.add("rotation")
            costs = any(pp.get(k) for k in ("cultivationCost", "spiritStoneCost", "luckyCost", "lifespanCost")) or pp.get("primaryStatCosts")
            gains = any(pp.get(k) for k in ("cultivationGain", "spiritStoneGain", "luckyGain", "lifespanGain")) or pp.get("primaryStatGains")
            if costs and gains:
                mech.add("transmutation")
    if it.get("category") == "throwable":
        mech.add("throw")
    if it.get("category") in ("pill", "talisman"):
        mech.add("exhaust")
    if has_charge:
        mech.add("charge")
    el_str = json.dumps(el, ensure_ascii=False)
    if it.get("category") == "spell":
        for e in _walk_effects(el):
            if e.get("activeCast"):
                mech.add("cooldown")
                break
    if "formationSwitch" in el_str:
        mech.add("formationSwitch")
    if "onBuffConsumed" in el_str:
        mech.add("consumption")
    elems = set(it.get("baseElements") or [])
    if it.get("element"):
        elems.add(it["element"])
    return {
        "buffs": sorted(buffs), "stats": sorted(stats), "mech": sorted(mech),
        "elements": sorted(e for e in elems if e), "features": it.get("features") or [],
    }


STAT_KEYS = {"health", "stamina", "mana", "bloodEssence", "spiritSense"}


def find_catalog(edition=None):
    """返回已加工的物品目录（含图标相对路径）。每次请求时实时读取对应版本游戏数据。"""
    edition = edition or _DEFAULT_CTX
    sdk_path = find_sdk_catalog_path(edition)
    assets = edition["assetsRoot"]
    if not sdk_path or not assets:
        return []
    try:
        raw = json.load(open(sdk_path, encoding="utf-8"))
    except Exception:
        return []
    items = raw["items"] if isinstance(raw, dict) and "items" in raw else raw
    icons_dir = os.path.join(assets, "assets", "generated", "item-icons")
    bags_dir = os.path.join(assets, "assets", "generated", "bag-backgrounds")
    icons = set(os.listdir(icons_dir)) if os.path.isdir(icons_dir) else set()
    bags = set(os.listdir(bags_dir)) if os.path.isdir(bags_dir) else set()
    out = []
    for it in items:
        base = it.get("iconBasename", "")
        icon = None
        if (base + ".webp") in icons:
            icon = "assets/generated/item-icons/" + base + ".webp"
        elif (base + ".webp") in bags:
            icon = "assets/generated/bag-backgrounds/" + base + ".webp"
        effs = []
        for e in (it.get("effectList") or []):
            if not isinstance(e, dict):
                continue
            labels = effect_labels(e)
            if labels:
                effs.extend(labels)
            else:
                kind = e.get("kind", "效果")
                effs.append(EFFECT_KIND_CN.get(kind, kind))
        deduped = []
        for x in effs:
            if x not in deduped:
                deduped.append(x)
        sig = item_signals(it)
        out.append({
            "id": it["numericId"], "name": it.get("name"), "cat": it.get("category"),
            "grade": it.get("grade"), "elem": it.get("element"),
            "shape": it.get("shape"), "features": it.get("features") or [],
            "tags": it.get("tags") or [],
            "obtainable": it.get("isObtainable", False), "published": it.get("isPublished", True),
            "icon": icon, "price": it.get("price"),
            "desc": (it.get("description") or "")[:120],
            "effs": [x for x in effs if x][:6],
            "sig": sig,
        })
    out.sort(key=lambda x: x["id"])
    return out


def catalog_payload(edition=None):
    """物品目录 + 该版本实际类别/品级/元素集合（前端据此动态生成筛选，基底不变即自动适配）。
    另附六维子分类可选值（facets），供分类下方交叉筛选。"""
    edition = edition or _DEFAULT_CTX
    items = find_catalog(edition)
    cats = sorted({i["cat"] for i in items if i.get("cat")})
    grades = sorted({i["grade"] for i in items if i.get("grade")})
    elems = sorted({i["elem"] for i in items if i.get("elem")})

    def collect(field):
        s = set()
        for i in items:
            sig = i.get("sig") or {}
            for v in (sig.get(field) or []):
                if v:
                    s.add(v)
        return sorted(s)

    facets = {
        "elem": sorted({e for i in items for e in ((i.get("sig") or {}).get("elements") or []) if e}),
        "buff": collect("buffs"),
        "mech": collect("mech"),
        "stat": collect("stats"),
        "feature": sorted({f for i in items for f in (i.get("features") or []) if f}),
        "grade": grades,
    }
    return {
        "edition": edition["id"],
        "items": items,
        "categories": cats,
        "grades": grades,
        "elements": elems,
        "facets": facets,
    }


def list_saves(edition=None):
    edition = edition or _DEFAULT_CTX
    saves_root = edition["savesRoot"]
    result = []
    if not saves_root or not os.path.isdir(saves_root):
        return result
    for steam in sorted(os.listdir(saves_root)):
        sdir = os.path.join(saves_root, steam)
        if not os.path.isdir(sdir) or steam.startswith("_"):
            continue
        for char in sorted(os.listdir(sdir)):
            cdir = os.path.join(sdir, char)
            if not os.path.isdir(cdir):
                continue
            for f in sorted(os.listdir(cdir)):
                if not f.endswith(".json") or f.endswith(".bak"):
                    continue
                p = os.path.join(cdir, f)
                try:
                    d = json.load(open(p, encoding="utf-8"))
                except Exception:
                    continue
                if not isinstance(d, dict) or "snapshot" not in d or d.get("characterId") is None:
                    continue
                snap = d.get("snapshot", {})
                result.append({
                    "path": p,
                    "characterId": d.get("characterId"),
                    "saveId": d.get("saveId"),
                    "saveKind": d.get("saveKind"),
                    "roleName": d.get("roleName"),
                    "lastSaveTime": d.get("lastLocalSaveTime"),
                    "spiritStones": snap.get("spiritStones"),
                    "itemCount": len(snap.get("inventoryItems", [])),
                    "placedCount": len(snap.get("placedItems", [])),
                })
    result.sort(key=lambda x: x.get("lastSaveTime") or 0, reverse=True)
    return result




def edition_for_save_path(path):
    """由存档路径判断所属版本。"""
    try:
        rp = os.path.realpath(path)
    except Exception:
        return None
    for e in EDITIONS:
        root = get_edition(e["id"])["savesRoot"]
        if root:
            try:
                if rp.startswith(os.path.realpath(root)):
                    return e["id"]
            except Exception:
                continue
    return None


def realm_display_name(realm_id, realm_layer, edition_id=None):
    """按对应版本的境界表推导层名（等价游戏 getRealmDisplayName）。"""
    if not realm_id:
        return None
    rl = None
    ctx = get_edition(edition_id) if edition_id else None
    realms = (ctx or {}).get("realms") if ctx else None
    if not realms:
        realms = FALLBACK_REALMS
    for r in realms:
        if r["id"] == realm_id:
            rl = r
            break
    if not rl or not rl.get("layers"):
        return None
    idx = min(max(int(realm_layer or 1) - 1, 0), len(rl["layers"]) - 1)
    return rl["layers"][idx]


def sync_catalog_summary(save_path, d):
    """同步存档同级目录 catalog.json 的摘要（游戏存档列表读它，含 roleName/境界/灵根）。
    匹配 saveId；无 catalog 或未匹配则静默跳过。"""
    steam_dir = os.path.dirname(os.path.dirname(save_path))
    cat_path = os.path.join(steam_dir, "catalog.json")
    if not os.path.isfile(cat_path):
        return False
    try:
        cat = json.load(open(cat_path, encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(cat, dict) or not isinstance(cat.get("summaries"), list):
        return False
    snap = d.get("snapshot", {})
    sid = d.get("saveId")
    edition_id = edition_for_save_path(save_path)
    changed = False
    for s in cat["summaries"]:
        if s.get("saveId") != sid:
            continue
        s["roleName"] = d.get("roleName", s.get("roleName"))
        s["daoCode"] = d.get("daoCode", s.get("daoCode"))
        s["lastLocalSaveTime"] = d.get("lastLocalSaveTime", s.get("lastLocalSaveTime"))
        s["rulesVersion"] = d.get("rulesVersion", s.get("rulesVersion"))
        if isinstance(snap.get("spiritRoots"), dict):
            s["spiritRoots"] = snap["spiritRoots"]
        rid = snap.get("realmId", s.get("realmId"))
        rlayer = snap.get("realmLayer", s.get("realmLayer"))
        s["realmId"] = rid
        s["realmLayer"] = rlayer
        name = realm_display_name(rid, rlayer, edition_id)
        if name:
            s["realmName"] = name
        changed = True
    if changed:
        try:
            tmp = cat_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(json.dumps(cat, ensure_ascii=False, separators=(",", ":")))
            os.replace(tmp, cat_path)
        except Exception:
            return False
    return changed


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=WEB_DIR, **kw)

    def log_message(self, fmt, *args):
        pass

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self):
        n = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(n).decode("utf-8"))

    def _edition_id(self, parsed):
        """从 query(edition=) 或请求头取版本；非法/缺省 → 默认版本。"""
        q = parse_qs(parsed.query)
        eid = (q.get("edition", [""])[0] or self.headers.get("X-Pocket-Edition") or "").strip()
        return eid if eid in EDITION_BY_ID else None

    def _ctx(self, parsed):
        return get_edition(self._edition_id(parsed))

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        # index.html: in frozen exe it lives in _MEIPASS, not next to the exe
        if path in ("/", "/index.html"):
            html = None
            meipass = getattr(sys, "_MEIPASS", None)
            candidates = []
            if meipass:
                candidates += [os.path.join(meipass, "web", "index.html"), os.path.join(meipass, "index.html")]
            candidates += [os.path.join(WEB_DIR, "index.html"), os.path.join(ROOT, "index.html")]
            for cand in candidates:
                if os.path.isfile(cand):
                    html = open(cand, "rb").read()
                    break
            if html is None:
                self.send_response(404)
                self.end_headers()
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(html)))
            self.end_headers()
            self.wfile.write(html)
            return
        if path == "/api/versions":
            q = parse_qs(parsed.query)
            refresh = (q.get("refresh", [""])[0] == "1")
            ctx = self._ctx(parsed)
            versions = []
            for e in EDITIONS:
                c = get_edition(e["id"], refresh=refresh)
                versions.append({
                    "id": e["id"], "label": e["label"], "appId": e.get("appId"),
                    "installed": c["installed"],
                    "gameRoot": c["gameRoot"], "savesRoot": c["savesRoot"],
                    "assetsRoot": c["assetsRoot"],
                    "version": c.get("version"),
                    "realms": c.get("realms"),
                })
            return self.send_json({"current": ctx["id"], "default": DEFAULT_EDITION_ID, "versions": versions})
        if path == "/api/saves":
            ctx = self._ctx(parsed)
            return self.send_json({"saves": list_saves(ctx), "savesRoot": ctx["savesRoot"] or "(未找到)",
                                   "edition": ctx["id"]})
        if path == "/api/catalog":
            ctx = self._ctx(parsed)
            data = json.dumps(catalog_payload(ctx), ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path == "/api/save":
            q = parse_qs(parsed.query)
            p = q.get("path", [""])[0]
            if not self._safe_save_path(p):
                return self.send_json({"error": "非法路径"}, 400)
            if not os.path.isfile(p):
                return self.send_json({"error": "存档不存在"}, 404)
            d = json.load(open(p, encoding="utf-8"))
            return self.send_json({"save": d, "hashOk": self._check_hash(d)})
        if path == "/api/status":
            ctx = self._ctx(parsed)
            lock = os.path.join(os.path.dirname(ctx["savesRoot"] or ""), "lockfile")
            running = False
            if lock and os.path.isfile(lock):
                running = (time.time() - os.path.getmtime(lock)) < 10
            return self.send_json({"gameRunning": running, "assetsFound": bool(ctx["assetsRoot"]),
                                   "savesFound": bool(ctx["savesRoot"]), "edition": ctx["id"]})
        # game assets (icons)
        if path.startswith("/game/"):
            ctx = self._ctx(parsed)
            assets = ctx["assetsRoot"]
            rel = path[len("/game/"):].lstrip("/\\")
            fp = os.path.normpath(os.path.join(assets or "", rel.replace("/", os.sep)))
            base = os.path.normpath(assets) if assets else ""
            inside = bool(base) and (fp == base or fp.startswith(base + os.sep))
            if not inside or not os.path.isfile(fp) or fp.endswith(".bak"):
                self.send_response(404)
                self.end_headers()
                return
            ctype = "image/webp" if fp.endswith(".webp") else ("image/png" if fp.endswith(".png") else "application/octet-stream")
            data = open(fp, "rb").read()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "max-age=86400")
            self.end_headers()
            self.wfile.write(data)
            return
        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/save":
            body = self.read_body()
            p = body.get("path", "")
            if not self._safe_save_path(p):
                return self.send_json({"error": "非法路径"}, 400)
            d = body.get("save")
            if not isinstance(d, dict):
                return self.send_json({"error": "缺少 save 数据"}, 400)
            payload = {k: v for k, v in d.items() if k not in ("snapshotHash", "localSaveHash")}
            h = compute_hash(payload)
            d["snapshotHash"] = h
            d["localSaveHash"] = h
            # 备份：先将选定存档复制到 backups/<日期>/<存档文件名>，再覆盖原存档
            backup_rel = None
            if os.path.isfile(p):
                day = time.strftime("%Y-%m-%d")
                day_dir = os.path.join(BACKUP_DIR, day)
                os.makedirs(day_dir, exist_ok=True)
                name = os.path.basename(p)
                dest = os.path.join(day_dir, name)
                if os.path.exists(dest):  # 同日多次保存同一存档，追加时间避免覆盖
                    stem, ext = os.path.splitext(name)
                    dest = os.path.join(day_dir, "%s.%s%s" % (stem, time.strftime("%H%M%S"), ext))
                shutil.copy2(p, dest)
                backup_rel = os.path.relpath(dest, BACKUP_DIR).replace("\\", "/")
            tmp = p + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(json.dumps(d, ensure_ascii=False, separators=(",", ":")))
            os.replace(tmp, p)
            # 同步游戏存档列表摘要（catalog.json）
            try:
                sync_catalog_summary(p, d)
            except Exception:
                pass
            return self.send_json({"ok": True, "hash": h, "backup": backup_rel})
        return self.send_json({"error": "unknown"}, 404)

    def _safe_save_path(self, p):
        if not p:
            return False
        try:
            rp = os.path.realpath(p)
        except Exception:
            return False
        if not rp.endswith(".json"):
            return False
        for e in EDITIONS:
            root = get_edition(e["id"])["savesRoot"]
            if root:
                try:
                    if rp.startswith(os.path.realpath(root)):
                        return True
                except Exception:
                    continue
        return False

    def _check_hash(self, d):
        payload = {k: v for k, v in d.items() if k not in ("snapshotHash", "localSaveHash")}
        return compute_hash(payload) == d.get("snapshotHash") == d.get("localSaveHash")


def open_browser(port):
    try:
        webbrowser.open(f"http://127.0.0.1:{port}")
    except Exception:
        pass


def main():
    port = 8765
    print("口袋修仙 存档工坊")
    for e in EDITIONS:
        ctx = get_edition(e["id"])
        mark = "已安装" if ctx["installed"] else "未安装"
        print(f"  [{mark}] {e['label']:<8} 游戏: {ctx['gameRoot'] or '(未找到)'}")
        print(f"            存档: {ctx['savesRoot'] or '(未找到)'}")
    print(f"备份目录: {BACKUP_DIR}")
    url = f"http://127.0.0.1:{port}"
    print(f"打开 {url} （浏览器已自动打开，可在页面顶部切换版本）")
    threading.Timer(0.8, open_browser, (port,)).start()
    try:
        HTTPServer(("127.0.0.1", port), Handler).serve_forever()
    except OSError as e:
        if "10048" in str(e) or "in use" in str(e).lower():
            print("端口被占用，可能已有一个实例在运行。直接使用已打开的页面即可。")
            open_browser(port)
        else:
            raise


if __name__ == "__main__":
    main()