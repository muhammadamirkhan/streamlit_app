import streamlit as st
import pandas as pd
import os
import json
from io import BytesIO

st.set_page_config(page_title="Muraba Veil – Unit Manager", layout="wide", page_icon="🏙️")

# ── Password gate ──────────────────────────────────────────────────────────────
def _check_password() -> bool:
    expected = st.secrets.get("password", os.environ.get("APP_PASSWORD", "muraba2026"))

    def _entered():
        st.session_state["auth_ok"] = st.session_state.get("pw_input", "") == expected
        st.session_state.pop("pw_input", None)

    if st.session_state.get("auth_ok"):
        return True

    st.markdown("## 🔒 Muraba Veil — Unit Manager")
    st.text_input("Enter password to continue", type="password",
                  on_change=_entered, key="pw_input")
    if st.session_state.get("auth_ok") is False:
        st.error("Incorrect password — try again.")
    st.caption("Access is restricted. Contact the administrator for the password.")
    return False

if not _check_password():
    st.stop()

# ── Data file (lives next to this script, so it works locally and on the cloud) ─
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Muraba Veil Unit list.xlsx")
COMMENTS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unit_comments.json")


def comment_key(unit, type_, floor):
    return f"{unit}|{type_}|{floor}"

def load_comments_file() -> dict:
    try:
        with open(COMMENTS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_comments_file(mapping: dict):
    try:
        with open(COMMENTS_PATH, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def persist_all_comments():
    """Write every non-empty comment in the register to the JSON file, keyed by unit/type/floor."""
    u = st.session_state.units
    mapping = {}
    for _, r in u.iterrows():
        c = r.get("Comment", "")
        if isinstance(c, str) and c.strip():
            mapping[comment_key(r["Unit"], r["Type"], r["Floor"])] = c
    save_comments_file(mapping)

UNIT_TYPES = [
    "2 Bedroom", "3 Bedroom - New", "3 Bedroom", "3 Bedroom Pool", "4 Bedroom Pool",
    "4 Bedroom XL", "3 Bedroom Duplex", "4 Bedroom Duplex", "5 Bedroom Duplex",
]
STATUS_OPTIONS = ["Available", "Sold"]

# "3 Bedroom - New" is a 2-Bedroom-sized unit under a new name; it shares the
# 2 Bedroom price ladder, area, terrace and escalation (same stats).
PRICE_FAMILY = {"3 Bedroom - New": "2 Bedroom"}

def family_types(t):
    """All types that share a price ladder with t (e.g. 3 Bedroom - New ↔ 2 Bedroom)."""
    base = PRICE_FAMILY.get(t, t)
    fam = {t, base}
    fam.update(k for k, v in PRICE_FAMILY.items() if v == base)
    return fam

TYPE_DEFAULTS = {
    "2 Bedroom":         {"internal": 2218.764851, "external": 1619.322681, "parking": 2, "terrace_rate": 0.30, "levels": 1},
    "3 Bedroom - New":   {"internal": 2218.764851, "external": 1619.322681, "parking": 2, "terrace_rate": 0.30, "levels": 1},
    "3 Bedroom":         {"internal": 2880.530062, "external": 2058.920781, "parking": 2, "terrace_rate": 0.30, "levels": 1},
    "3 Bedroom Pool":    {"internal": 2880.530062, "external": 2059.243699, "parking": 2, "terrace_rate": 0.65, "levels": 2},
    "4 Bedroom Pool":    {"internal": 4643.550947, "external": 5258.816065, "parking": 3, "terrace_rate": 0.55, "levels": 2},
    "4 Bedroom XL": {"internal": 7474.889938, "external": 7857.654592, "parking": 4, "terrace_rate": 0.65, "levels": 1},
    "3 Bedroom Duplex":  {"internal": 4733.537238, "external": 3334.228886, "parking": 3, "terrace_rate": 0.75, "levels": 2},
    "4 Bedroom Duplex":  {"internal": 7485.653849, "external": 7260.042287, "parking": 4, "terrace_rate": 0.75, "levels": 2},
    "5 Bedroom Duplex":  {"internal": 11648.17,    "external": 15018.56,    "parking": 6, "terrace_rate": 1.00, "levels": 2},
}

LEVEL_CAPACITY = {"2 Bedroom": 2, "3 Bedroom": 1}   # standard residential floor

# Fallback escalation defaults (overridden by what we read from the sheets)
ESC_DEFAULTS = {
    "2 Bedroom": 150.0, "3 Bedroom - New": 150.0, "3 Bedroom": 150.0,
    "3 Bedroom Pool": 104.0, "4 Bedroom Pool": 104.0,
    "4 Bedroom XL": 497.0, "3 Bedroom Duplex": 308.0,
    "4 Bedroom Duplex": 305.0, "5 Bedroom Duplex": 0.0,
}
# Terrace-rate groups (% of internal), variable; from Launches XL SX DX
TERRACE_DEFAULTS = {
    "standard": 0.30,        # 2BR & 3BR
    "3 Bedroom Pool": 0.65,  # 3 Pool Terrace
    "4 Bedroom Pool": 0.55,  # 4 Pool Terrace Rate
    "duplex": 0.75,          # DX Terrace Rate (3/4/5 BR Duplex)
    "simplex": 0.65,         # SX Terrace Rate (4 BR Simplex / XL)
}

BLUE_DARK = "#1F4E78"
BLUE_MED  = "#2E75B6"
BLUE_LITE = "#DDEBF7"

SQFT_PER_SQM = 10.7639   # 1 m² = 10.7639 ft²  →  sqm = sqft / 10.7639


# ── Excel-style table renderer (first column left, numeric columns centered) ───

def excel_table(df: pd.DataFrame):
    sty = (df.style.hide(axis="index").set_table_styles([
        {"selector": "", "props": "border-collapse:collapse;font-size:13px;width:100%;"
                                   "font-family:Calibri,Arial,sans-serif;"},
        {"selector": "thead th", "props": f"background-color:{BLUE_DARK};color:#FFFFFF;font-weight:bold;"
                                           "text-align:center;border:1px solid #9DC3E6;padding:6px 10px;"},
        {"selector": "tbody td", "props": "border:1px solid #BDD7EE;padding:5px 10px;text-align:center;"},
        {"selector": "tbody td:first-child", "props": "text-align:left;font-weight:600;"},
        {"selector": "thead th:first-child", "props": "text-align:left;"},
        {"selector": "tbody tr:nth-child(even)", "props": f"background-color:{BLUE_LITE};"},
        {"selector": "tbody tr:nth-child(odd)",  "props": "background-color:#FFFFFF;"},
    ]))
    st.markdown(f'<div style="overflow-x:auto">{sty.to_html()}</div>', unsafe_allow_html=True)


def aed(x):  return f"AED {x:,.0f}"

def ordinal(n):
    """13 -> '13th', 21 -> '21st'. Accepts ints or numeric strings; passes through non-numeric."""
    try:
        n = int(float(n))
    except (ValueError, TypeError):
        return str(n)
    suf = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"

def area_fmt(x, sqm=False):
    v = x / SQFT_PER_SQM if sqm else x
    return f"{v:,.0f}"


def column_picker(all_cols, key, locked=None, hidden_default=None):
    """Dropdown (popover) with a multi-select to show / hide table columns.

    `locked` columns are always shown and cannot be unticked.
    `hidden_default` columns start unticked (hidden) but can be shown.
    Returns the ordered list of column names to display.
    """
    all_cols = list(all_cols)
    locked = locked or []
    hidden_default = hidden_default or []
    selectable = [c for c in all_cols if c not in locked]
    default = [c for c in selectable if c not in hidden_default]
    with st.popover("🔧 Columns", use_container_width=False):
        st.caption("Tick to show, untick to hide. Key columns are always shown.")
        chosen = st.multiselect("Columns", selectable, default=default,
                                key=key, label_visibility="collapsed")
    return [c for c in all_cols if c in locked or c in chosen]


# ── Data loading ───────────────────────────────────────────────────────────────

def load_unit_data() -> pd.DataFrame:
    raw  = pd.read_excel(EXCEL_PATH, sheet_name="Muraba Unit Wise Details ", header=None)
    data = raw.iloc[3:].copy().reset_index(drop=True)
    df = pd.DataFrame({
        "Type":          data[0],
        "Status":        data[1],
        "Unit":          data[2].astype(str),
        "Floor":         data[3].astype(str),
        "Parking":       pd.to_numeric(data[4], errors="coerce").fillna(0).astype(int),
        "Internal_sqft": pd.to_numeric(data[5], errors="coerce"),
        "External_sqft": pd.to_numeric(data[6], errors="coerce"),
        "Terrace_Rate":  pd.to_numeric(data[8], errors="coerce"),
        "Price_sqft":    pd.to_numeric(data[10], errors="coerce"),
    })
    df = df[df["Type"].notna() & (df["Type"] != "Total")].reset_index(drop=True)
    df["Type"]   = df["Type"].replace("4 Bedroom Simplex", "4 Bedroom XL")  # renamed typology
    df["Floor"]  = df["Floor"].apply(ordinal)                          # normalise 33 / 33.0 / "4th" -> "33rd" / "4th"
    df["Status"] = df["Status"].replace("Bank Locked", "Available")   # Bank Locked reclassified as Available

    # Fill any missing Price/sqft from the nearest same-type floor (source-data gaps, e.g. unit 6802).
    if df["Price_sqft"].isna().any():
        fn = pd.to_numeric(df["Floor"].str.replace(r"[^0-9]", "", regex=True), errors="coerce")
        for t, idx in df.groupby("Type").groups.items():
            order = fn.loc[idx].sort_values().index            # same-type rows in floor order
            filled = df.loc[order, "Price_sqft"].ffill().bfill()
            df.loc[order, "Price_sqft"] = filled
        df["Price_sqft"] = df["Price_sqft"].fillna(0.0)        # last-resort guard

    df["Terrace_Override"] = pd.NA                                     # per-unit terrace-rate override (set by bulk tool)
    cmts = load_comments_file()                                       # free-text notes per unit, persisted to JSON
    df["Comment"] = [cmts.get(comment_key(u, t, fl), "")
                     for u, t, fl in zip(df["Unit"], df["Type"], df["Floor"])]
    df["uid"] = [f"u{i}" for i in range(len(df))]   # stable unique row id (unit numbers are NOT unique)
    return df


def load_floor_data(units_df: pd.DataFrame) -> list:
    # link Launches-sheet units to register rows by (unit_no, type) — unique in the standard range
    umap = {(str(r["Unit"]), r["Type"]): r["uid"] for _, r in units_df.iterrows()}
    raw = pd.read_excel(EXCEL_PATH, sheet_name="Launches Residences", header=None)
    floors = []
    for i in range(14, 66):
        if i >= len(raw):
            break
        row = raw.iloc[i]
        try:
            fnum = int(float(row[6]))
        except (ValueError, TypeError):
            continue
        if fnum == 0:
            continue
        units = []
        def mk(col_no, col_rate, typ):
            no = str(int(float(row[col_no])))
            return {"unit_no": no, "type": typ, "rate": float(row[col_rate]),
                    "uid": umap.get((no, typ))}
        if pd.notna(row[3]) and pd.notna(row[7]):
            units.append(mk(3, 7, "3 Bedroom"))
        if pd.notna(row[4]) and pd.notna(row[8]):
            units.append(mk(4, 8, "2 Bedroom"))
        if pd.notna(row[5]) and pd.notna(row[9]):
            units.append(mk(5, 9, "2 Bedroom"))
        floors.append({"floor": fnum, "kind": "Standard", "levels": 1, "units": units})
    return floors


def load_pool_floors(units_df: pd.DataFrame) -> list:
    pool = units_df[units_df["Type"].isin(["3 Bedroom Pool", "4 Bedroom Pool"])].copy()
    pool["fn"] = pd.to_numeric(
        pool["Floor"].str.replace("th", "").str.replace("st", "")
                     .str.replace("nd", "").str.replace("rd", "").str.strip(), errors="coerce")
    floors = []
    for fnum, grp in pool.groupby("fn"):
        if pd.isna(fnum):
            continue
        units = [{"unit_no": str(r["Unit"]), "type": r["Type"], "rate": float(r["Price_sqft"]),
                  "uid": r["uid"]} for _, r in grp.iterrows()]
        floors.append({"floor": int(fnum), "kind": "Pool", "levels": 2, "units": units})
    return floors


def build_floor_list(units_df: pd.DataFrame) -> list:
    """Build the floor list directly from the register so EVERY unit is represented
    (Standard, Pool, Duplex, XL, Penthouse). This guarantees the Floor Manager grand
    total matches the portfolio value — earlier it only read the Standard+Pool sheets
    and silently dropped the 4 Duplex/XL/PH units (~AED 825M)."""
    fnum = pd.to_numeric(units_df["Floor"].astype(str).str.replace(r"[^0-9]", "", regex=True),
                         errors="coerce")
    buckets = {}
    for pos in range(len(units_df)):
        f = fnum.iloc[pos]
        if pd.isna(f):
            continue
        r = units_df.iloc[pos]
        buckets.setdefault(int(f), []).append({
            "unit_no": str(r["Unit"]), "type": r["Type"],
            "rate": float(r["Price_sqft"]), "uid": r["uid"],
        })
    out = []
    for f in sorted(buckets):
        units = buckets[f]
        types = [u["type"] for u in units]
        levels = max(TYPE_DEFAULTS.get(t, {"levels": 1})["levels"] for t in types)
        if   any("Pool" in t for t in types):    kind = "Pool"
        elif any("Duplex" in t for t in types):  kind = "Duplex"
        elif any(t == "4 Bedroom XL" for t in types): kind = "XL"
        else:                                     kind = "Standard"
        out.append({"floor": f, "kind": kind, "levels": levels, "units": units})
    return out


def load_params() -> dict:
    esc     = dict(ESC_DEFAULTS)
    terrace = dict(TERRACE_DEFAULTS)
    duplex_premium = 0.0
    try:
        lr = pd.read_excel(EXCEL_PATH, sheet_name="Launches Residences", header=None)
        esc["2 Bedroom"]     = float(lr.iloc[3][20])
        esc["3 Bedroom"]     = float(lr.iloc[4][20])
        terrace["standard"]  = float(lr.iloc[6][20])
    except Exception:
        pass
    try:
        xl = pd.read_excel(EXCEL_PATH, sheet_name="Launches   XL   SX   DX", header=None)
        esc["3 Bedroom Pool"]    = float(xl.iloc[1][16])
        esc["4 Bedroom Pool"]    = float(xl.iloc[2][16])
        esc["3 Bedroom Duplex"]  = float(xl.iloc[4][16])
        esc["4 Bedroom XL"] = float(xl.iloc[5][16])
        esc["4 Bedroom Duplex"]  = float(xl.iloc[6][16])
        terrace["3 Bedroom Pool"] = float(xl.iloc[7][16])
        terrace["4 Bedroom Pool"] = float(xl.iloc[8][16])
        terrace["duplex"]         = float(xl.iloc[9][16])
        terrace["simplex"]        = float(xl.iloc[10][16])
        dp = xl.iloc[11][16]
        duplex_premium = float(dp) if pd.notna(dp) else 0.0
    except Exception:
        pass
    esc["3 Bedroom - New"] = esc["2 Bedroom"]   # new type mirrors 2 Bedroom escalation
    area = {t: {"internal": TYPE_DEFAULTS[t]["internal"], "external": TYPE_DEFAULTS[t]["external"]}
            for t in UNIT_TYPES}
    return {"escalation": esc, "terrace": terrace, "duplex_premium": duplex_premium, "area": area}


def load_blocked_floors() -> dict:
    """MEP / Majlis levels from the building view — these floor numbers cannot be added."""
    blocked = {}
    try:
        bv = pd.read_excel(EXCEL_PATH, sheet_name="Building  Usman", header=None)
        for _, row in bv.iterrows():
            lvl, desc = row[2], row[3]
            if pd.isna(lvl) or pd.isna(desc):
                continue
            try:
                lvl_int = int(float(lvl))
            except (ValueError, TypeError):
                continue
            d = str(desc).upper()
            if "MEP" in d or "MAJILIS" in d or "MAJLIS" in d:
                blocked[lvl_int] = str(desc).strip()
    except Exception:
        pass
    return blocked


# ── Calculations ───────────────────────────────────────────────────────────────

def terrace_for(t, params):
    tr = params["terrace"]
    if t in ("2 Bedroom", "3 Bedroom - New", "3 Bedroom"): return tr["standard"]
    if t == "3 Bedroom Pool":                           return tr["3 Bedroom Pool"]
    if t == "4 Bedroom Pool":                           return tr["4 Bedroom Pool"]
    if t == "4 Bedroom XL":                        return tr["simplex"]
    if t in ("3 Bedroom Duplex", "4 Bedroom Duplex"):   return tr["duplex"]
    # 5 Bedroom Duplex (penthouse) keeps its own 100% terrace
    return TYPE_DEFAULTS[t]["terrace_rate"]

def escalation_for(t, params):
    return params["escalation"].get(t, 0.0)

def last_available_price(t, units_df):
    """Return the Price_sqft of the highest-floor Available unit for type t (floor-sequence aware).
    Pools price-family types together (3 Bedroom - New rides the 2 Bedroom ladder)."""
    fam = family_types(t)
    sub = units_df[(units_df["Type"].isin(fam)) & (units_df["Status"] == "Available")].copy()
    if sub.empty:
        sub = units_df[units_df["Type"].isin(fam)].copy()
    if sub.empty:
        return 5000.0
    # sort by numeric floor so escalation always references the topmost available unit
    sub["_fnum"] = pd.to_numeric(
        sub["Floor"].str.replace(r"[^0-9]", "", regex=True), errors="coerce")
    sub = sub.sort_values("_fnum")
    return float(sub.iloc[-1]["Price_sqft"])

def _fnum_series(units_df):
    return pd.to_numeric(units_df["Floor"].astype(str).str.replace(r"[^0-9]", "", regex=True),
                         errors="coerce")

def escalation_reference(t, target_floor, units_df):
    """Pick the escalation reference for a new/edited unit of type t on target_floor.

    Floor-direction aware (prices always rise with height):

    • If type-t units exist BELOW target → price UP from below:
        A = highest-floor Available unit below; candidates = A + units between A and target;
        reference R = the HIGHEST-priced candidate; steps = distinct floors with a type-t
        unit strictly between R and target, + 1.  rate = R.psf + esc × steps.

    • Else if type-t units exist ABOVE target → price DOWN from above:
        A = lowest-floor Available unit above; candidates = A + units between target and A;
        reference R = the LOWEST-priced candidate; steps = distinct floors with a type-t
        unit strictly between target and R, + 1.  rate = R.psf − esc × steps.

    Returns ({ref_unit, ref_floor, ref_psf, steps}, direction) where direction is +1 / −1,
    or (None, 0) when no comparable type-t unit exists anywhere.
    """
    sub = units_df[units_df["Type"].isin(family_types(t))].copy()   # pool 3BR-New with 2BR ladder
    if sub.empty:
        return None, 0
    sub["fn"] = _fnum_series(sub)
    sub = sub.dropna(subset=["fn"])
    below = sub[sub["fn"] < target_floor]
    above = sub[sub["fn"] > target_floor]

    if not below.empty:                                    # ── price UP from below ──
        avail = below[below["Status"] == "Available"]
        if not avail.empty:
            a_floor = avail["fn"].max()
            cand = below[below["fn"] >= a_floor]
        else:
            cand = below
        R = cand.loc[cand["Price_sqft"].astype(float).idxmax()]
        r_floor = float(R["fn"])
        steps = int(sub[(sub["fn"] > r_floor) & (sub["fn"] < target_floor)]["fn"].nunique()) + 1
        return {"ref_unit": str(R["Unit"]), "ref_floor": int(r_floor),
                "ref_psf": float(R["Price_sqft"]), "steps": steps}, +1

    if not above.empty:                                    # ── price DOWN from above ──
        avail = above[above["Status"] == "Available"]
        if not avail.empty:
            a_floor = avail["fn"].min()
            cand = above[above["fn"] <= a_floor]
        else:
            cand = above
        R = cand.loc[cand["Price_sqft"].astype(float).idxmin()]
        r_floor = float(R["fn"])
        steps = int(sub[(sub["fn"] < r_floor) & (sub["fn"] > target_floor)]["fn"].nunique()) + 1
        return {"ref_unit": str(R["Unit"]), "ref_floor": int(r_floor),
                "ref_psf": float(R["Price_sqft"]), "steps": steps}, -1

    return None, 0

def new_unit_rate(t, target_floor, units_df, params):
    esc = escalation_for(t, params)
    ref, direction = escalation_reference(t, target_floor, units_df)
    if ref is None:
        rate = last_available_price(t, units_df) + esc
    else:
        rate = ref["ref_psf"] + direction * esc * ref["steps"]
    if "Duplex" in t:
        rate += params.get("duplex_premium", 0.0)
    return max(rate, 0.0)

def reladder_typology(t, params):
    """Re-price every **Available** unit of the price-family up the ladder using the current
    escalation (3 Bedroom - New is re-laddered together with 2 Bedroom). Anchors — units with
    no comparable unit below them (the entry price) — keep their price. Sold units are never
    touched and act as fixed reference points. Mutates st.session_state.units.
    Returns the number of units whose price changed."""
    fam = family_types(t)
    u = st.session_state.units.copy()
    fn = _fnum_series(u)
    famask = u["Type"].isin(fam)
    order = (u[famask].assign(_fn=fn[famask])
             .dropna(subset=["_fn"]).sort_values("_fn").index)
    changed = 0
    for idx in order:                       # bottom-up so each unit sees updated floors below it
        if u.at[idx, "Status"] != "Available":
            continue
        tt = u.at[idx, "Type"]
        tf = float(fn.loc[idx])
        ref, direction = escalation_reference(tt, tf, u)
        if ref is None or direction < 0:    # nothing below → this is the anchor, keep its price
            continue
        rate = ref["ref_psf"] + escalation_for(tt, params) * ref["steps"]
        if "Duplex" in tt:
            rate += params.get("duplex_premium", 0.0)
        rate = max(rate, 0.0)
        if abs(float(u.at[idx, "Price_sqft"]) - rate) > 1e-9:
            u.at[idx, "Price_sqft"] = rate
            changed += 1
    st.session_state.units = u
    return changed

def sync_floor_rates():
    """Push current register prices back onto the floor-objects so the Floors table,
    totals and export stay consistent after a reprice."""
    pmap = dict(zip(st.session_state.units["uid"], st.session_state.units["Price_sqft"]))
    for fl in st.session_state.floors:
        for un in fl["units"]:
            if un.get("uid") in pmap:
                un["rate"] = float(pmap[un["uid"]])

def area_for(t, params):
    a = params.get("area", {}).get(t)
    if a:
        return a["internal"], a["external"]
    return TYPE_DEFAULTS[t]["internal"], TYPE_DEFAULTS[t]["external"]

def unit_val(t, rate, params):
    internal, external = area_for(t, params)
    tr = terrace_for(t, params)
    return {"internal": internal*rate, "terrace": external*tr*rate,
            "total": (internal + tr*external)*rate}

def floor_total(fl, params):
    tot = 0.0
    for u in fl["units"]:
        v = unit_val(u["type"], u["rate"], params)["total"]
        tot += 0.0 if pd.isna(v) else v
    return tot

def recalc(df, params):
    df = df.copy()
    for t in df["Type"].unique():
        internal, external = area_for(t, params)
        m = df["Type"] == t
        df.loc[m, "Internal_sqft"] = internal
        df.loc[m, "External_sqft"] = external
        df.loc[m, "Terrace_Rate"]  = terrace_for(t, params)
    # per-unit terrace overrides win over the type default (set by the floor-range tool)
    if "Terrace_Override" in df.columns:
        ov = df["Terrace_Override"].notna()
        df.loc[ov, "Terrace_Rate"] = pd.to_numeric(df.loc[ov, "Terrace_Override"], errors="coerce")
    df["Sellable_sqft"] = df["Internal_sqft"] + df["Terrace_Rate"]*df["External_sqft"]
    df["Total_sqft"]    = df["Internal_sqft"] + df["External_sqft"]
    df["Price"]         = df["Price_sqft"] * df["Sellable_sqft"]
    return df


# ── Session state ──────────────────────────────────────────────────────────────

def _init():
    if "units" not in st.session_state:
        st.session_state.units = load_unit_data()
        st.session_state.uid_counter = len(st.session_state.units)
    if "fm_params" not in st.session_state: st.session_state.fm_params = load_params()
    if "floors"    not in st.session_state: st.session_state.floors    = build_floor_list(st.session_state.units)
    if "blocked"   not in st.session_state: st.session_state.blocked   = load_blocked_floors()

_init()

params  = st.session_state.fm_params
blocked = st.session_state.blocked
df      = recalc(st.session_state.units, params)


# ── Helpers for mutating register + floors ─────────────────────────────────────

def next_uid():
    n = st.session_state.get("uid_counter", 0) + 1
    st.session_state.uid_counter = n
    return f"u{n}"

def gen_unit_nos(floor_num, types_in_order):
    existing = set(st.session_state.units["Unit"].values)
    nos, n = [], 1
    for _ in types_in_order:
        while True:
            cand = str(floor_num*100 + n)
            if cand not in existing and cand not in nos:
                break
            n += 1
        nos.append(cand); n += 1
    return nos

def add_units_to_register(unit_list, floor_num, params):
    """Assigns a fresh uid to each new unit (mutates the dicts) and appends to the register."""
    for u in unit_list:
        uid = next_uid()
        u["uid"] = uid
        d = TYPE_DEFAULTS[u["type"]]
        internal, external = area_for(u["type"], params)
        st.session_state.units = pd.concat([st.session_state.units, pd.DataFrame([{
            "Type": u["type"], "Status": "Available", "Unit": u["unit_no"], "Floor": ordinal(floor_num),
            "Parking": d["parking"], "Internal_sqft": internal, "External_sqft": external,
            "Terrace_Rate": terrace_for(u["type"], params), "Price_sqft": u["rate"],
            "Terrace_Override": pd.NA, "Comment": "", "uid": uid,
        }])], ignore_index=True)

def remove_units_from_register(uids):
    st.session_state.units = st.session_state.units[
        ~st.session_state.units["uid"].isin(uids)].reset_index(drop=True)

def mix_from_editor(edited):
    out = []
    for r in edited.itertuples():
        if pd.notna(r.Type) and pd.notna(r.Qty) and int(r.Qty) > 0:
            out.append((str(r.Type), int(r.Qty)))
    return out

def uid_status_map():
    u = st.session_state.units
    return dict(zip(u["uid"], u["Status"]))

def unit_status(u, smap):
    return smap.get(u.get("uid"), "Available")

def split_floor_units(fl):
    """Split a floor's units into locked (Sold/Bank Locked) and editable (Available), by uid."""
    smap = uid_status_map()
    locked, avail = [], []
    for u in fl["units"]:
        (avail if unit_status(u, smap) == "Available" else locked).append(u)
    return locked, avail

def clear_builder(state_key):
    for k in list(st.session_state.keys()):
        if k.startswith(state_key):
            st.session_state.pop(k, None)

def unit_mix_builder(state_key, default_rows, qmin=1):
    """Rows of (Type dropdown + Qty stepper). Returns list of (type, qty). Uses +/- steppers."""
    sk_rows, sk_ctr = f"{state_key}__rows", f"{state_key}__ctr"
    if sk_rows not in st.session_state:
        st.session_state[sk_ctr] = 0
        st.session_state[sk_rows] = []
        for r in default_rows:
            st.session_state[sk_rows].append({"id": st.session_state[sk_ctr], "type": r["type"], "qty": r["qty"]})
            st.session_state[sk_ctr] += 1
    rows = st.session_state[sk_rows]
    h1, h2, h3 = st.columns([3, 2, 1])
    h1.caption("Topology"); h2.caption("Quantity (use − / +)"); h3.caption(" ")
    to_del = None
    for r in rows:
        rid = r["id"]
        c1, c2, c3 = st.columns([3, 2, 1])
        r["type"] = c1.selectbox("Type", UNIT_TYPES,
                                 index=UNIT_TYPES.index(r["type"]) if r["type"] in UNIT_TYPES else 0,
                                 key=f"{state_key}__t{rid}", label_visibility="collapsed")
        r["qty"] = c2.number_input("Qty", min_value=qmin, max_value=50, value=max(int(r["qty"]), qmin),
                                   step=1, key=f"{state_key}__q{rid}", label_visibility="collapsed")
        if c3.button("🗑️", key=f"{state_key}__x{rid}", help="Remove this row"):
            to_del = rid
    if to_del is not None:
        st.session_state[sk_rows] = [r for r in rows if r["id"] != to_del]
        st.session_state.pop(f"{state_key}__t{to_del}", None)
        st.session_state.pop(f"{state_key}__q{to_del}", None)
        st.rerun()
    if st.button("➕ Add topology", key=f"{state_key}__add"):
        st.session_state[sk_rows].append({"id": st.session_state[sk_ctr], "type": "2 Bedroom", "qty": qmin})
        st.session_state[sk_ctr] += 1
        st.rerun()
    return [(r["type"], int(r["qty"])) for r in st.session_state[sk_rows]]


# ── Sidebar ────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("Muraba Veil")
    st.caption("Unit Manager")
    if st.button("Reload from Excel", use_container_width=True):
        for k in ["units", "fm_params", "floors", "blocked", "uid_counter"]:
            st.session_state.pop(k, None)
        st.rerun()
    st.divider()
    st.caption("Add / edit / remove floors in the **Floor Manager** tab.")
    if blocked:
        st.caption("**Blocked floors (MEP / Majlis):** " + ", ".join(str(k) for k in sorted(blocked)))


# ── Top KPIs (full values) ─────────────────────────────────────────────────────

st.markdown("""<style>
[data-testid="stMetricValue"] { font-size: 1.35rem; }
[data-testid="stMetricLabel"] { font-size: 0.85rem; }
</style>""", unsafe_allow_html=True)

st.title("Muraba Veil — Unit Manager")

_flash = st.session_state.pop("flash", None)
if _flash:
    getattr(st, _flash[0])(_flash[1])

# Global scorecards (full building) — shown on top of every page, no per-tab duplication
ALLOWABLE_SELLABLE = 818186.683338944          # fixed design cap; shown rounded as 818,187
_tot_area = df["Total_sqft"].sum()
_variance = _tot_area - ALLOWABLE_SELLABLE
g1, g2, g3, g4, g5, g6 = st.columns(6)
g1.metric("Units shown", len(df))
g2.metric("Total Area (sqft)", f"{_tot_area:,.0f}")
g3.metric("Total Allowable Sellable (sqft)", f"{ALLOWABLE_SELLABLE:,.0f}")
g4.metric("Variance: Total − Allowable (sqft)", f"{_variance:,.0f}",
          delta=f"{_variance:,.0f}", delta_color="inverse")
g5.metric("Total Price/sqft", aed(df["Price"].sum()/_tot_area) if _tot_area else "—")
g6.metric("Portfolio Value", aed(df["Price"].sum()))

st.divider()

tab1, tab2, tab5, tab3, tab4 = st.tabs(
    ["Unit Register", "Summary by Type", "Topology View", "Floor Manager", "Edit / Remove Units"])


# ── Tab 1: Unit Register ───────────────────────────────────────────────────────

with tab1:
    # Row-level escalation & price variance vs the unit one floor BELOW in the same typology
    order = df.copy()
    order["fnum"] = pd.to_numeric(order["Floor"].str.replace(r"[^0-9]", "", regex=True), errors="coerce")
    order = order.sort_values(["Type", "fnum"])
    esc_map = dict(zip(order["uid"], order.groupby("Type")["Price_sqft"].diff()))
    var_map = dict(zip(order["uid"], order.groupby("Type")["Price"].diff()))

    fc1, fc2 = st.columns(2)
    f_types  = fc1.multiselect("Type",   sorted(df["Type"].unique().tolist()), default=sorted(df["Type"].unique().tolist()))
    f_status = fc2.multiselect("Status", STATUS_OPTIONS, default=STATUS_OPTIONS)
    view = df[df["Type"].isin(f_types) & df["Status"].isin(f_status)].copy()

    # Default sort: by floor (ascending), then unit
    view["_fnum"] = pd.to_numeric(view["Floor"].str.replace(r"[^0-9]", "", regex=True), errors="coerce")
    view = view.sort_values(["_fnum", "Unit"]).drop(columns="_fnum")

    # Derived per-unit columns
    view["PSF_total"]  = view["Price"] / view["Total_sqft"]
    view["Int_Value"]  = view["Price_sqft"] * view["Internal_sqft"]
    view["Terr_Value"] = view["Price_sqft"] * view["Terrace_Rate"] * view["External_sqft"]
    view["Esc_row"]    = view["uid"].map(esc_map)
    view["Var_row"]    = view["uid"].map(var_map)
    view.loc[view["Status"] == "Sold", "Var_row"] = pd.NA   # Floor Wise Variance blank for Sold units

    if "Comment" not in view.columns:
        view["Comment"] = ""

    def _money(v):  return "" if pd.isna(v) else f"AED {v:,.0f}"
    def _num1(v):   return "" if pd.isna(v) else f"{v:,.1f}"

    # Pre-format every value to a string so the look matches; Comment stays editable.
    disp = pd.DataFrame({
        "Type": view["Type"].values, "Status": view["Status"].values,
        "Unit": view["Unit"].values, "Floor": view["Floor"].values,
        "Parking": view["Parking"].astype(int).astype(str).values,
        "Internal (sqft)": view["Internal_sqft"].map(_num1).values,
        "External (sqft)": view["External_sqft"].map(_num1).values,
        "Total Area (sqft)": view["Total_sqft"].map(_num1).values,
        "Sellable (sqft)": view["Sellable_sqft"].map(_num1).values,
        "Terrace Rate": view["Terrace_Rate"].map(lambda v: "" if pd.isna(v) else f"{v:.0%}").values,
        "Price/Sellable sqft": view["Price_sqft"].map(_money).values,
        "Price/Total sqft": view["PSF_total"].map(_money).values,
        "Internal Value (AED)": view["Int_Value"].map(_money).values,
        "Terrace Value (AED)": view["Terr_Value"].map(_money).values,
        "Total Price (AED)": view["Price"].map(_money).values,
        "Escalation vs below (/sqft)": view["Esc_row"].map(_money).values,
        "Floor Wise Variance (AED)": view["Var_row"].map(_money).values,
        "Comment": view["Comment"].fillna("").astype(str).values,
        "uid": view["uid"].values,
    })
    display_cols = [c for c in disp.columns if c != "uid"]
    show_cols = column_picker(display_cols, key="reg_cols", locked=["Type", "Unit"])

    # Styled read-only table with the blue Sold highlight (reliable, no extra deps)
    sold_by_idx = disp.set_index("uid")["Status"] == "Sold"
    vis = disp[show_cols].copy()
    vis.index = disp["uid"].values
    def _hl_sold(row):
        return ["background-color:#9DC3E6" if bool(sold_by_idx.loc[row.name]) else "" for _ in row]
    st.dataframe(vis.style.apply(_hl_sold, axis=1), use_container_width=True,
                 hide_index=True, height=460)
    st.caption(f"Showing {len(view)} of {len(df)} units · Sold units highlighted in blue · "
               f"“vs below” compares each unit to the one a floor lower in the same typology")

    # Inline comment editor (toggle to show/hide); edits persist to file and reload on launch
    show_editor = st.toggle("✏️ Show inline comment editor", value=False, key="show_cmt_editor")
    if show_editor:
        cdf = disp[["Unit", "Type", "Floor", "Status", "Comment", "uid"]].copy()
        ed = st.data_editor(
            cdf, hide_index=True, use_container_width=True, key="cmt_editor",
            column_order=["Unit", "Type", "Floor", "Status", "Comment"],
            disabled=["Unit", "Type", "Floor", "Status"],
            column_config={"Comment": st.column_config.TextColumn("Comment", width="large")},
        )
        new_cmt = ed["Comment"].fillna("").astype(str).values
        if not (new_cmt == cdf["Comment"].values).all():
            cmap = dict(zip(ed["uid"].values, new_cmt))
            u = st.session_state.units
            u["Comment"] = u.apply(lambda r: cmap.get(r["uid"], r.get("Comment", "")), axis=1)
            persist_all_comments()
            st.rerun()


# ── Tab 2: Summary by Type (no Bank Locked column, full values) ────────────────

with tab2:
    st.caption("Mirrors the Excel **Muraba Veil Sale Summary** tab (same columns & structure). "
               "Based on **all units** (Sold included) and computed live from the unit register "
               "using the app's sellable-area pricing.")

    # Column order exactly as in the Excel Sale Summary tab
    SUM_COLS = ["Typology", "Number of Units", "Price/Sq.ft", "Area (sqft)", "Area (sqm)",
                "Internal (sqft/unit)", "Terrace (sqft/unit)", "Total Internal (sqft)",
                "Total Terrace (sqft)", "Total Sellable", "Counted Terraces",
                "Total Sellable Counted", "Price (per unit)", "Total Sales",
                "Parking", "Total Parking"]

    # build one numeric row per typology, ordered like the master type list
    present = [t for t in UNIT_TYPES if t in set(df["Type"])]
    present += [t for t in sorted(df["Type"].unique()) if t not in present]
    num_rows = []
    for t in present:
        g = df[df["Type"] == t]
        n = len(g)
        internal_u = g["Internal_sqft"].mean()                 # constant per type after recalc
        terrace_u  = g["External_sqft"].mean()
        area_sqft  = internal_u + terrace_u                     # full footprint per unit
        tot_internal = g["Internal_sqft"].sum()
        tot_terrace  = g["External_sqft"].sum()                 # 100% terrace footprint
        tot_sellable = g["Total_sqft"].sum()                   # Internal + full Terrace
        tot_counted  = g["Sellable_sqft"].sum()                # rate-adjusted (app) sellable
        counted_terr = tot_counted - tot_internal              # terrace portion actually counted
        total_sales  = g["Price"].sum()
        num_rows.append({
            "Typology": t,
            "Number of Units": n,
            "Price/Sq.ft": total_sales / tot_sellable if tot_sellable else 0.0,
            "Area (sqft)": area_sqft,
            "Area (sqm)": area_sqft / SQFT_PER_SQM,
            "Internal (sqft/unit)": internal_u,
            "Terrace (sqft/unit)": terrace_u,
            "Total Internal (sqft)": tot_internal,
            "Total Terrace (sqft)": tot_terrace,
            "Total Sellable": tot_sellable,
            "Counted Terraces": counted_terr,
            "Total Sellable Counted": tot_counted,
            "Price (per unit)": total_sales / n if n else 0.0,
            "Total Sales": total_sales,
            "Parking": int(g["Parking"].mode().iloc[0]) if not g["Parking"].mode().empty else 0,
            "Total Parking": int(g["Parking"].sum()),
        })
    nm = pd.DataFrame(num_rows)

    # Total row (sum additive; Price/Sq.ft value-weighted; per-unit cells blank like Excel)
    tot_counted_all = nm["Total Sellable Counted"].sum()
    tot_sellable_all = nm["Total Sellable"].sum()
    total = {
        "Typology": "Total",
        "Number of Units": int(nm["Number of Units"].sum()),
        "Price/Sq.ft": (nm["Total Sales"].sum() / tot_sellable_all) if tot_sellable_all else 0.0,
        "Area (sqft)": None, "Area (sqm)": None,
        "Internal (sqft/unit)": None, "Terrace (sqft/unit)": None,
        "Total Internal (sqft)": nm["Total Internal (sqft)"].sum(),
        "Total Terrace (sqft)": nm["Total Terrace (sqft)"].sum(),
        "Total Sellable": nm["Total Sellable"].sum(),
        "Counted Terraces": nm["Counted Terraces"].sum(),
        "Total Sellable Counted": tot_counted_all,
        "Price (per unit)": None,
        "Total Sales": nm["Total Sales"].sum(),
        "Parking": None,
        "Total Parking": int(nm["Total Parking"].sum()),
    }
    nm = pd.concat([nm, pd.DataFrame([total])], ignore_index=True)

    # formatting per column
    def _f(v, kind):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        if kind == "int":   return f"{int(v):,}"
        if kind == "area":  return f"{v:,.2f}"
        if kind == "areak": return f"{v:,.0f}"
        if kind == "aed0":  return f"AED {v:,.0f}"
        return str(v)

    disp = pd.DataFrame({
        "Typology": nm["Typology"],
        "Number of Units": nm["Number of Units"].apply(lambda v: _f(v, "int")),
        "Price/Sq.ft": nm["Price/Sq.ft"].apply(lambda v: _f(v, "aed0")),
        "Area (sqft)": nm["Area (sqft)"].apply(lambda v: _f(v, "area")),
        "Area (sqm)": nm["Area (sqm)"].apply(lambda v: _f(v, "area")),
        "Internal (sqft/unit)": nm["Internal (sqft/unit)"].apply(lambda v: _f(v, "area")),
        "Terrace (sqft/unit)": nm["Terrace (sqft/unit)"].apply(lambda v: _f(v, "area")),
        "Total Internal (sqft)": nm["Total Internal (sqft)"].apply(lambda v: _f(v, "areak")),
        "Total Terrace (sqft)": nm["Total Terrace (sqft)"].apply(lambda v: _f(v, "areak")),
        "Total Sellable": nm["Total Sellable"].apply(lambda v: _f(v, "areak")),
        "Counted Terraces": nm["Counted Terraces"].apply(lambda v: _f(v, "areak")),
        "Total Sellable Counted": nm["Total Sellable Counted"].apply(lambda v: _f(v, "areak")),
        "Price (per unit)": nm["Price (per unit)"].apply(lambda v: _f(v, "aed0")),
        "Total Sales": nm["Total Sales"].apply(lambda v: _f(v, "aed0")),
        "Parking": nm["Parking"].apply(lambda v: _f(v, "int")),
        "Total Parking": nm["Total Parking"].apply(lambda v: _f(v, "int")),
    })[SUM_COLS]

    sum_show = column_picker(list(disp.columns), key="sum_cols", locked=["Typology"])
    excel_table(disp[sum_show])
    st.caption(f"Conversion: 1 m² = {SQFT_PER_SQM} ft²  ·  "
               "Total Sellable = Internal + full Terrace  ·  "
               "Counted Terraces = rate-adjusted terrace  ·  "
               "Total Sellable Counted = Internal + Counted Terraces (drives Price per unit).")

    st.divider()
    chart = nm[nm["Typology"] != "Total"][["Typology", "Total Sales"]].set_index("Typology")
    chart["AED M"] = chart["Total Sales"] / 1e6
    st.bar_chart(chart["AED M"])

    # ── Furniture Pack (static reference, from Excel Sale Summary) ──────────────
    with st.expander("🛋️  Muraba Veil Furniture Pack (reference, from Excel)", expanded=False):
        fp = pd.DataFrame({
            "Type": ["2 Bedroom", "3 Bedroom", "3 Bedroom Pool", "3 Bedroom XL",
                     "3 Bedroom Duplex", "4 Bedroom", "4 Bedroom Duplex", "5 Bedroom PH"],
            "Amount in AED": [475000, 550000, 600000, 800000, 800000, 850000, 850000, 1250000],
        })
        fp["Amount in AED"] = fp["Amount in AED"].apply(lambda x: f"AED {x:,.0f}")
        excel_table(fp)


# ── Tab 5: Topology View (min/max/avg stats) ───────────────────────────────────

with tab5:
    st.subheader("Topology Summary Statistics")
    st.caption("**Total Units** and **Total Value** include all units (Sold + Available). "
               "All other columns (Min / Median / Avg / Max /sqft and Avg Unit Price) are "
               "based on **Available units only** (Sold excluded). "
               "Price/sqft = unit Price ÷ Total Area (Internal + External); "
               "Avg /sqft is value-weighted = Total Value ÷ Total Area.")
    all_types = sorted(df["Type"].unique().tolist())
    pick = st.multiselect("Filter topologies", all_types, default=all_types, key="topo_filter")

    # all-status aggregate (Sold + Available) → drives Total Units & Total Value
    alldf = df.copy()
    if pick:
        alldf = alldf[alldf["Type"].isin(pick)]
    allagg = alldf.groupby("Type").agg(
        Total_Units=("Unit", "count"),
        Total_Value_All=("Price", "sum"),
    ).reset_index()

    # available-only aggregate → drives every pricing stat
    tvdf = df[df["Status"] != "Sold"].copy()
    if pick:
        tvdf = tvdf[tvdf["Type"].isin(pick)]
    tvdf["PSF_total"] = tvdf["Price"] / tvdf["Total_sqft"]

    if tvdf.empty:
        st.info("No available units for the selected topologies.")
    else:
        tv = tvdf.groupby("Type").agg(
            Min_PSF=("PSF_total","min"),
            Median_PSF=("PSF_total","median"),
            Max_PSF=("PSF_total","max"),
            Avg_Price=("Price","mean"),
            Total_Area=("Total_sqft","sum"),
            Avail_Value=("Price","sum"),
        ).reset_index()
        tv["Avg_PSF"] = tv["Avail_Value"] / tv["Total_Area"]   # value-weighted (available)
        tv = tv.merge(allagg, on="Type", how="left")           # bring in all-status Units & Value

        tvd = tv[["Type","Total_Units","Min_PSF","Median_PSF","Avg_PSF","Max_PSF",
                  "Avg_Price","Total_Value_All"]].copy()
        for c in ["Min_PSF","Median_PSF","Avg_PSF","Max_PSF"]:
            tvd[c] = tvd[c].apply(lambda x: f"AED {x:,.0f}")
        tvd["Avg_Price"]       = tvd["Avg_Price"].apply(lambda x: aed(x))
        tvd["Total_Value_All"] = tvd["Total_Value_All"].apply(lambda x: aed(x))
        tvd["Total_Units"]     = tvd["Total_Units"].astype(int)
        tvd.columns = ["Type","Total Units (incl. Sold)",
                       "Min /sqft (lowest)","Median /sqft (mid)","Avg /sqft (wtd)","Max /sqft (highest)",
                       "Avg Unit Price (avail)","Total Value (incl. Sold)"]
        topo_show = column_picker(list(tvd.columns), key="topo_cols", locked=["Type"])
        excel_table(tvd[topo_show])

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.caption("Avg price per total sqft by topology (excl. Sold)")
            st.bar_chart(tv.set_index("Type")["Avg_PSF"])
        with c2:
            st.caption("Total value by topology, incl. Sold (AED M)")
            chart2 = tv.set_index("Type")[["Total_Value_All"]].copy()
            chart2["AED M"] = chart2["Total_Value_All"] / 1e6
            st.bar_chart(chart2["AED M"])


# ── Tab 3: Floor Manager ───────────────────────────────────────────────────────

with tab3:
    floors = st.session_state.floors

    # Parameters
    def terrace_group_key(t):
        if t in ("2 Bedroom", "3 Bedroom - New", "3 Bedroom"): return "standard"
        if t == "3 Bedroom Pool":                          return "3 Bedroom Pool"
        if t == "4 Bedroom Pool":                          return "4 Bedroom Pool"
        if t == "4 Bedroom XL":                            return "simplex"
        if t in ("3 Bedroom Duplex", "4 Bedroom Duplex"):  return "duplex"
        return None   # 5 Bedroom Duplex → fixed 100%

    with st.expander("⚙️  Escalation & Terrace Settings (per typology, from launch sheets)", expanded=True):
        st.caption("Pick a typology, then set its **escalation** (price/sqft added per floor as you go up) "
                   "and **terrace %**. These are the building-wide defaults used when pricing new or "
                   "edited floors. The table below shows the current settings for every typology.")
        s_type = st.selectbox("Typology", UNIT_TYPES, key="set_type")

        sc1, sc2, sc3 = st.columns(3)
        cur_esc = float(params["escalation"].get(s_type, 0.0))
        new_e = sc1.number_input("Escalation (AED/sqft per floor)", min_value=0.0, step=1.0,
                                 value=cur_esc, key=f"set_esc_{s_type}")

        gkey = terrace_group_key(s_type)
        if gkey is not None:
            cur_tr = float(params["terrace"][gkey]) * 100
            new_t = sc2.number_input("Terrace %", min_value=0.0, max_value=100.0, step=1.0,
                                     value=cur_tr, key=f"set_tr_{s_type}")
        else:
            sc2.number_input("Terrace %", min_value=0.0, max_value=100.0, step=1.0, value=100.0,
                             key=f"set_tr_{s_type}", disabled=True,
                             help="5 Bedroom Duplex (penthouse) terrace is fixed at 100%.")
            new_t = None

        cur_dpx = float(params.get("duplex_premium", 0.0))
        if "Duplex" in s_type:
            new_dpx = sc3.number_input("Duplex Premium (AED/sqft)", min_value=0.0, step=50.0,
                                       value=cur_dpx, key="set_dpx")
        else:
            sc3.number_input("Duplex Premium (AED/sqft)", min_value=0.0, step=50.0, value=cur_dpx,
                             key="set_dpx_off", disabled=True, help="Applies to Duplex typologies only.")
            new_dpx = cur_dpx

        old = st.session_state.fm_params
        esc_changed = abs(float(old["escalation"].get(s_type, 0.0)) - new_e) > 1e-9
        dpx_changed = abs(float(old.get("duplex_premium", 0.0)) - new_dpx) > 1e-9

        new_esc_map = dict(params["escalation"]); new_esc_map[s_type] = new_e
        new_tr_map  = dict(params["terrace"])
        if gkey is not None and new_t is not None:
            new_tr_map[gkey] = new_t / 100.0
        np_ = {"escalation": new_esc_map, "terrace": new_tr_map,
               "duplex_premium": new_dpx, "area": dict(params.get("area", {}))}
        if np_ != old:
            st.session_state.fm_params = np_
            n = 0
            if esc_changed:
                n += reladder_typology(s_type, np_)                 # re-ladder this typology live
            if dpx_changed:
                for dt in ("3 Bedroom Duplex", "4 Bedroom Duplex", "5 Bedroom Duplex"):
                    n += reladder_typology(dt, np_)
            if esc_changed or dpx_changed:
                sync_floor_rates()
                st.session_state["flash"] = ("success",
                    f"✅ Escalation updated — {n} Available unit price(s) re-laddered.")
            st.rerun()

        st.caption("Changing **Escalation** re-prices the Available units of that typology up the "
                   "ladder immediately (Sold units stay fixed; the entry/anchor floor holds). "
                   "**Terrace %** and **Area** reprice every unit of the type. The table shows the "
                   "current settings for all typologies.")
        ref_rows = []
        for t in UNIT_TYPES:
            gk = terrace_group_key(t)
            trv = params["terrace"][gk] * 100 if gk else 100.0
            ref_rows.append({"Typology": t,
                             "Escalation (AED/sqft)": f"{params['escalation'].get(t, 0):,.0f}",
                             "Terrace %": f"{trv:.0f}%"})
        excel_table(pd.DataFrame(ref_rows))

    with st.expander("📐  Area Settings (Internal & External sqft per topology — cascades to all units of that type)", expanded=False):
        st.caption("Change a topology's area and every unit of that type updates — sellable area, price and all stats recompute.")
        cur_area = params.get("area", {})
        new_area = {}
        ah1, ah2, ah3 = st.columns([2, 1.5, 1.5])
        ah1.markdown("**Topology**"); ah2.markdown("**Internal (sqft)**"); ah3.markdown("**External (sqft)**")
        for t in UNIT_TYPES:
            a = cur_area.get(t, {"internal": TYPE_DEFAULTS[t]["internal"], "external": TYPE_DEFAULTS[t]["external"]})
            r1, r2, r3 = st.columns([2, 1.5, 1.5])
            r1.markdown(f"<div style='padding-top:6px'>{t}</div>", unsafe_allow_html=True)
            ni = r2.number_input("int", min_value=0.0, step=10.0, value=float(a["internal"]),
                                 key=f"area_int_{t}", label_visibility="collapsed")
            ne = r3.number_input("ext", min_value=0.0, step=10.0, value=float(a["external"]),
                                 key=f"area_ext_{t}", label_visibility="collapsed")
            new_area[t] = {"internal": ni, "external": ne}
        if new_area != params.get("area", {}):
            st.session_state.fm_params = {**st.session_state.fm_params, "area": new_area}
            st.rerun()

    with st.expander("📈  Bulk Update by Typology & Floor Range (escalation + terrace %)", expanded=False):
        st.caption("Pick a typology and a floor range (From → To, or All floors), then optionally "
                   "**add escalation** (AED/sqft) and/or **set the terrace %**. Both apply only to "
                   "**Available** units of that typology in range — Sold units are never changed.")
        u_all = st.session_state.units
        u_all_fn = pd.to_numeric(u_all["Floor"].astype(str).str.replace(r"[^0-9]", "", regex=True),
                                 errors="coerce")

        bt1, bt2 = st.columns([2, 1.4])
        be_type  = bt1.selectbox("Typology", UNIT_TYPES, key="be_type")
        be_scope = bt2.selectbox("Floors", ["All floors", "Floor range", "Single floor"], key="be_scope")

        # floors that actually have an Available unit of this typology
        elig = u_all[(u_all["Type"] == be_type) & (u_all["Status"] == "Available")].copy()
        elig_fn = sorted(set(pd.to_numeric(
            elig["Floor"].astype(str).str.replace(r"[^0-9]", "", regex=True),
            errors="coerce").dropna().astype(int)))

        if not elig_fn:
            st.info(f"No Available {be_type} units to update.")
        else:
            if be_scope == "All floors":
                f_from, f_to = elig_fn[0], elig_fn[-1]
                st.caption(f"Range: **{ordinal(f_from)} → {ordinal(f_to)}** (all {len(elig_fn)} eligible floor(s)).")
            elif be_scope == "Single floor":
                f_one = st.selectbox("Floor", elig_fn, format_func=ordinal, key="be_one")
                f_from = f_to = f_one
            else:  # Floor range
                fc1, fc2 = st.columns(2)
                f_from = fc1.selectbox("From floor", elig_fn, index=0,
                                       format_func=ordinal, key="be_from")
                to_opts = [f for f in elig_fn if f >= f_from]
                f_to = fc2.selectbox("To floor", to_opts, index=len(to_opts) - 1,
                                     format_func=ordinal, key="be_to")

            ac1, ac2, ac3 = st.columns([1.2, 1, 1.2])
            do_esc = ac1.checkbox("Add escalation", value=True, key="be_do_esc")
            amount = ac2.number_input("AED/sqft", value=100.0, step=10.0,
                                      key="be_amount", disabled=not do_esc)
            cur_tr = terrace_for(be_type, params) * 100
            tr_opts = [0, 30, 55, 65, 75, 100]
            if round(cur_tr) not in tr_opts:
                tr_opts = sorted(set(tr_opts + [int(round(cur_tr))]))
            do_tr = ac3.checkbox("Set terrace %", value=False, key="be_do_tr")
            tr_pct = ac3.selectbox("Terrace %", tr_opts,
                                   index=tr_opts.index(int(round(cur_tr))) if int(round(cur_tr)) in tr_opts else 0,
                                   key="be_trpct", disabled=not do_tr)

            mask = ((u_all["Type"] == be_type) & (u_all["Status"] == "Available") &
                    (u_all_fn >= f_from) & (u_all_fn <= f_to))
            n_hit = int(mask.sum())
            bits = []
            if do_esc: bits.append(f"escalate **{amount:+,.0f} AED/sqft**")
            if do_tr:  bits.append(f"set terrace **{tr_pct}%**")
            action_txt = " and ".join(bits) if bits else "make no change (tick an action)"
            st.caption(f"Will update **{n_hit}** Available {be_type} unit(s) on floors "
                       f"{ordinal(f_from)}–{ordinal(f_to)}: {action_txt}.")

            can_apply = n_hit > 0 and (do_esc or do_tr)
            if st.button("Apply", type="primary", disabled=not can_apply, key="be_apply"):
                if do_esc:
                    st.session_state.units.loc[mask, "Price_sqft"] = \
                        st.session_state.units.loc[mask, "Price_sqft"] + amount
                if do_tr:
                    st.session_state.units.loc[mask, "Terrace_Override"] = tr_pct / 100.0
                # keep floor-object rates in sync for the floor table / export
                for fl in st.session_state.floors:
                    if f_from <= fl["floor"] <= f_to:
                        for un in fl["units"]:
                            uid = un.get("uid")
                            if uid is not None:
                                r = st.session_state.units[st.session_state.units["uid"] == uid]
                                if not r.empty and r.iloc[0]["Type"] == be_type and r.iloc[0]["Status"] == "Available":
                                    un["rate"] = float(r.iloc[0]["Price_sqft"])
                st.session_state["flash"] = ("success",
                    f"✅ Updated {n_hit} Available {be_type} unit(s) on floors "
                    f"{ordinal(f_from)}–{ordinal(f_to)}: {action_txt.replace('**','')}.")
                st.rerun()

    # Floors table + totals
    smap_all = uid_status_map()
    rows, grand = [], 0
    TYPE_ABBR = {"2 Bedroom":"2BR","3 Bedroom - New":"3BR New","3 Bedroom":"3BR","3 Bedroom Pool":"3BR Pool","4 Bedroom Pool":"4BR Pool",
                 "4 Bedroom XL":"4BR XL","3 Bedroom Duplex":"3BR DX","4 Bedroom Duplex":"4BR DX","5 Bedroom Duplex":"5BR DX"}
    for fl in floors:
        ft = floor_total(fl, params); grand += ft
        mix = ", ".join(f"{sum(1 for u in fl['units'] if u['type']==t)}x {t}"
                        for t in dict.fromkeys(u["type"] for u in fl["units"]))
        unit_list = ", ".join(f"{u['unit_no']} ({TYPE_ABBR.get(u['type'], u['type'])})" for u in fl["units"])
        n_avail = sum(1 for u in fl["units"] if unit_status(u, smap_all) == "Available")
        editable = "🔒 Locked" if n_avail == 0 else f"{n_avail} avail"
        rows.append({"Floor": ordinal(fl["floor"]), "Kind": fl["kind"], "Levels": fl["levels"],
                     "Unit Mix": mix, "Unit Numbers": unit_list, "Units": len(fl["units"]),
                     "Editable": editable, "Floor Total (AED)": aed(ft)})
    col_t, col_k = st.columns([3, 1])
    with col_t:
        st.subheader("Floors")
        floors_tbl = pd.DataFrame(rows)
        fcols_show = column_picker(list(floors_tbl.columns), key="floor_cols", locked=["Floor"])
        excel_table(floors_tbl[fcols_show])
    with col_k:
        st.subheader("Totals")
        st.metric("Floors", len(floors))
        st.metric("Units",  sum(len(fl["units"]) for fl in floors))
        # Portfolio value is shown in the global header on every page (no duplication here)

    st.divider()
    action = st.radio("Action", ["Add a New Floor", "Edit a Floor"], horizontal=True, key="fm_action")
    st.divider()

    # ─────────────────────── ADD A NEW FLOOR ──────────────────────────────────
    if action == "Add a New Floor":
        st.subheader("Add a New Floor")
        st.caption("Enter a floor number and choose the unit mix. Each unit is priced from the "
                   "**highest-priced reference** at/above the last available unit of that type, "
                   "stepped up by escalation × number of floors to this one.")
        existing = [fl["floor"] for fl in floors]
        nf = st.number_input("Floor number", min_value=1, max_value=999,
                             value=(max(existing)+1 if existing else 59), step=1, key="newfl")

        blocked_hit = nf in blocked
        exists_hit  = nf in existing
        if blocked_hit:
            st.error(f"Floor {ordinal(nf)} is a **{blocked[nf]}** floor (MEP/Majlis) — cannot place residential units here.")
        elif exists_hit:
            st.error(f"Floor {ordinal(nf)} already exists. Use **Edit a Floor** to change it.")

        st.markdown("**Unit mix** — pick topology and use **− / +** to set quantity:")
        mix = unit_mix_builder("addmix", [{"type": "3 Bedroom", "qty": 1}, {"type": "2 Bedroom", "qty": 2}])

        if mix:
            preview, total = [], 0
            for t, q in mix:
                rate = new_unit_rate(t, nf, st.session_state.units, params)
                ref, direction = escalation_reference(t, nf, st.session_state.units)
                uv   = unit_val(t, rate, params)["total"]
                total += uv*q
                if ref:
                    ref_txt = f"Unit {ref['ref_unit']} (Flr {ordinal(ref['ref_floor'])})"
                    base_psf = f"AED {ref['ref_psf']:,.0f}"
                    sign = "+" if direction >= 0 else "−"
                    step_txt = f"{sign}{escalation_for(t, params):,.0f} × {ref['steps']}"
                else:
                    ref_txt, base_psf, step_txt = "—", "—", f"+{escalation_for(t, params):,.0f} × 1"
                preview.append({"Type": t, "Qty": q,
                                "Reference": ref_txt,
                                "Ref Rate/sqft": base_psf,
                                "Escalation (esc × steps)": step_txt,
                                "Rate/sqft": f"AED {rate:,.0f}",
                                "Value each": aed(uv), "Subtotal": aed(uv*q)})
            excel_table(pd.DataFrame(preview))
            m1, m2, m3 = st.columns(3)
            m1.metric("Units on floor", sum(q for _, q in mix))
            m2.metric("Floor Total", aed(total))
            m3.metric("New Portfolio", aed(grand + total), delta=f"+{aed(total)}")

            disabled = blocked_hit or exists_hit
            if st.button(f"Add Floor {ordinal(nf)}", type="primary", disabled=disabled, key="btn_addfl"):
                try:
                    ordered = []
                    for t, q in mix:
                        ordered += [t]*q
                    ordered.sort(key=lambda t: (t != "3 Bedroom", t))
                    nos = gen_unit_nos(nf, ordered)
                    new_units = [{"unit_no": no, "type": t,
                                  "rate": new_unit_rate(t, nf, st.session_state.units, params)}
                                 for no, t in zip(nos, ordered)]
                    add_units_to_register(new_units, nf, params)
                    st.session_state.floors.append({"floor": nf, "kind": "Added",
                                                    "levels": max(TYPE_DEFAULTS[t]["levels"] for t in ordered),
                                                    "units": new_units})
                    st.session_state.floors.sort(key=lambda x: x["floor"])
                    clear_builder("addmix")
                    st.session_state["flash"] = ("success", f"✅ Floor {ordinal(nf)} added with {len(new_units)} unit(s).")
                except Exception as e:
                    st.session_state["flash"] = ("error", f"❌ Could not add floor {nf}: {e}")
                st.rerun()

    # ─────────────────────── EDIT A FLOOR ─────────────────────────────────────
    else:
        st.subheader("Edit a Floor")
        st.caption("Pick a floor, then change the **Available** units — add, remove or swap types. "
                   "Sold units are protected and cannot be changed.")
        _fl_labels = {
            fl["floor"]: f"Floor {ordinal(fl['floor'])}  —  " + ", ".join(
                f"{u['unit_no']} ({TYPE_ABBR.get(u['type'], u['type'])})"
                for u in fl["units"]
            )
            for fl in floors
        }
        sel = st.selectbox(
            "Select floor", ["— select —"] + [fl["floor"] for fl in floors],
            format_func=lambda x: _fl_labels.get(x, f"Floor {ordinal(x)}") if x != "— select —" else x,
            key="edit_floor_sel",
        )

        if sel != "— select —":
            fl = next(f for f in floors if f["floor"] == sel)
            locked_units, avail_units = split_floor_units(fl)
            smap = uid_status_map()

            st.markdown(f"**Floor {ordinal(sel)}** ({fl['kind']}, {fl['levels']} level"
                        f"{'s' if fl['levels']>1 else ''}) — current units:")
            cur = [{"Unit": u["unit_no"], "Type": u["type"],
                    "Status": unit_status(u, smap),
                    "Editable": "Yes" if unit_status(u, smap) == "Available" else "🔒 No",
                    "Rate/sqft": f"AED {u['rate']:,.0f}",
                    "Value": aed(unit_val(u["type"], u["rate"], params)["total"])} for u in fl["units"]]
            excel_table(pd.DataFrame(cur))

            # Block fully-sold / fully-locked floors
            if not avail_units:
                st.error(f"🔒 Floor {ordinal(sel)} has no Available units — every unit is Sold. "
                         f"This floor cannot be edited.")
            else:
                if locked_units:
                    st.info(f"{len(locked_units)} unit(s) on this floor are Sold and will be kept "
                            f"unchanged. You are editing the **{len(avail_units)} Available** unit(s) only.")

                # editable composition = available units only
                comp = {}
                for u in avail_units:
                    comp[u["type"]] = comp.get(u["type"], 0) + 1
                default_rows = [{"type": t, "qty": q} for t, q in comp.items()] or [{"type": "2 Bedroom", "qty": 1}]
                st.markdown("**Editable (Available) unit mix — use − / + to set quantity:**")
                mix = unit_mix_builder(f"editmix_{sel}", default_rows, qmin=0)

                # preview: locked value (fixed) + new available value
                locked_total = sum(unit_val(u["type"], u["rate"], params)["total"] for u in locked_units)
                avail_total = 0
                for t, q in mix:
                    same = [u for u in avail_units if u["type"] == t]
                    for j in range(q):
                        if j < len(same):
                            avail_total += unit_val(t, same[j]["rate"], params)["total"]
                        else:
                            avail_total += unit_val(t, new_unit_rate(t, sel, st.session_state.units, params), params)["total"]
                new_total  = locked_total + avail_total
                old_total  = floor_total(fl, params)
                grand_excl = grand - old_total

                m1, m2, m3 = st.columns(3)
                m1.metric("Units after edit", len(locked_units) + sum(q for _, q in mix))
                m2.metric("Floor value", aed(new_total), delta=aed(new_total - old_total))
                m3.metric("New Portfolio", aed(grand_excl + new_total))

                b1, b2 = st.columns(2)
                if b1.button("Apply Changes", type="primary", key="btn_edit_apply"):
                    try:
                        # keep ALL locked units; rebuild available portion from the mix
                        keep_avail, new_meta = [], []
                        for t, q in mix:
                            same = [u for u in avail_units if u["type"] == t]
                            for j in range(q):
                                if j < len(same):
                                    keep_avail.append(same[j])
                                else:
                                    new_meta.append(t)
                        kept_uids = {u["uid"] for u in keep_avail}
                        removed_uids = [u["uid"] for u in avail_units if u["uid"] not in kept_uids]
                        if removed_uids:
                            remove_units_from_register(removed_uids)
                        added_units = []
                        if new_meta:
                            new_meta.sort(key=lambda t: (t != "3 Bedroom", t))
                            nos = gen_unit_nos(sel, new_meta)
                            added_units = [{"unit_no": no, "type": t,
                                            "rate": new_unit_rate(t, sel, st.session_state.units, params)}
                                           for no, t in zip(nos, new_meta)]
                            add_units_to_register(added_units, sel, params)
                        final_units = locked_units + keep_avail + added_units
                        if final_units:
                            for f2 in st.session_state.floors:
                                if f2["floor"] == sel:
                                    f2["units"] = final_units
                                    if f2["kind"] not in ("Pool",):
                                        f2["kind"] = "Edited"
                                    break
                        else:
                            st.session_state.floors = [f2 for f2 in st.session_state.floors if f2["floor"] != sel]
                        n_add, n_rem = len(added_units), len(removed_uids)
                        clear_builder(f"editmix_{sel}")
                        st.session_state["flash"] = ("success",
                            f"✅ Floor {ordinal(sel)} updated — {n_add} added, {n_rem} removed, "
                            f"{len(locked_units)} protected unit(s) kept.")
                    except Exception as e:
                        st.session_state["flash"] = ("error", f"❌ Could not update floor {sel}: {e}")
                    st.rerun()

                # Remove entire floor only allowed when nothing is locked
                if locked_units:
                    b2.button("Remove Entire Floor", key="btn_edit_remove", disabled=True,
                              help="Cannot remove a floor that has Sold units.")
                elif b2.button("Remove Entire Floor", key="btn_edit_remove"):
                    remove_units_from_register([u["uid"] for u in fl["units"]])
                    st.session_state.floors = [f2 for f2 in st.session_state.floors if f2["floor"] != sel]
                    clear_builder(f"editmix_{sel}")
                    st.session_state["flash"] = ("success", f"✅ Floor {ordinal(sel)} removed entirely.")
                    st.rerun()


# ── Tab 4: Edit / Remove individual units ──────────────────────────────────────

with tab4:
    def uid_label(uid, frame):
        r = frame[frame["uid"] == uid]
        if r.empty:
            return uid
        r = r.iloc[0]
        return f"Unit {r['Unit']} - {r['Type']} (Floor {r['Floor']}) - {r['Status']}"

    st.subheader("Edit a Unit")
    sel_uid = st.selectbox("Select unit", ["— select —"] + df["uid"].tolist(),
                           format_func=lambda x: uid_label(x, df) if x != "— select —" else x, key="edit_sel")
    if sel_uid != "— select —":
        idx = st.session_state.units[st.session_state.units["uid"] == sel_uid].index[0]
        row = st.session_state.units.loc[idx]
        e1, e2 = st.columns(2)
        ns = e1.selectbox("Status", STATUS_OPTIONS,
                          index=STATUS_OPTIONS.index(row["Status"]) if row["Status"] in STATUS_OPTIONS else 0)
        npx = e2.number_input("Price/sqft (AED)", min_value=0.0, value=float(row["Price_sqft"]), step=50.0)
        if st.button("Save Changes", type="primary", key="btn_save"):
            st.session_state.units.at[idx, "Status"] = ns
            st.session_state.units.at[idx, "Price_sqft"] = npx
            st.success(f"Unit {row['Unit']} updated.")
            st.rerun()

    st.divider()
    st.subheader("Remove Units")
    rt = st.multiselect("Filter by type",   sorted(df["Type"].unique().tolist()), key="rem_type")
    rs = st.multiselect("Filter by status", STATUS_OPTIONS, key="rem_status")
    rdf = df.copy()
    if rt: rdf = rdf[rdf["Type"].isin(rt)]
    if rs: rdf = rdf[rdf["Status"].isin(rs)]
    torem = st.multiselect("Select unit(s) to remove", rdf["uid"].tolist(),
                           format_func=lambda u: uid_label(u, rdf), key="rem_ms")
    if torem:
        st.warning(f"This will remove {len(torem)} unit(s). Use Reload to restore.")
        if st.button("Remove Selected Units", type="primary", key="btn_rem"):
            remove_units_from_register(torem)
            st.success(f"Removed {len(torem)} unit(s).")
            st.rerun()


# ── Export ─────────────────────────────────────────────────────────────────────

st.divider()

def build_export():
    d = recalc(st.session_state.units, st.session_state.fm_params)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        eu = d[["Type","Status","Unit","Floor","Parking","Internal_sqft","External_sqft",
                "Total_sqft","Sellable_sqft","Terrace_Rate","Price_sqft","Price"]].copy()
        eu.columns = ["Type","Status","Unit","Floor","Parking","Internal (sqft)","External (sqft)",
                      "Total (sqft)","Sellable (sqft)","Terrace Rate","Price/sqft (AED)","Price (AED)"]
        eu.to_excel(writer, index=False, sheet_name="Unit Register")
        grp = d.groupby("Type").agg(
            Units=("Unit","count"), Sold=("Status", lambda x: (x=="Sold").sum()),
            Available=("Status", lambda x: (x=="Available").sum()),
            Avg_Price_sqft=("Price_sqft","mean"), Min_Price_sqft=("Price_sqft","min"),
            Max_Price_sqft=("Price_sqft","max"), Total_Sellable=("Sellable_sqft","sum"),
            Total_Value=("Price","sum")).reset_index()
        grp.to_excel(writer, index=False, sheet_name="Topology Summary")
        fe = []
        for fl in st.session_state.floors:
            ft = floor_total(fl, st.session_state.fm_params)
            for u in fl["units"]:
                v = unit_val(u["type"], u["rate"], st.session_state.fm_params)
                fe.append({"Floor": fl["floor"], "Kind": fl["kind"], "Levels": fl["levels"],
                           "Unit No": u["unit_no"], "Type": u["type"], "Rate (AED/sqft)": round(u["rate"],2),
                           "Unit Total (AED)": round(v["total"],0), "Floor Total (AED)": round(ft,0)})
        if fe:
            pd.DataFrame(fe).to_excel(writer, index=False, sheet_name="Floor Manager")
    return out.getvalue()

st.download_button("Download Updated Excel", data=build_export(),
                   file_name="Muraba_Veil_Updated.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Downloads current state. Does not overwrite the source file.")
