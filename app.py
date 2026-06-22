import json
import os
from datetime import date, datetime

import streamlit as st

from generator import generate_excel

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "data", "config.json")

DEFAULT_CONFIG = {
    "church_name": "St Example Parish Church",
    "duties": [
        {"name": "Sides Person / Welcomer", "people": []},
        {"name": "Reader (1st Lesson)", "people": []},
        {"name": "Intercessor", "people": []},
        {"name": "Ad Hoc Duty 1", "people": []},
        {"name": "Ad Hoc Duty 2", "people": []},
        {"name": "Ad Hoc Duty 3", "people": []},
    ],
    "hymns": [],
    "readings": [],   # list of {"ref": str, "text": str}
    "hymns_per_service": 3,
    "readings_per_service": 2,
    "prayers": [],    # list of {"name": str, "text": str}
}


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()


def save_config(cfg: dict) -> None:
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ── page config ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Church Roster Generator",
    page_icon="✝",
    layout="wide",
)

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    h1 { color: #1F3864; }
    h2 { color: #1F3864; border-bottom: 2px solid #C9A94B; padding-bottom: 4px; }
    h3 { color: #1F3864; }
    .stButton > button { background-color: #1F3864; color: white; border-radius: 6px; }
    .stButton > button:hover { background-color: #C9A94B; color: white; }
    .stDownloadButton > button { background-color: #C9A94B; color: white;
                                  border-radius: 6px; font-size: 1.1rem;
                                  padding: 0.5rem 2rem; }
</style>
""", unsafe_allow_html=True)

# ── session state ─────────────────────────────────────────────────────────
if "cfg" not in st.session_state:
    st.session_state.cfg = load_config()
if "service_dates" not in st.session_state:
    st.session_state.service_dates = []

cfg = st.session_state.cfg

# ── header ────────────────────────────────────────────────────────────────
st.title("Church of England — Order of Service & Roster Generator")
st.caption("Fill in the sections below then click **Generate Excel** to download your roster.")

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 1 — Church Settings
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 1. Church Settings")
col1, col2, col3 = st.columns([3, 1, 1])
with col1:
    cfg["church_name"] = st.text_input("Church name", value=cfg["church_name"])
with col2:
    cfg["hymns_per_service"] = st.number_input(
        "Hymns per service", min_value=1, max_value=8,
        value=cfg.get("hymns_per_service", 3))
with col3:
    cfg["readings_per_service"] = st.number_input(
        "Readings per service", min_value=1, max_value=6,
        value=cfg.get("readings_per_service", 2))

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 2 — Service Dates
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 2. Service Dates")
st.caption("Add dates one at a time, or use **Bulk Add** to fill a whole season in one click.")

tab_single, tab_bulk = st.tabs(["Add single date", "Bulk add (by weekday)"])

with tab_single:
    col_a, col_b = st.columns([2, 3])
    with col_a:
        picked = st.date_input("Service date", value=None, key="date_picker")
        if st.button("Add date"):
            if picked and picked not in st.session_state.service_dates:
                st.session_state.service_dates.append(picked)
                st.session_state.service_dates.sort()
                st.rerun()

with tab_bulk:
    st.markdown("Add every **nth weekday** between two dates (e.g. every Sunday for a year).")
    bc1, bc2, bc3 = st.columns([2, 2, 2])
    with bc1:
        bulk_start = st.date_input("From", value=date.today(), key="bulk_start")
    with bc2:
        bulk_end = st.date_input(
            "To (max 12 months ahead)",
            value=date(date.today().year + 1, date.today().month, date.today().day),
            key="bulk_end",
            min_value=date.today(),
            max_value=date(date.today().year + 1, 12, 31),
        )
    with bc3:
        WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday",
                    "Thursday", "Friday", "Saturday"]
        bulk_day = st.selectbox("Weekday", WEEKDAYS, index=0, key="bulk_day")
        bulk_every = st.selectbox("Frequency", ["Every week", "Every 2 weeks", "Every 4 weeks"], key="bulk_every")

    # exclusions
    if "bulk_exclusions" not in st.session_state:
        st.session_state.bulk_exclusions = []

    st.markdown("**Exclude dates** (e.g. bank holidays, special services):")
    ex_col1, ex_col2 = st.columns([3, 1])
    with ex_col1:
        excl_pick = st.date_input("Date to exclude", value=None, key="excl_picker")
    with ex_col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Add exclusion"):
            if excl_pick and excl_pick not in st.session_state.bulk_exclusions:
                st.session_state.bulk_exclusions.append(excl_pick)
                st.session_state.bulk_exclusions.sort()
                st.rerun()

    if st.session_state.bulk_exclusions:
        excl_remove = None
        excl_cols = st.columns(4)
        for ei, ed in enumerate(st.session_state.bulk_exclusions):
            with excl_cols[ei % 4]:
                ec1, ec2 = st.columns([4, 1])
                ec1.markdown(f"~~{ed.strftime('%d %b %Y')}~~")
                if ec2.button("✕", key=f"excl_rm_{ei}", help="Remove exclusion"):
                    excl_remove = ei
        if excl_remove is not None:
            st.session_state.bulk_exclusions.pop(excl_remove)
            st.rerun()
        if st.button("Clear exclusions", key="clear_excl"):
            st.session_state.bulk_exclusions = []
            st.rerun()

    if st.button("Bulk add", type="primary"):
        from datetime import timedelta
        step = {"Every week": 7, "Every 2 weeks": 14, "Every 4 weeks": 28}[bulk_every]
        py_target = 6 if WEEKDAYS.index(bulk_day) == 0 else WEEKDAYS.index(bulk_day) - 1
        cursor = bulk_start
        days_ahead = (py_target - cursor.weekday()) % 7
        cursor = cursor + timedelta(days=days_ahead)
        exclusion_set = set(st.session_state.bulk_exclusions)
        added, skipped = 0, 0
        while cursor <= bulk_end:
            if cursor in exclusion_set:
                skipped += 1
            elif cursor not in st.session_state.service_dates:
                st.session_state.service_dates.append(cursor)
                added += 1
            cursor += timedelta(days=step)
        st.session_state.service_dates.sort()
        msg = f"Added {added} date(s)."
        if skipped:
            msg += f" Skipped {skipped} excluded date(s)."
        st.success(msg)
        st.rerun()

# date list with remove buttons
st.markdown("---")
if st.session_state.service_dates:
    n_dates = len(st.session_state.service_dates)
    st.markdown(f"**{n_dates} service date(s) scheduled:**")
    # show in 3-column grid
    to_remove = None
    grid_cols = st.columns(3)
    for i, d in enumerate(st.session_state.service_dates):
        with grid_cols[i % 3]:
            cc1, cc2 = st.columns([5, 1])
            cc1.write(f"{d.strftime('%a, %d %b %Y')}")
            if cc2.button("✕", key=f"rm_{i}", help="Remove"):
                to_remove = i
    if to_remove is not None:
        st.session_state.service_dates.pop(to_remove)
        st.rerun()
    if st.button("Clear all dates", type="secondary"):
        st.session_state.service_dates = []
        st.rerun()
else:
    st.info("No dates added yet.")

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 3 — Duties & People
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 3. Duties & People")
st.caption("Enter one name per line in each duty pool. Leave the duty name blank to hide it.")

# Standard duties (first 3 locked names, last 3 ad hoc)
STANDARD_DUTIES = 3
while len(cfg["duties"]) < 6:
    cfg["duties"].append({"name": "", "people": []})

cols = st.columns(3)
for i in range(6):
    with cols[i % 3]:
        default_names = [
            "Sides Person / Welcomer",
            "Reader (1st Lesson)",
            "Intercessor",
            "Ad Hoc Duty 1",
            "Ad Hoc Duty 2",
            "Ad Hoc Duty 3",
        ]
        label = "Duty name" if i >= STANDARD_DUTIES else "Duty"
        duty_name = st.text_input(
            label,
            value=cfg["duties"][i]["name"] or default_names[i],
            key=f"duty_name_{i}",
            disabled=(i < STANDARD_DUTIES),
        )
        if i >= STANDARD_DUTIES:
            cfg["duties"][i]["name"] = duty_name

        people_text = "\n".join(cfg["duties"][i]["people"])
        updated = st.text_area(
            f"People for **{cfg['duties'][i]['name'] or f'Duty {i+1}'}**",
            value=people_text,
            height=130,
            key=f"duty_people_{i}",
            placeholder="One name per line",
        )
        cfg["duties"][i]["people"] = [
            p.strip() for p in updated.splitlines() if p.strip()
        ]

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 4 — Hymns & Readings
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 4. Hymns & Scripture Readings")

# ── Hymns ──
st.markdown("### Hymns")
st.caption("One hymn per line. These are randomly assigned across services.")
hymns_text = "\n".join(cfg["hymns"])
updated_hymns = st.text_area(
    "Master hymn list",
    value=hymns_text,
    height=180,
    placeholder="The Lord's My Shepherd\nBe Thou My Vision\nHow Great Thou Art\n...",
)
cfg["hymns"] = [h.strip() for h in updated_hymns.splitlines() if h.strip()]

# ── Readings ──
st.markdown("### Scripture Readings")
st.caption(
    "Add each reading below — a reference (e.g. *Psalm 23*) is required; "
    "the full text is optional but will be printed in full on each service sheet."
)

# migrate legacy plain-string readings to dict format
if cfg["readings"] and isinstance(cfg["readings"][0], str):
    cfg["readings"] = [{"ref": r, "text": ""} for r in cfg["readings"]]

# add / remove readings
if "readings" not in cfg:
    cfg["readings"] = []

ra_col, rb_col = st.columns([5, 1])
with ra_col:
    new_ref = st.text_input("New reading reference", placeholder="e.g. John 14:1-6", key="new_reading_ref")
with rb_col:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Add reading"):
        if new_ref.strip():
            cfg["readings"].append({"ref": new_ref.strip(), "text": ""})
            st.rerun()

readings_to_remove = None
for ri, reading in enumerate(cfg["readings"]):
    with st.expander(f"📖 {reading['ref'] or f'Reading {ri+1}'}", expanded=False):
        rc1, rc2 = st.columns([8, 1])
        with rc1:
            new_r = st.text_input(
                "Reference", value=reading["ref"],
                key=f"reading_ref_{ri}", label_visibility="collapsed"
            )
            cfg["readings"][ri]["ref"] = new_r
        with rc2:
            if st.button("Remove", key=f"rm_reading_{ri}"):
                readings_to_remove = ri
        cfg["readings"][ri]["text"] = st.text_area(
            "Full scripture text (optional — printed word-for-word in the service sheet)",
            value=reading.get("text", ""),
            height=150,
            key=f"reading_text_{ri}",
            placeholder="Paste the full passage text here...",
        )
if readings_to_remove is not None:
    cfg["readings"].pop(readings_to_remove)
    st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 5 — Sheet Music
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 5. Sheet Music")
st.caption(
    "Select an instrument and click a hymn to open free sheet music. "
    "Only public-domain hymns have direct PDF links — others open a search."
)

INSTRUMENT_OPTIONS = ["Organ", "Guitar", "Flute"]
instrument = st.radio("Instrument", INSTRUMENT_OPTIONS, horizontal=True)

# Hymnary.org instrument query keywords
INSTRUMENT_QUERY = {
    "Organ":  "organ",
    "Guitar": "guitar chords",
    "Flute":  "flute",
}

# Known public-domain hymns with direct Hymnary tune page slugs
# Format: lowercase hymn name fragment → hymnary tune slug
HYMNARY_SLUGS = {
    "the lord's my shepherd":        "crimond",
    "be thou my vision":             "slane",
    "how great thou art":            "how_great_thou_art",
    "guide me o thou great redeemer":"cwm_rhondda",
    "immortal invisible":            "st_denio",
    "all things bright and beautiful":"royal_oak",
    "dear lord and father of mankind":"repton",
    "praise my soul the king of heaven":"lauda_anima",
    "to god be the glory":           "to_god_be_the_glory",
    "o lord my god":                 "how_great_thou_art",
    "abide with me":                 "eventide",
    "o come all ye faithful":        "adeste_fideles",
    "hark the herald angels sing":   "mendelssohn",
    "away in a manger":              "away_in_a_manger",
    "love divine all loves excelling":"blaenwern",
}

def _hymnary_url(hymn: str, instrument: str) -> tuple[str, str]:
    """Return (url, label) for a hymn + instrument combo."""
    key = hymn.lower().strip()
    slug = next((v for k, v in HYMNARY_SLUGS.items() if k in key), None)
    instr_q = INSTRUMENT_QUERY[instrument]
    if slug:
        url = f"https://hymnary.org/tune/{slug}"
        label = f"Open on Hymnary.org ({instrument} score)"
    else:
        query = hymn.replace(" ", "+") + f"+{instr_q}+sheet+music"
        url = f"https://hymnary.org/search?qu={query}"
        label = f"Search Hymnary.org for {instrument} music"
    return url, label

hymn_list = cfg.get("hymns", [])
if not hymn_list:
    st.info("Add hymns in Section 4 to see sheet music links here.")
else:
    cols_sm = st.columns(3)
    for idx, hymn in enumerate(hymn_list):
        with cols_sm[idx % 3]:
            url, label = _hymnary_url(hymn, instrument)
            st.markdown(
                f"**{hymn}**  \n"
                f"[{label}]({url})"
            )

# ═══════════════════════════════════════════════════════════════════════════
# SECTION 6 — Prayer Templates
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("## 6. Prayer Templates")
st.caption(
    "Add standing prayers that appear word-for-word on every service sheet "
    "(e.g. Collect, Confession, Lord's Prayer). "
    "Order them here — they print in this order."
)

if "prayers" not in cfg:
    cfg["prayers"] = []

pa_col, pb_col = st.columns([5, 1])
with pa_col:
    new_prayer_name = st.text_input(
        "New prayer name", placeholder="e.g. The Collect, Confession, Lord's Prayer",
        key="new_prayer_name"
    )
with pb_col:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Add prayer"):
        if new_prayer_name.strip():
            cfg["prayers"].append({"name": new_prayer_name.strip(), "text": ""})
            st.rerun()

prayer_to_remove = None
for pi, prayer in enumerate(cfg["prayers"]):
    with st.expander(f"🙏 {prayer['name'] or f'Prayer {pi+1}'}", expanded=False):
        pc1, pc2 = st.columns([8, 1])
        with pc1:
            cfg["prayers"][pi]["name"] = st.text_input(
                "Prayer name", value=prayer["name"],
                key=f"prayer_name_{pi}", label_visibility="collapsed"
            )
        with pc2:
            if st.button("Remove", key=f"rm_prayer_{pi}"):
                prayer_to_remove = pi
        cfg["prayers"][pi]["text"] = st.text_area(
            "Prayer text (word for word)",
            value=prayer.get("text", ""),
            height=180,
            key=f"prayer_text_{pi}",
            placeholder="Type or paste the full prayer text here...",
        )
if prayer_to_remove is not None:
    cfg["prayers"].pop(prayer_to_remove)
    st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# SAVE + GENERATE
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("---")
col_save, col_gen = st.columns([1, 2])

with col_save:
    if st.button("Save configuration"):
        save_config(cfg)
        st.success("Configuration saved.")

with col_gen:
    # validation
    errors = []
    if not st.session_state.service_dates:
        errors.append("Add at least one service date.")
    n_services = len(st.session_state.service_dates)
    hymns_needed = n_services * cfg["hymns_per_service"]
    readings_needed = n_services * cfg["readings_per_service"]
    if len(cfg["hymns"]) < 1:
        errors.append(f"Add at least 1 hymn (need {hymns_needed} total, cycling is fine).")
    valid_readings = [r for r in cfg["readings"] if (r["ref"] if isinstance(r, dict) else r).strip()]
    if len(valid_readings) < 1:
        errors.append(f"Add at least 1 reading (need {readings_needed} total, cycling is fine).")
    active_duties_with_people = [
        d for d in cfg["duties"] if d["name"].strip() and d["people"]
    ]

    if errors:
        for e in errors:
            st.warning(e)
    else:
        if st.button("Generate Excel Roster", type="primary"):
            save_config(cfg)
            try:
                excel_bytes = generate_excel(
                    church_name=cfg["church_name"],
                    service_dates=st.session_state.service_dates,
                    duties=cfg["duties"],
                    hymns=cfg["hymns"],
                    readings=cfg["readings"],
                    hymns_per_service=cfg["hymns_per_service"],
                    readings_per_service=cfg["readings_per_service"],
                    prayers=cfg.get("prayers", []),
                )
                fname = f"church_roster_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="Download Roster Excel",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success(
                    f"Roster generated! {n_services} service sheet(s) + Home summary tab."
                )
            except Exception as ex:
                st.error(f"Generation failed: {ex}")

st.markdown("---")
st.caption("Church Roster Generator — built with Streamlit & openpyxl")
