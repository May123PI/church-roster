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
    "readings": [],
    "hymns_per_service": 3,
    "readings_per_service": 2,
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

col_a, col_b = st.columns([2, 4])
with col_a:
    picked = st.date_input("Add a service date", value=None, key="date_picker")
    if st.button("Add date"):
        if picked and picked not in st.session_state.service_dates:
            st.session_state.service_dates.append(picked)
            st.session_state.service_dates.sort()
            st.rerun()

with col_b:
    if st.session_state.service_dates:
        st.markdown("**Scheduled services:**")
        to_remove = None
        for i, d in enumerate(st.session_state.service_dates):
            c1, c2 = st.columns([6, 1])
            c1.write(f"**{d.strftime('%A, %d %B %Y')}**")
            if c2.button("✕", key=f"rm_{i}", help="Remove this date"):
                to_remove = i
        if to_remove is not None:
            st.session_state.service_dates.pop(to_remove)
            st.rerun()
    else:
        st.info("No dates added yet. Use the date picker on the left.")

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
st.caption("Enter your master lists below (one item per line). The generator will randomly assign them across services.")

col_h, col_r = st.columns(2)
with col_h:
    hymns_text = "\n".join(cfg["hymns"])
    updated_hymns = st.text_area(
        "Master hymn list",
        value=hymns_text,
        height=250,
        placeholder="The Lord's My Shepherd\nBe Thou My Vision\nHow Great Thou Art\n...",
    )
    cfg["hymns"] = [h.strip() for h in updated_hymns.splitlines() if h.strip()]

with col_r:
    readings_text = "\n".join(cfg["readings"])
    updated_readings = st.text_area(
        "Master scripture readings list",
        value=readings_text,
        height=250,
        placeholder="Genesis 1:1-5\nPsalm 23\nJohn 3:16\n...",
    )
    cfg["readings"] = [r.strip() for r in updated_readings.splitlines() if r.strip()]

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
    if len(cfg["readings"]) < 1:
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
