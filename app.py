import io
import os
import json
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
import anthropic

from review import CHECKLIST, TITLE_NAMES, build_prompt, call_claude, write_excel

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Municipal Code Review",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).parent
MTAS_DIR = BASE_DIR / "mtas"
MTAS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚖️ Municipal Code Review")
    st.divider()

    city_name = st.text_input("City / Town Name", placeholder="e.g. Greeneville")

    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        help="Starts with sk-ant-. Set ANTHROPIC_API_KEY in your shell to avoid pasting it each time.",
    )

    st.divider()
    st.caption(
        "Open each title below, paste in the city's code text, "
        "and optionally paste the MTAS model. "
        "Titles with no city code text are skipped automatically."
    )
    st.divider()

    if st.button("Clear city code", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key.startswith("city_"):
                st.session_state[key] = ""
        st.rerun()

# ---------------------------------------------------------------------------
# Helper: load saved MTAS text for a title
# ---------------------------------------------------------------------------
def load_mtas(title_num: int) -> str:
    path = MTAS_DIR / f"title_{title_num}.txt"
    return path.read_text(encoding="utf-8", errors="replace") if path.exists() else ""


def save_mtas(title_num: int, text: str):
    path = MTAS_DIR / f"title_{title_num}.txt"
    path.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------------
# Title input panels
# ---------------------------------------------------------------------------
st.header("Paste Code by Title")
st.caption("Open a title, paste the city's code on the left and the MTAS model on the right. MTAS text is saved automatically for future reviews.")

for t in range(1, 21):
    title_label = f"Title {t} — {TITLE_NAMES[t]}"
    city_key = f"city_{t}"
    mtas_key = f"mtas_{t}"

    # Pre-populate MTAS from saved file (once per session key)
    if mtas_key not in st.session_state:
        st.session_state[mtas_key] = load_mtas(t)

    with st.expander(title_label):
        col_city, col_mtas = st.columns(2)

        with col_city:
            st.caption("**City Code** (required to include this title)")
            st.text_area(
                label=f"city_code_{t}",
                label_visibility="collapsed",
                key=city_key,
                height=280,
                placeholder=f"Paste the city's Title {t} text here...",
            )

        with col_mtas:
            st.caption("**MTAS Model Code** (optional — saved for future reviews)")
            mtas_val = st.text_area(
                label=f"mtas_model_{t}",
                label_visibility="collapsed",
                key=mtas_key,
                height=280,
                placeholder=f"Paste the MTAS Title {t} model text here...",
            )
            if mtas_val:
                save_mtas(t, mtas_val)

# ---------------------------------------------------------------------------
# Run button
# ---------------------------------------------------------------------------
st.divider()

available_titles = [t for t in range(1, 21) if st.session_state.get(f"city_{t}", "").strip()]
selected_titles = st.multiselect(
    "Titles to run",
    options=available_titles,
    default=available_titles,
    format_func=lambda t: f"Title {t} — {TITLE_NAMES[t]}",
    placeholder="Paste city code into titles above to enable them here",
)

run_col, _ = st.columns([1, 3])
run_clicked = run_col.button("▶  Run Review", type="primary", use_container_width=True)

if run_clicked:
    if not city_name.strip():
        st.error("Enter a city name before running.")
        st.stop()
    if not api_key.strip():
        st.error("Enter your Anthropic API key.")
        st.stop()

    titles_to_run = [t for t in selected_titles if st.session_state.get(f"city_{t}", "").strip()]
    if not titles_to_run:
        st.error("Paste city code text into at least one title and select it before running.")
        st.stop()

    client = anthropic.Anthropic(api_key=api_key.strip())
    all_findings = []

    status_area = st.empty()
    progress_bar = st.progress(0)
    log_area = st.empty()
    log_lines = []

    for i, t in enumerate(titles_to_run):
        city_text = st.session_state[f"city_{t}"].strip()
        mtas_text = st.session_state.get(f"mtas_{t}", "").strip() or None
        title_name = TITLE_NAMES[t]

        status_area.info(f"Reviewing Title {t} — {title_name}…")
        progress_bar.progress((i) / len(titles_to_run))

        try:
            prompt = build_prompt(t, city_name.strip(), city_text, mtas_text)
            findings = call_claude(client, prompt)
            all_findings.extend(findings)
            log_lines.append(f"✓ Title {t} — {title_name}: {len(findings)} finding(s)")
        except Exception as e:
            log_lines.append(f"✗ Title {t} — {title_name}: ERROR — {e}")

        log_area.text("\n".join(log_lines))

    progress_bar.progress(1.0)
    status_area.success(f"Review complete — {len(all_findings)} total findings across {len(titles_to_run)} title(s).")

    # Write Excel to memory
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    write_excel(all_findings, city_name.strip(), tmp_path)
    excel_bytes = tmp_path.read_bytes()
    tmp_path.unlink(missing_ok=True)

    st.session_state["last_findings"] = all_findings
    st.session_state["last_excel"] = excel_bytes
    st.session_state["last_city"] = city_name.strip()

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if st.session_state.get("last_findings"):
    findings = st.session_state["last_findings"]
    city = st.session_state["last_city"]
    excel_bytes = st.session_state["last_excel"]

    st.divider()
    st.subheader(f"Results — {city}")

    # Summary metrics
    checklist_hits = sum(1 for f in findings if f.get("source") == "checklist")
    new_findings = sum(1 for f in findings if f.get("source") == "new_finding")
    cols = st.columns(3)
    cols[0].metric("Total Findings", len(findings))
    cols[1].metric("Standard Checklist Hits", checklist_hits)
    cols[2].metric("New Findings", new_findings)

    # Download button
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{city.replace(' ', '_')}_review_{date_str}.xlsx"
    st.download_button(
        label="⬇  Download Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()

    # Findings table
    import pandas as pd
    df = pd.DataFrame(findings)[["title", "chapter_title", "section", "section_title", "section_text", "comment", "source"]]
    df.columns = ["Title", "Chapter", "Section", "Section Title", "Section Text", "Comment", "Source"]
    df["Title"] = df["Title"].astype(str)

    # Color new findings
    def highlight_new(row):
        if row["Source"] == "new_finding":
            return ["background-color: #fff2cc"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df.style.apply(highlight_new, axis=1),
        use_container_width=True,
        height=500,
        column_config={
            "Section Text": st.column_config.TextColumn(width="large"),
            "Comment": st.column_config.TextColumn(width="large"),
        },
    )
    st.caption("Yellow rows = new findings beyond the standard checklist.")
