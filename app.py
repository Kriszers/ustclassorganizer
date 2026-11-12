"""UST Class Organizer: one-click fetch and generate workflow."""
import io
import zipfile

import streamlit as st

from scheduler.data import normalize, read_upload
from scheduler.fetch import TERMS, fetch_courses
from scheduler.scc import make_scc
from scheduler.word_signs import make_word_signs

st.set_page_config(page_title="UST Class Organizer", page_icon="🏛️", layout="wide", initial_sidebar_state="expanded")

# All design changes stay in this file; schedule generation code is unchanged.
st.markdown("""
<style>
/* Subtle St. Thomas-inspired plum-and-gold interface; no university logo used. */
:root { --ust-plum: #592c63; --ust-gold: #e5b65c; }
[data-testid="stAppViewContainer"] {
  background: radial-gradient(ellipse at 96% 0%, rgba(107,62,129,.08), transparent 32%),
              var(--background-color);
}
.block-container { padding-top: 2.0rem; padding-bottom: 3rem; max-width: 1200px; }
[data-testid="stSidebar"] { border-right: 1px solid rgba(124,95,145,.14); }
.ust-hero {
  border-radius: 18px; padding: 31px 36px; margin-bottom: 26px;
  background: linear-gradient(115deg,#382044 0%,#592c63 64%,#724b78 100%);
  color: white; box-shadow: 0 12px 30px rgba(48,23,61,.14);
  border-bottom: 4px solid #e5b65c;
}
.ust-eyebrow { letter-spacing: .16em; font-size: .72rem; text-transform: uppercase;
  color: #f2d59b; font-weight: 700; margin-bottom: 10px; }
.ust-hero h1 { color: #fff !important; font-size: clamp(1.8rem,3.4vw,2.65rem);
  letter-spacing: -.025em; line-height: 1.16; margin: 0 0 11px 0; font-weight: 780; }
.ust-hero p { color: #efe4f0; font-size: .98rem; margin: 0; max-width: 680px; }
.ust-step { color: #8e69a0; letter-spacing: .12em; font-size: .70rem;
  font-weight: 800; text-transform: uppercase; margin: 4px 0 5px; }
.ust-section-heading { font-weight: 770; font-size: 1.45rem; margin: 0 0 10px; }
.ust-instructions { color: var(--text-color); opacity: .75; font-size: .91rem; margin-bottom: 12px; }
.ust-mini-card { border: 1px solid rgba(124,95,145,.23); border-radius: 12px;
  padding: 16px 18px; min-height: 90px; background: rgba(130,105,155,.045); }
.ust-mini-card strong { display:block; font-size: .98rem; margin-bottom: 6px; }
.ust-mini-card span { font-size: .82rem; opacity: .70; }
.ust-semester { border-radius: 12px; border: 1px solid rgba(123,91,146,.24);
  padding: 14px 16px; margin: 14px 0 18px; background: rgba(129,95,152,.055); }
.ust-semester small { opacity:.65; font-size:.72rem; text-transform:uppercase; letter-spacing:.09em; }
.ust-semester b { display:block; font-size:1.18rem; padding-top:3px; }
.ust-foot { font-size:.79rem; opacity:.72; margin-top:14px; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
  background: #592c63 !important; color: #fff !important;
  border: 1px solid #713e7c !important; border-radius: 10px !important;
  min-height: 3.1rem; font-weight: 700;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
  background: #72417e !important; border-color: #e5b65c !important;
}
.stDownloadButton > button { border-radius: 9px; }
[data-testid="stExpander"] { border-radius: 11px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ust-hero">
 <div class="ust-eyebrow">School of Engineering · Scheduling Tools</div>
 <h1>🏛️ UST Class Organizer</h1>
 <p>Classroom schedules, made simple. Fetch course data from ClassFinder and create print-ready Excel and Word signs in one place.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🗓️ Semester settings")
    st.caption("Choose the academic term you want to organize.")
    year = st.number_input("Year", min_value=2020, max_value=2040, value=2027, step=1)
    term_name = st.selectbox("Term", list(TERMS), index=list(TERMS).index("Spring"))
    st.info(f"Selected: {term_name} {int(year)}")
    st.caption("Default: Spring 2027")
    st.divider()
    st.markdown("**How it works**")
    st.caption("① Choose semester  →  ② Generate schedules  →  ③ Download files")


term = TERMS[term_name]
semester = f"{term_name} {year}"
slug = f"{term_name.lower()}_{year}"


def build_schedules(df, semester_label, filename_slug):
    """Return downloadable file contents and an informational summary."""
    df = normalize(df)
    if df.empty:
        raise ValueError("There are no valid course rows to create schedules from.")
    xlsx, scc_count = make_scc(df, semester_label)
    word_docs = make_word_signs(df, semester_label)
    files = {f"SCC_{filename_slug}_sign.xlsx": xlsx}
    for building, (content, _count) in word_docs.items():
        files[f"{building}_{filename_slug}_signs.docx"] = content
    summary = f"{len(df):,} course rows processed; SCC: {scc_count} rooms; " + ", ".join(
        f"{building}: {count} rooms" for building, (_content, count) in word_docs.items()
    )
    missing = [building for building, (_content, count) in word_docs.items() if count == 0]
    if not scc_count:
        missing.insert(0, "SCC")
    return files, summary, missing


# Never show an old semester's downloads when the user changes selection.
selection = (int(year), term)
if st.session_state.get("selection") != selection:
    st.session_state["selection"] = selection
    st.session_state.pop("generated", None)
    st.session_state.pop("courses", None)
    st.session_state.pop("generation_summary", None)
    st.session_state.pop("generation_warnings", None)
    st.session_state.pop("failed_subjects", None)

st.markdown('<div class="ust-step">Step 01 · Create schedules</div>', unsafe_allow_html=True)
st.markdown('<div class="ust-section-heading">Generate classroom schedules</div>', unsafe_allow_html=True)
st.markdown(f'<div class="ust-semester"><small>Current selection</small><b>{semester}</b></div>', unsafe_allow_html=True)
st.markdown('<div class="ust-instructions">One click fetches courses from UST ClassFinder and creates your classroom schedules. Fetching may take several minutes.</div>', unsafe_allow_html=True)

if st.button("🚀 Generate Classroom Schedules", type="primary", use_container_width=True):
    # Clear previous results so failed requests cannot leave misleading downloads.
    for key in ("generated", "courses", "generation_summary", "generation_warnings", "failed_subjects"):
        st.session_state.pop(key, None)
    progress_bar = st.progress(0)
    status = st.empty()

    def update(current, total, subject):
        progress_bar.progress(current / total)
        status.caption(f"Fetching subject {current}/{total}: {subject}")

    try:
        with st.spinner(f"Fetching {semester} courses from UST ClassFinder..."):
            df, failed = fetch_courses(int(year), term, progress=update)
        st.session_state["courses"] = df
        st.session_state["failed_subjects"] = failed
        status.caption("Course download complete. Creating Excel and Word schedules...")
        with st.spinner("Generating SCC Excel and Word classroom signs..."):
            generated, summary, missing = build_schedules(df, semester, slug)
        st.session_state["generated"] = generated
        st.session_state["generation_summary"] = summary
        st.session_state["generation_warnings"] = missing
        progress_bar.progress(1.0)
        status.empty()
    except Exception as exc:
        status.empty()
        st.error(f"Could not complete automatic generation: {exc}")
        st.info("If ClassFinder is unavailable, expand 'Backup: Upload CSV / Excel' below.")

with st.expander("📁 Optional backup · Upload CSV / Excel"):
    st.caption("Only needed if ClassFinder cannot be reached or you want to use a previously saved course file.")
    upload = st.file_uploader("Course data", type=["csv", "xlsx", "xls"])
    if upload is not None and st.button("Generate schedules from uploaded file"):
        try:
            df = read_upload(upload)
            generated, summary, missing = build_schedules(df, semester, slug)
            st.session_state["courses"] = df
            st.session_state["generated"] = generated
            st.session_state["generation_summary"] = summary
            st.session_state["generation_warnings"] = missing
            st.session_state["failed_subjects"] = []
        except Exception as exc:
            st.error(f"Could not process uploaded file: {exc}")

if "generated" in st.session_state:
    st.markdown('<div class="ust-step">Step 02 · Download results</div>', unsafe_allow_html=True)
    st.success("Your classroom schedule files are ready.")
    st.write(st.session_state["generation_summary"])
    failed = st.session_state.get("failed_subjects", [])
    if failed:
        st.warning(f"ClassFinder returned no parsed sections for {len(failed)} subjects: {', '.join(failed)}. Verify completeness before printing.")
    missing = st.session_state.get("generation_warnings", [])
    if missing:
        st.warning("No eligible physical classroom meetings found for: " + ", ".join(missing) + ". Confirm room formats and schedules.")

    outputs = st.session_state["generated"]
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zipf:
        for filename, content in outputs.items():
            zipf.writestr(filename, content)
    st.download_button("⬇ Download all schedules (.zip)", archive.getvalue(),
                       file_name=f"UST_classroom_signs_{slug}.zip", mime="application/zip", type="primary", use_container_width=True)
    st.markdown("**Individual files**")
    for filename, content in outputs.items():
        mime = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if filename.endswith(".xlsx") else
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.download_button(("📊 " if filename.endswith(".xlsx") else "📄 ") + filename, content, file_name=filename, mime=mime, key=f"download_{slug}_{filename}")

if "courses" in st.session_state:
    with st.expander("🔎 Review source course data"):
        courses = normalize(st.session_state["courses"])
        st.dataframe(courses, use_container_width=True, height=260)
        st.download_button("Download course CSV", courses.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"ust_{slug}_courses.csv", mime="text/csv")

st.divider()
st.markdown("""<div class="ust-foot">Data source: UST ClassFinder (external university service). Verify course dates and classroom assignments before printing. This independent tool does not request UST credentials.</div>""", unsafe_allow_html=True)
