"""Defect coding app.

Run with:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import os
from datetime import date

import pandas as pd
import streamlit as st

import defect_core as core

st.set_page_config(page_title="Defect Coding", page_icon="🔬", layout="wide")


# --------------------------------------------------------------------------
# Configuration (environment variables or .streamlit/secrets.toml)
# --------------------------------------------------------------------------
def setting(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:  # no secrets file
        pass
    return os.environ.get(name, default)


API_KEY = setting("ANTHROPIC_API_KEY")
MODEL = setting("ANTHROPIC_MODEL", "claude-sonnet-5-5")
EDITOR_PASSWORD = setting("EDITOR_PASSWORD")

FLAG = {"None": "🔴", "Medium": "🟡", "High": "🟢", "": "⚪"}
DECISIONS = ["Auto", "Accepted", "Changed", "Excluded", "Needs review"]

ss = st.session_state
ss.setdefault("draft", None)        # list of row dicts for the month being coded
ss.setdefault("draft_key", None)    # "YYYY-MM"
ss.setdefault("draft_file", "")
ss.setdefault("sheets", {})
ss.setdefault("sheet", "")
ss.setdefault("upload_id", None)
ss.setdefault("ed_ver", 0)          # bump to reset the review table after bulk changes
ss.setdefault("flash", None)


def bump():
    ss.ed_ver += 1


def flash(kind: str, text: str):
    ss.flash = (kind, text)


# --------------------------------------------------------------------------
# Sidebar: who you are and whether you can edit
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("Defect coding")
    user_name = st.text_input("Your name", key="user_name",
                              help="Recorded on the History tab when you save a month.")
    if EDITOR_PASSWORD:
        pw = st.text_input("Editor password", type="password",
                           help="Needed to code and save months or change codes and rules.")
        can_edit = pw == EDITOR_PASSWORD
        st.caption("Can edit" if can_edit else "View only: enter the editor password to code months.")
    else:
        can_edit = True
    st.divider()
    if API_KEY:
        st.caption(f"Auto-coding is on ({MODEL}).")
    else:
        st.warning("Auto-coding is off: no ANTHROPIC_API_KEY is set. "
                   "Rows matching a past description are still coded, and you can upload a file "
                   "that was coded elsewhere.")

months = core.load_history()
settings = core.load_settings()
codes = core.all_codes(months, settings)

st.title("Defect coding")
if months:
    total = sum(len(m["rows"]) for m in months.values())
    keys = sorted(months)
    st.caption(f"{len(months)} months in the history ({core.month_label(keys[0])} to "
               f"{core.month_label(keys[-1])}), {total} defects")
else:
    st.caption("The history is empty. Code and save a month, or import a history workbook on the History tab.")

if ss.flash:
    kind, text = ss.flash
    getattr(st, kind)(text)
    ss.flash = None

tab_code, tab_summary, tab_history, tab_settings = st.tabs(
    ["Code a month", "Monthly summary", "History", "Codes & rules"])


# --------------------------------------------------------------------------
# Code a month
# --------------------------------------------------------------------------
def load_draft_from_sheet(name: str):
    rows = ss.sheets[name]
    ss.sheet = name
    ss.draft = core.new_draft_rows(rows, codes)
    ss.draft_key = core.guess_month(rows)
    core.code_exact_matches(ss.draft, months)
    bump()


with tab_code:
    if not can_edit:
        st.info("Coding a month needs edit access. Enter the editor password in the sidebar. "
                "You can still use the Monthly summary tab.")
    else:
        up = st.file_uploader("This month's defect file",
                              type=["xlsx", "xlsm", "csv"],
                              help="Part number, date, program, and defect description in columns A to D. "
                                   "Leave column E empty, or upload a file that's already coded to review it here.")
        if up is not None:
            uid = hashlib.md5(up.getvalue()).hexdigest()
            if uid != ss.upload_id:
                try:
                    sheets = core.read_upload(up.getvalue(), up.name)
                except Exception:
                    st.error("That file couldn't be read. Save it as .xlsx and try again.")
                    sheets = {}
                sheets = {k: v for k, v in sheets.items() if v}
                if sheets:
                    ss.upload_id, ss.sheets, ss.draft_file = uid, sheets, up.name
                    load_draft_from_sheet(max(sheets, key=lambda k: len(sheets[k])))
                elif up is not None and not ss.draft:
                    st.error("No defect descriptions found. Put descriptions in column D, starting on row 2.")

        draft = ss.draft
        if draft:
            # ---- month and worksheet
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                if len(ss.sheets) > 1:
                    names = list(ss.sheets)
                    pick = st.selectbox("Worksheet", names, index=names.index(ss.sheet),
                                        format_func=lambda n: f"{n} ({len(ss.sheets[n])} rows)")
                    if pick != ss.sheet:
                        load_draft_from_sheet(pick)
                        st.rerun()
                st.subheader(core.month_label(ss.draft_key))
                st.caption(f"{ss.draft_file}, {len(draft)} defects")
            y, m = int(ss.draft_key[:4]), int(ss.draft_key[5:7])
            with c2:
                new_m = st.selectbox("Month", range(1, 13), index=m - 1,
                                     format_func=lambda i: core.MONTHS[i - 1])
            with c3:
                years = list(range(date.today().year - 3, date.today().year + 2))
                if y not in years:
                    years.append(y)
                new_y = st.selectbox("Year", sorted(years), index=sorted(years).index(y))
            ss.draft_key = f"{new_y}-{new_m:02d}"

            # ---- auto-code
            uncoded = sum(1 for r in draft if not r["conf"])
            if uncoded:
                label = f"Auto-code {uncoded} defects" if uncoded == len(draft) else f"Code remaining {uncoded} rows"
                if st.button(label, type="primary", disabled=not API_KEY):
                    import anthropic
                    client = anthropic.Anthropic(api_key=API_KEY)
                    bar = st.progress(0.0, text="Coding… each batch can take up to a minute.")
                    try:
                        core.auto_code(draft, months, settings, client, MODEL,
                                       progress=lambda d, t: bar.progress(d / t, text=f"Coded batch {d} of {t}"))
                    except anthropic.AuthenticationError:
                        flash("error", "The Anthropic API key was rejected. Check ANTHROPIC_API_KEY.")
                    except anthropic.RateLimitError:
                        flash("error", "Claude is busy right now. Wait a minute, then select Code remaining rows.")
                    except Exception as e:  # noqa: BLE001
                        flash("error", f"Auto-coding stopped ({type(e).__name__}). Rows coded so far are kept; "
                                       "select Code remaining rows to continue.")
                    bump()
                    st.rerun()
                if not API_KEY:
                    st.caption("Auto-coding needs an API key. Download the rows below, code them in the "
                               "Defect Analysis project, then upload the coded file here to review and save.")

            # ---- counts and bulk actions
            red = [r for r in draft if r["conf"] == "None"]
            yellow = [r for r in draft if r["conf"] == "Medium"]
            pending = sum(1 for r in draft if r["decision"] == "Needs review")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("🔴 No matching code", len(red))
            k2.metric("🟡 To review", len(yellow))
            k3.metric("🟢 Coded", sum(1 for r in draft if r["conf"] == "High"))
            k4.metric("Awaiting your decision", pending)

            b1, b2, _ = st.columns([1, 1, 3])
            if b1.button("Accept all yellow", disabled=not any(r["decision"] == "Needs review" and r["code"] for r in yellow)):
                for r in yellow:
                    if r["decision"] == "Needs review" and r["code"]:
                        r["decision"] = "Accepted"
                bump(); st.rerun()
            if b2.button("Exclude all red", disabled=not any(r["decision"] == "Needs review" for r in red)):
                for r in red:
                    if r["decision"] == "Needs review":
                        r["decision"] = "Excluded"
                bump(); st.rerun()

            # ---- review table
            show = st.radio("Rows shown", ["Needs review", "All rows"], horizontal=True,
                            index=0 if (red or yellow) else 1, on_change=bump, key="show_filter")
            idx = [i for i, r in enumerate(draft) if show == "All rows" or core.is_flagged(r) or not r["conf"]]
            if not idx:
                st.success("Nothing needs review. Every row matched a code with high confidence.")
            else:
                df = pd.DataFrame([{
                    "Flag": FLAG[draft[i]["conf"]], "Row": i + 2, "Part": draft[i]["pn"], "Date": draft[i]["date"],
                    "Program": draft[i]["prog"], "Defect description": draft[i]["desc"],
                    "Defect code": draft[i]["code"] or None, "Confidence": draft[i]["conf"],
                    "Note": draft[i]["note"], "Decision": draft[i]["decision"] or None,
                } for i in idx])
                edited = st.data_editor(
                    df, key=f"editor_{ss.ed_ver}", hide_index=True, use_container_width=True,
                    disabled=["Flag", "Row", "Part", "Date", "Program", "Defect description", "Confidence", "Note"],
                    column_config={
                        "Flag": st.column_config.TextColumn(" ", width="small"),
                        "Row": st.column_config.NumberColumn(width="small"),
                        "Defect description": st.column_config.TextColumn(width="large"),
                        "Defect code": st.column_config.SelectboxColumn(options=codes, width="medium"),
                        "Decision": st.column_config.SelectboxColumn(options=DECISIONS, width="small"),
                        "Note": st.column_config.TextColumn(width="medium"),
                    })
                changed = False
                for _, e in edited.iterrows():
                    r = draft[int(e["Row"]) - 2]
                    new_code = e["Defect code"] or ""
                    new_dec = e["Decision"] or r["decision"]
                    if new_code != r["code"]:
                        r["code"] = new_code
                        if new_dec in ("Auto", "Needs review"):
                            new_dec = "Changed"
                        changed = True
                    if new_dec != r["decision"]:
                        r["decision"] = new_dec
                        changed = True
                if changed:
                    bump(); st.rerun()
                st.caption("Set Decision to Accepted to keep a flagged code, pick a different code to change it, "
                           "or Excluded to leave the row out of the history.")

            # ---- save
            st.divider()
            problem = core.save_problem(draft)
            exists = ss.draft_key in months
            s1, s2, s3 = st.columns([1.3, 1, 2.5])
            s1.download_button("Download coded month (.xlsx)", core.coded_month_xlsx(draft, ss.draft_key),
                               file_name=f"Defect_Data_{core.month_label(ss.draft_key).replace(' ', '_')}_coded.xlsx",
                               help="Has the red and yellow highlighting. Upload it again later to pick up where you left off.")
            replace = True
            if exists and not problem:
                replace = s3.checkbox(f"{core.month_label(ss.draft_key)} is already in the history. Replace it.")
            if problem:
                s3.caption(problem)
            if s2.button("Save to history", type="primary", disabled=bool(problem) or not replace):
                keep = core.rows_to_save(draft)
                core.save_month(ss.draft_key, keep, user_name.strip(), ss.draft_file)
                excluded = len(draft) - len(keep)
                flash("success", f"{core.month_label(ss.draft_key)} saved: {len(keep)} defects added to the history"
                                 + (f", {excluded} excluded." if excluded else "."))
                ss.draft = None
                bump(); st.rerun()
            if st.button("Discard this upload"):
                ss.draft = None
                bump(); st.rerun()


# --------------------------------------------------------------------------
# Monthly summary
# --------------------------------------------------------------------------
def heat(df: pd.DataFrame):
    month_cols = [c for c in df.columns if c != "Total"]
    peak = max(1, int(df[month_cols].to_numpy().max())) if len(df) else 1

    def color(v):
        if not v:
            return "color: #8a958f"
        a = 0.12 + 0.68 * v / peak
        return f"background-color: rgba(29,106,76,{a:.2f}); color: {'white' if a > 0.45 else 'inherit'}"

    return (df.style.map(color, subset=month_cols)
            .set_properties(subset=["Total"], **{"font-weight": "bold", "color": "#A95F27"}))


with tab_summary:
    if not months:
        st.info("No months yet. Code and save a month and its counts appear here.")
    else:
        table = core.summary_table(months, settings)
        st.subheader("Defects by code and month")
        st.caption("Darker cells mean more defects with that code that month.")
        st.dataframe(heat(table), use_container_width=True, height=min(38 * (len(table) + 1) + 4, 1200))
        totals = table.drop(columns="Total").sum()
        st.caption("Defects per month: " + ", ".join(f"{k} {v}" for k, v in totals.items()))
        d1, d2, _ = st.columns([1, 1, 2])
        d1.download_button("Download summary (.xlsx)", core.summary_xlsx(months, settings),
                           file_name="Defect_Code_Monthly_Summary.xlsx", type="primary")
        d2.download_button("Download history (.xlsx)", core.history_xlsx(months),
                           file_name="Defect_Data_History.xlsx")


# --------------------------------------------------------------------------
# History
# --------------------------------------------------------------------------
with tab_history:
    st.subheader("Months in the history")
    st.caption("These months feed the auto-coding examples and the summary.")
    if months:
        st.dataframe(pd.DataFrame([{
            "Month": core.month_label(k), "Defects": len(v["rows"]), "Saved by": v.get("saved_by", ""),
            "Saved on": v.get("saved_at", ""), "Source file": v.get("source", ""),
        } for k, v in sorted(months.items(), reverse=True)]), hide_index=True, use_container_width=True)
    if can_edit and months:
        with st.expander("Delete a month"):
            k = st.selectbox("Month", sorted(months, reverse=True), format_func=core.month_label)
            sure = st.checkbox(f"Yes, permanently delete {core.month_label(k)}")
            if st.button("Delete month", disabled=not sure):
                core.delete_month(k)
                flash("success", f"{core.month_label(k)} was removed from the history.")
                st.rerun()
    if can_edit:
        with st.expander("Import a history workbook"):
            st.caption("One tab per month, named after the month (Jan, February, …), with the code in column E. "
                       "Months already in the history are replaced.")
            hf = st.file_uploader("History workbook", type=["xlsx"], key="hist_up")
            yr = st.number_input("Year of these months", 2000, 2100, date.today().year)
            if hf and st.button("Import"):
                found = core.import_history_xlsx(hf.getvalue(), int(yr))
                if not found:
                    st.error("No month tabs with coded rows were found.")
                else:
                    for key, rows in found.items():
                        core.save_month(key, rows, user_name.strip() or "Imported", hf.name)
                    flash("success", f"Imported {len(found)} months: " + ", ".join(core.month_label(k) for k in sorted(found)))
                    st.rerun()


# --------------------------------------------------------------------------
# Codes & rules
# --------------------------------------------------------------------------
with tab_settings:
    left, right = st.columns(2)
    with left:
        st.subheader("Defect codes")
        st.caption("Auto-coding only suggests codes from this list (plus any code already used in the history).")
        st.dataframe(pd.DataFrame({"Defect code": settings["codes"]}), hide_index=True,
                     use_container_width=True, height=420)
        if can_edit:
            new = st.text_input("New defect code")
            a, r = st.columns(2)
            if a.button("Add code", disabled=not new.strip()):
                core.save_settings(settings["codes"] + [new], settings["rules"])
                flash("success", f"Added {new.strip()}.")
                st.rerun()
            drop = r.selectbox("Remove a code", [""] + settings["codes"])
            if drop and r.button(f"Remove {drop}"):
                core.save_settings([c for c in settings["codes"] if c != drop], settings["rules"])
                flash("success", f"Removed {drop}. Past months that used it still keep it.")
                st.rerun()
    with right:
        st.subheader("Coding rules")
        st.caption("One rule per line. Auto-coding follows these ahead of past examples.")
        text = st.text_area("Rules", "\n".join(settings["rules"]), height=320,
                            disabled=not can_edit, label_visibility="collapsed")
        if can_edit and st.button("Save rules", type="primary"):
            core.save_settings(settings["codes"], text.split("\n"))
            flash("success", "Rules saved.")
            st.rerun()
