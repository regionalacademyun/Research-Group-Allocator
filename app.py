# -*- coding: utf-8 -*-
"""
RAUN Project Group Allocation Dashboard
Decision-support tool for matching participants to research projects
based on declared preferences and project interest scores.
"""

import io
import math
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="RAUN Project Allocator",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# STYLE
# =========================================================
st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 1450px;
    }
    .step-card {
        border-radius: 20px;
        padding: 1rem 1.2rem;
        background: linear-gradient(135deg, rgba(248,250,252,1), rgba(239,246,255,0.88));
        border: 1px solid rgba(148,163,184,0.22);
        box-shadow: 0 10px 24px rgba(15,23,42,0.04);
        margin-bottom: 1rem;
    }
    .soft-note {
        border-left: 4px solid #2563eb;
        background: rgba(37,99,235,0.08);
        padding: 0.8rem 1rem;
        border-radius: 12px;
        margin: 0.6rem 0 1rem 0;
    }
    .warn-note {
        border-left: 4px solid #d97706;
        background: rgba(245,158,11,0.10);
        padding: 0.8rem 1rem;
        border-radius: 12px;
        margin: 0.6rem 0 1rem 0;
    }
    .good-note {
        border-left: 4px solid #16a34a;
        background: rgba(34,197,94,0.10);
        padding: 0.8rem 1rem;
        border-radius: 12px;
        margin: 0.6rem 0 1rem 0;
    }
    .tiny {
        font-size: 0.92rem;
        color: #475569;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# PROJECT MASTER
# =========================================================
PROJECTS = {
    1: "Filling the Digital Gap: A Model for Rural Digital Transformation (UNDP)",
    2: "Artificial Intelligence and National Gender Assessment in Asia (UNDP)",
    3: "Nature-Based Solutions for Resilient Energy and Water Infrastructure in the Caribbean: Unlocking Innovation (UNFCCC)",
    4: "Biting less than could be chewed? The food governance of post-harvest losses (Kenya, Malawi and Zambia) (UNIDO)",
    5: "Why Coffee - Contextualising the UNIDO ACT Programme`s objectives (UNIDO)",
    6: "Why Coffee - Contextualising the UNIDO ACT Programme`s objectives (UNIDO)",
    7: "Creating Fair and Inclusive Value Distribution in Global Supply Chains (UNIDO)",
    8: "Youth, Science, and the Transformation of Agrifood Systems: Advancing FAO’s Vision for the Future (FAO)",
    9: "The role of Security Sector Governance and Reform (SSG/R) in the OSCE conflict cycle toolbox (OSCE)",
    10: "The road to the future: international organisations and their strategies (OSCE)",
    11: "The role of multilateral institutions and the importance of multilateralism in the current complex geopolitical context (OSCE)",
    12: "Responsible AI in the defence sector (OSCE)",
    13: "Youth news consumption patterns and expectations: media accountability and codified ethical standards in the era of news influencers (OSCE)",
    14: "Disinformation legislation and media freedom (OSCE)",
}


def build_default_projects_df() -> pd.DataFrame:
    rows = []
    for pid, title in PROJECTS.items():
        rows.append(
            {
                "Project ID": pid,
                "Project Label": f"Research project {pid}",
                "Project Title": title,
                "Active": True,
                "Min Capacity": 2,
                "Target Capacity": 3,
                "Max Capacity": 4,
                "Priority Weight": 1.0,
                "Notes": "",
            }
        )
    return pd.DataFrame(rows)


# =========================================================
# HELPERS
# =========================================================
def safe_int(value, default=0):
    try:
        if pd.isna(value):
            return default
        return int(float(value))
    except Exception:
        return default


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def find_col(df: pd.DataFrame, options: List[str]) -> str:
    norm_map = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}
    for opt in options:
        key = re.sub(r"\s+", " ", opt.strip().lower())
        if key in norm_map:
            return norm_map[key]
    return ""


def extract_project_number(text: str):
    if pd.isna(text):
        return np.nan
    s = str(text).strip().lower()
    m = re.search(r"project\s*(\d+)", s)
    if m:
        return int(m.group(1))
    m = re.search(r"research\s*project\s*(\d+)", s)
    if m:
        return int(m.group(1))
    digits = re.findall(r"\d+", s)
    if digits:
        return int(digits[0])
    return np.nan


def detect_response_sheet(excel_file):
    workbook = pd.ExcelFile(excel_file)
    best_sheet = None
    best_df = None
    best_score = -1

    for sheet_name in workbook.sheet_names:
        try:
            temp_df = pd.read_excel(workbook, sheet_name=sheet_name)
        except Exception:
            continue

        cols = [str(c).strip().lower() for c in temp_df.columns]
        score = 0

        if any("email" in c for c in cols):
            score += 3
        if any("first choice" in c for c in cols):
            score += 4
        if any("second choice" in c for c in cols):
            score += 4
        if any("third choice" in c for c in cols):
            score += 4
        score += sum(1 for c in cols if "project" in c)
        score += sum(1 for c in cols if "research project" in c)

        if score > best_score:
            best_score = score
            best_sheet = sheet_name
            best_df = temp_df.copy()

    return best_df, best_sheet, workbook.sheet_names


def normalize_responses(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def safe_series(frame: pd.DataFrame, col_name: str) -> pd.Series:
        if col_name and col_name in frame.columns:
            return frame[col_name].fillna("").astype(str).str.strip()
        return pd.Series([""] * len(frame), index=frame.index, dtype="object")

    email_col = find_col(df, ["Email", "E-mail", "Email address"])
    name_col = find_col(df, ["Name and surname", "Full name", "Name"])
    first_col = find_col(df, ["First name(s)", "First name", "First Name"])
    middle_col = find_col(df, ["Middle name(s)", "Middle name", "Middle Name"])
    last_col = find_col(df, ["Surname", "Last name", "Last Name"])
    ts_col = find_col(df, ["Timestamp"])

    if email_col:
        df = df.rename(columns={email_col: "Email"})
    else:
        df["Email"] = ""

    if name_col:
        df = df.rename(columns={name_col: "Full Name"})
    else:
        first_series = safe_series(df, first_col)
        middle_series = safe_series(df, middle_col)
        last_series = safe_series(df, last_col)
        full_name = (first_series + " " + middle_series + " " + last_series)
        df["Full Name"] = full_name.str.replace(r"\s+", " ", regex=True).str.strip()

    if ts_col:
        df = df.rename(columns={ts_col: "Timestamp"})
    else:
        df["Timestamp"] = pd.NaT

    score_cols = {}
    for pid, title in PROJECTS.items():
        candidates = [
            f"Project {pid}:",
            f"Project {pid}",
            f"Research project {pid}",
            title,
            title.split("(")[0].strip(),
        ]
        for c in df.columns:
            cl = str(c).strip().lower()
            if any(str(token).strip().lower() in cl for token in candidates if str(token).strip()):
                score_cols[pid] = c
                break

    for pid in PROJECTS:
        colname = f"Score P{pid}"
        if pid in score_cols:
            df[colname] = pd.to_numeric(df[score_cols[pid]], errors="coerce")
        else:
            df[colname] = np.nan

    first_choice_col = find_col(df, ["What would be your first choice?", "First choice"])
    second_choice_col = find_col(df, ["What would be your second choice?", "Second choice"])
    third_choice_col = find_col(df, ["What would be your third choice?", "Third choice"])

    df["Choice 1"] = df[first_choice_col].apply(extract_project_number) if first_choice_col else np.nan
    df["Choice 2"] = df[second_choice_col].apply(extract_project_number) if second_choice_col else np.nan
    df["Choice 3"] = df[third_choice_col].apply(extract_project_number) if third_choice_col else np.nan

    if "Participant ID" not in df.columns:
        df["Participant ID"] = range(1, len(df) + 1)

    df["Full Name"] = df["Full Name"].fillna("").astype(str).str.strip()
    df["Email"] = df["Email"].fillna("").astype(str).str.strip().str.lower()

    for pid in PROJECTS:
        df[f"Score P{pid}"] = pd.to_numeric(df[f"Score P{pid}"], errors="coerce")

    valid_pids = set(PROJECTS.keys())
    for col in ["Choice 1", "Choice 2", "Choice 3"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df[col] = df[col].where(df[col].isin(valid_pids), np.nan)

    df["Any Missing Scores"] = df[[f"Score P{pid}" for pid in PROJECTS]].isna().any(axis=1)
    df["Duplicate Choices"] = df[["Choice 1", "Choice 2", "Choice 3"]].apply(
        lambda r: len([x for x in r.dropna().tolist()]) != len(set(r.dropna().tolist())),
        axis=1,
    )
    df["Missing Any Choice"] = df[["Choice 1", "Choice 2", "Choice 3"]].isna().any(axis=1)

    return df


def normalize_projects_input(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    needed = {
        "Project ID": 0,
        "Project Label": "",
        "Project Title": "",
        "Active": True,
        "Min Capacity": 2,
        "Target Capacity": 3,
        "Max Capacity": 4,
        "Priority Weight": 1.0,
        "Notes": "",
    }
    for col, default in needed.items():
        if col not in df.columns:
            df[col] = default

    df["Project ID"] = pd.to_numeric(df["Project ID"], errors="coerce")
    df = df[df["Project ID"].isin(list(PROJECTS.keys()))].copy()
    df["Project ID"] = df["Project ID"].astype(int)
    df["Project Label"] = df["Project Label"].fillna("").astype(str)
    df["Project Title"] = df["Project Title"].fillna("").astype(str)
    df["Active"] = df["Active"].astype(bool)
    df["Min Capacity"] = pd.to_numeric(df["Min Capacity"], errors="coerce").fillna(2).astype(int)
    df["Target Capacity"] = pd.to_numeric(df["Target Capacity"], errors="coerce").fillna(3).astype(int)
    df["Max Capacity"] = pd.to_numeric(df["Max Capacity"], errors="coerce").fillna(4).astype(int)
    df["Priority Weight"] = pd.to_numeric(df["Priority Weight"], errors="coerce").fillna(1.0)
    df["Notes"] = df["Notes"].fillna("").astype(str)

    df["Target Capacity"] = df[["Target Capacity", "Min Capacity"]].max(axis=1)
    df["Max Capacity"] = df[["Max Capacity", "Target Capacity"]].max(axis=1)

    return df.sort_values("Project ID").reset_index(drop=True)


def compute_data_quality(df: pd.DataFrame) -> dict:
    duplicate_email_count = int(df["Email"].duplicated(keep=False).sum()) if "Email" in df.columns else 0
    missing_email_count = int((df["Email"].astype(str).str.strip() == "").sum()) if "Email" in df.columns else 0
    missing_name_count = int((df["Full Name"].astype(str).str.strip() == "").sum()) if "Full Name" in df.columns else 0

    return {
        "Rows": len(df),
        "Missing Emails": missing_email_count,
        "Duplicate Emails": duplicate_email_count,
        "Missing Names": missing_name_count,
        "Missing Any Choice": int(df["Missing Any Choice"].sum()),
        "Duplicate Choices": int(df["Duplicate Choices"].sum()),
        "Missing Any Scores": int(df["Any Missing Scores"].sum()),
    }


def deduplicate_latest(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    if "Timestamp" in work.columns:
        work["Timestamp Parsed"] = pd.to_datetime(work["Timestamp"], errors="coerce")
    else:
        work["Timestamp Parsed"] = pd.NaT

    work["Email Key"] = work["Email"].fillna("").astype(str).str.strip().str.lower()
    work["Email Key"] = work["Email Key"].replace("", np.nan)

    no_email = work[work["Email Key"].isna()].copy()
    has_email = work[work["Email Key"].notna()].copy()

    if not has_email.empty:
        has_email = has_email.sort_values(["Email Key", "Timestamp Parsed"], ascending=[True, True])
        has_email = has_email.drop_duplicates(subset=["Email Key"], keep="last")

    out = pd.concat([has_email, no_email], ignore_index=True)
    out = out.drop(columns=[c for c in ["Timestamp Parsed", "Email Key"] if c in out.columns])
    out = out.sort_values("Participant ID").reset_index(drop=True)
    return out


def rank_bonus(choice_rank: int, w1: float, w2: float, w3: float) -> float:
    if choice_rank == 1:
        return w1
    if choice_rank == 2:
        return w2
    if choice_rank == 3:
        return w3
    return 0.0


def get_choice_rank(row: pd.Series, pid: int) -> int:
    if row.get("Choice 1") == pid:
        return 1
    if row.get("Choice 2") == pid:
        return 2
    if row.get("Choice 3") == pid:
        return 3
    return 0


def allocation_priority(row: pd.Series) -> Tuple:
    scores = [row.get(f"Score P{pid}") for pid in PROJECTS]
    valid_scores = [s for s in scores if pd.notna(s)]
    max_score = max(valid_scores) if valid_scores else 0
    mean_score = np.mean(valid_scores) if valid_scores else 0
    spread = max_score - mean_score
    missing_choices = int(row.get("Missing Any Choice", False))
    duplicate_choices = int(row.get("Duplicate Choices", False))
    return (-spread, missing_choices, duplicate_choices, row.get("Participant ID", 0))


def recommend_capacity_pattern(n_participants: int, n_projects: int, strategy: str) -> dict:
    if n_projects <= 0:
        return {
            "preferred": 3,
            "comfort_max": 4,
            "rare_low": 2,
            "absolute_max": 4,
            "avg": 0,
            "mix_text": "No active projects yet.",
        }

    avg = n_participants / n_projects if n_projects > 0 else 0

    if strategy == "Strict small groups":
        preferred = 3
        comfort_max = 3
        rare_low = 2
        absolute_max = 4
    elif strategy == "High-pressure mode":
        preferred = 4
        comfort_max = 4
        rare_low = 2
        absolute_max = 5
    else:
        preferred = 3
        comfort_max = 4
        rare_low = 2
        absolute_max = 4

    n_hi = math.ceil(n_participants - (preferred * n_projects)) if avg > preferred else 0
    n_hi = max(0, min(n_projects, n_hi))
    n_lo = n_projects - n_hi

    hi_size = min(comfort_max, max(preferred, math.ceil(avg)))
    lo_size = preferred if avg >= preferred else max(rare_low, math.floor(avg))

    if strategy == "Manual only":
        mix_text = "Manual mode selected. The app will not push a recommended pattern."
    else:
        if avg >= preferred:
            mix_text = f"Suggested pattern: about {n_hi} groups of {hi_size} and {n_lo} groups of {lo_size}."
        else:
            mix_text = f"Suggested pattern: several groups may need {lo_size} participants or fewer. Consider reducing active topics if this is undesirable."

    return {
        "preferred": preferred,
        "comfort_max": comfort_max,
        "rare_low": rare_low,
        "absolute_max": absolute_max,
        "avg": round(avg, 2),
        "mix_text": mix_text,
    }


def build_topic_demand_table(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for pid, title in PROJECTS.items():
        first_count = int((df["Choice 1"] == pid).sum()) if "Choice 1" in df.columns else 0
        second_count = int((df["Choice 2"] == pid).sum()) if "Choice 2" in df.columns else 0
        third_count = int((df["Choice 3"] == pid).sum()) if "Choice 3" in df.columns else 0
        score_col = f"Score P{pid}"
        avg_score = round(pd.to_numeric(df[score_col], errors="coerce").mean(), 2) if score_col in df.columns else np.nan
        top_score_count = int((pd.to_numeric(df[score_col], errors="coerce") == 5).sum()) if score_col in df.columns else 0
        weighted_interest = first_count * 3 + second_count * 2 + third_count * 1

        rows.append(
            {
                "Project ID": pid,
                "Project": f"Research project {pid}",
                "Project Title": title,
                "1st Choice Count": first_count,
                "2nd Choice Count": second_count,
                "3rd Choice Count": third_count,
                "Weighted Preference Demand": weighted_interest,
                "Average Topic Score": avg_score,
                "Score 5 Count": top_score_count,
            }
        )
    return pd.DataFrame(rows)


@dataclass
class AllocationConfig:
    choice1_weight: float = 20.0
    choice2_weight: float = 12.0
    choice3_weight: float = 6.0
    score_weight: float = 8.0
    priority_weight_factor: float = 3.0
    target_balance_penalty: float = 4.0
    min_size_penalty: float = 5.0
    max_capacity_penalty: float = 100000.0
    outside_top3_allowed: bool = True
    outside_top3_penalty: float = 18.0


def candidate_project_utility(
    row: pd.Series,
    pid: int,
    assigned_count: int,
    project_row: pd.Series,
    cfg: AllocationConfig,
) -> float:
    score = safe_float(row.get(f"Score P{pid}"), default=0.0)
    choice_rank = get_choice_rank(row, pid)
    choice_component = rank_bonus(choice_rank, cfg.choice1_weight, cfg.choice2_weight, cfg.choice3_weight)
    score_component = score * cfg.score_weight
    priority_component = safe_float(project_row.get("Priority Weight", 1.0), 1.0) * cfg.priority_weight_factor

    min_capacity = max(0, safe_int(project_row.get("Min Capacity", 0), 0))
    target_capacity = max(1, safe_int(project_row.get("Target Capacity", 1), 1))
    max_capacity = safe_int(project_row.get("Max Capacity", 0), 0)

    balance_penalty = (assigned_count / target_capacity) * cfg.target_balance_penalty
    hard_penalty = cfg.max_capacity_penalty if assigned_count >= max_capacity else 0.0
    underfill_bonus = (min_capacity - assigned_count) * cfg.min_size_penalty if assigned_count < min_capacity else 0.0

    outside_penalty = 0.0
    if choice_rank == 0:
        if not cfg.outside_top3_allowed:
            outside_penalty = cfg.max_capacity_penalty
        else:
            outside_penalty = cfg.outside_top3_penalty

    return choice_component + score_component + priority_component + underfill_bonus - balance_penalty - outside_penalty - hard_penalty


def allocate_participants(responses_df: pd.DataFrame, projects_df: pd.DataFrame, cfg: AllocationConfig):
    active_projects = projects_df[projects_df["Active"] == True].copy()
    active_map = {int(r["Project ID"]): r for _, r in active_projects.iterrows()}

    assignments = []
    exceptions = []
    project_counts = {pid: 0 for pid in active_map.keys()}

    work = responses_df.copy()
    work["_priority"] = work.apply(allocation_priority, axis=1)
    work = work.sort_values(by="_priority", kind="stable").drop(columns=["_priority"]).reset_index(drop=True)

    for _, row in work.iterrows():
        pid_options = list(active_map.keys())
        scored_options = []

        for pid in pid_options:
            utility = candidate_project_utility(
                row=row,
                pid=pid,
                assigned_count=project_counts.get(pid, 0),
                project_row=active_map[pid],
                cfg=cfg,
            )
            scored_options.append((pid, utility))

        scored_options.sort(key=lambda x: x[1], reverse=True)

        if not scored_options:
            exceptions.append(
                {
                    "Participant ID": row["Participant ID"],
                    "Full Name": row["Full Name"],
                    "Issue Type": "No active project",
                    "Details": "There are no active projects available.",
                }
            )
            continue

        chosen_pid, chosen_utility = scored_options[0]
        chosen_project = active_map[chosen_pid]
        project_counts[chosen_pid] += 1

        chosen_rank = get_choice_rank(row, chosen_pid)
        chosen_score = safe_float(row.get(f"Score P{chosen_pid}"), 0.0)

        reason = []
        if chosen_rank == 1:
            reason.append("Matched to 1st choice")
        elif chosen_rank == 2:
            reason.append("Matched to 2nd choice")
        elif chosen_rank == 3:
            reason.append("Matched to 3rd choice")
        else:
            reason.append("Allocated outside top 3")

        reason.append(f"Topic score = {chosen_score:.0f}/5")

        assignments.append(
            {
                "Participant ID": row["Participant ID"],
                "Full Name": row["Full Name"],
                "Email": row["Email"],
                "Choice 1": row.get("Choice 1"),
                "Choice 2": row.get("Choice 2"),
                "Choice 3": row.get("Choice 3"),
                "Allocated Project ID": chosen_pid,
                "Allocated Project": f"Research project {chosen_pid}",
                "Allocated Project Title": chosen_project["Project Title"],
                "Allocated Score": chosen_score,
                "Matched Preference Rank": chosen_rank if chosen_rank > 0 else "Outside top 3",
                "Utility Score": round(chosen_utility, 2),
                "Allocation Reason": " | ".join(reason),
            }
        )

        if chosen_rank == 0:
            exceptions.append(
                {
                    "Participant ID": row["Participant ID"],
                    "Full Name": row["Full Name"],
                    "Issue Type": "Outside top 3 allocation",
                    "Details": f"Assigned to project {chosen_pid} with score {chosen_score:.0f}/5 because preferred projects were full or less competitive under the current settings.",
                }
            )

        if chosen_score <= 2:
            exceptions.append(
                {
                    "Participant ID": row["Participant ID"],
                    "Full Name": row["Full Name"],
                    "Issue Type": "Low interest allocation",
                    "Details": f"Assigned to project {chosen_pid} with a low topic score ({chosen_score:.0f}/5). Manual review suggested.",
                }
            )

    alloc_df = pd.DataFrame(assignments).sort_values("Participant ID").reset_index(drop=True) if assignments else pd.DataFrame()

    stats_rows = []
    for _, prow in active_projects.iterrows():
        pid = int(prow["Project ID"])
        assigned = int((alloc_df["Allocated Project ID"] == pid).sum()) if not alloc_df.empty else 0
        target = safe_int(prow["Target Capacity"], 0)
        max_cap = safe_int(prow["Max Capacity"], 0)
        min_cap = safe_int(prow["Min Capacity"], 0)

        stats_rows.append(
            {
                "Project ID": pid,
                "Project": f"Research project {pid}",
                "Project Title": prow["Project Title"],
                "Assigned": assigned,
                "Min Capacity": min_cap,
                "Target Capacity": target,
                "Max Capacity": max_cap,
                "Gap to Target": target - assigned,
                "Over Max By": max(0, assigned - max_cap),
                "Below Min By": max(0, min_cap - assigned),
                "Fill % of Target": round((assigned / target) * 100, 1) if target > 0 else 0,
            }
        )

    stats_df = pd.DataFrame(stats_rows).sort_values("Project ID").reset_index(drop=True)
    exc_df = pd.DataFrame(exceptions) if exceptions else pd.DataFrame(columns=["Participant ID", "Full Name", "Issue Type", "Details"])
    return alloc_df, stats_df, exc_df


def compute_allocation_summary(alloc_df: pd.DataFrame, stats_df: pd.DataFrame) -> dict:
    if alloc_df is None or alloc_df.empty:
        return {
            "Participants": 0,
            "1st Choice": 0,
            "2nd Choice": 0,
            "3rd Choice": 0,
            "Outside Top 3": 0,
            "Avg Assigned Score": 0,
            "Projects Used": 0,
        }

    ranks = alloc_df["Matched Preference Rank"].astype(str)
    return {
        "Participants": len(alloc_df),
        "1st Choice": int((ranks == "1").sum()),
        "2nd Choice": int((ranks == "2").sum()),
        "3rd Choice": int((ranks == "3").sum()),
        "Outside Top 3": int((ranks == "Outside top 3").sum()),
        "Avg Assigned Score": round(pd.to_numeric(alloc_df["Allocated Score"], errors="coerce").fillna(0).mean(), 2),
        "Projects Used": int(stats_df[stats_df["Assigned"] > 0]["Project ID"].nunique()) if stats_df is not None and not stats_df.empty else 0,
    }


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output.getvalue()


# =========================================================
# SESSION STATE
# =========================================================
for key, default in {
    "projects_working_df": None,
    "alloc_df": None,
    "project_stats_df": None,
    "exceptions_df": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.projects_working_df is None:
    st.session_state.projects_working_df = build_default_projects_df().copy()


# =========================================================
# HEADER
# =========================================================
st.title("RAUN Project Group Allocation Dashboard")
st.caption("A step-by-step decision support tool for allocating participants to research projects using preferences, topic scores, and capacity rules.")

with st.expander("What this tool is doing", expanded=False):
    st.write(
        "This app helps the admin create research groups in a transparent way. "
        "It first reads the form responses, then checks the quality of the data, then lets the admin define project capacities, "
        "and finally generates a suggested allocation. The app does not replace human judgement. It gives the admin a strong starting point."
    )

# =========================================================
# STEP 1 — UPLOAD
# =========================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown("## Step 1 — Upload the Google Form export")
st.markdown('<div class="tiny">Upload the Excel or CSV export from the participant preference form.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload participant responses",
    type=["xlsx", "csv"],
    help="Use the Google Form export. Excel or CSV is fine.",
)

if uploaded_file is None:
    st.info("Please upload the participant response file to begin.")
    st.stop()

sheet_debug_text = ""

if uploaded_file.name.lower().endswith(".xlsx"):
    responses_raw, detected_sheet, all_sheets = detect_response_sheet(uploaded_file)
    sheet_debug_text = f"Detected response sheet: {detected_sheet} | Available sheets: {', '.join(all_sheets)}"
else:
    responses_raw = pd.read_csv(uploaded_file)
    detected_sheet = "CSV upload"
    all_sheets = ["CSV upload"]
    sheet_debug_text = "CSV file detected."

responses = normalize_responses(responses_raw)
quality = compute_data_quality(responses)

st.success(f"File loaded. Found {len(responses)} response rows.")
st.caption(sheet_debug_text)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", quality["Rows"])
m2.metric("Missing any choice", quality["Missing Any Choice"])
m3.metric("Duplicate choices", quality["Duplicate Choices"])
m4.metric("Missing any scores", quality["Missing Any Scores"])

with st.expander("Preview raw normalized data", expanded=False):
    preview_cols = [
        "Participant ID",
        "Full Name",
        "Email",
        "Choice 1",
        "Choice 2",
        "Choice 3",
    ] + [f"Score P{i}" for i in range(1, 15)]
    show_cols = [c for c in preview_cols if c in responses.columns]
    st.dataframe(responses[show_cols], width="stretch", height=360)

with st.expander("Show column mapping diagnostics", expanded=False):
    mapping_rows = []
    for pid in PROJECTS:
        score_col = f"Score P{pid}"
        non_null = int(pd.to_numeric(responses.get(score_col), errors="coerce").notna().sum()) if score_col in responses.columns else 0
        mapping_rows.append({"Field": score_col, "Non-empty values detected": non_null})

    mapping_rows.extend(
        [
            {"Field": "Choice 1", "Non-empty values detected": int(pd.to_numeric(responses.get("Choice 1"), errors="coerce").notna().sum())},
            {"Field": "Choice 2", "Non-empty values detected": int(pd.to_numeric(responses.get("Choice 2"), errors="coerce").notna().sum())},
            {"Field": "Choice 3", "Non-empty values detected": int(pd.to_numeric(responses.get("Choice 3"), errors="coerce").notna().sum())},
            {"Field": "Full Name", "Non-empty values detected": int((responses.get("Full Name", pd.Series(dtype=str)).astype(str).str.strip() != "").sum())},
            {"Field": "Email", "Non-empty values detected": int((responses.get("Email", pd.Series(dtype=str)).astype(str).str.strip() != "").sum())},
        ]
    )
    st.dataframe(pd.DataFrame(mapping_rows), width="stretch", height=340)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# STEP 2 — CLEANING / QUALITY
# =========================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown("## Step 2 — Check the data quality")
st.markdown('<div class="tiny">This step helps the admin spot issues before allocation.</div>', unsafe_allow_html=True)

q1, q2, q3, q4 = st.columns(4)
q1.metric("Duplicate emails", quality["Duplicate Emails"])
q2.metric("Missing emails", quality["Missing Emails"])
q3.metric("Missing names", quality["Missing Names"])
q4.metric("Usable rows after dedupe", len(deduplicate_latest(responses)))

st.markdown(
    """
    <div class="soft-note">
    <b>Plain explanation:</b><br>
    Sometimes participants submit the form more than once. The app can keep the latest response per email address.<br>
    This reduces confusion before group allocation begins.
    </div>
    """,
    unsafe_allow_html=True,
)

keep_latest = st.checkbox(
    "If duplicate emails exist, keep only the latest response per email",
    value=True,
    help="Recommended when participants may have resubmitted the form.",
)

if keep_latest:
    responses_working = deduplicate_latest(responses)
else:
    responses_working = responses.copy()

issue_rows = responses_working[
    responses_working[["Any MissingScores" if False else "Any Missing Scores", "Duplicate Choices", "Missing Any Choice"]].copy()
] if False else responses_working[
    responses_working[["Any Missing Scores", "Duplicate Choices", "Missing Any Choice"]].any(axis=1)
].copy()

with st.expander("Show rows that may need admin attention", expanded=False):
    if issue_rows.empty:
        st.success("No major response issues detected in the current working data.")
    else:
        st.dataframe(
            issue_rows[
                [
                    "Participant ID",
                    "Full Name",
                    "Email",
                    "Choice 1",
                    "Choice 2",
                    "Choice 3",
                    "Any Missing Scores",
                    "Duplicate Choices",
                    "Missing Any Choice",
                ]
            ],
            width="stretch",
            height=320,
        )

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# STEP 3 — STRATEGY + PROJECT TABLE
# =========================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown("## Step 3 — Choose the group size strategy and configure the research projects")
st.markdown('<div class="tiny">First choose the strategy. Then adjust project capacities if needed. The app will recommend a sensible size pattern based on the number of participants and active topics.</div>', unsafe_allow_html=True)

strategy_col1, strategy_col2 = st.columns([1.3, 1.7])
with strategy_col1:
    capacity_strategy = st.selectbox(
        "Capacity strategy",
        ["Balanced RAUN mode", "Strict small groups", "High-pressure mode", "Manual only"],
        index=0,
        help="Balanced RAUN mode is the recommended default.",
    )

with strategy_col2:
    st.markdown(
        '<div class="soft-note"><b>Strategy guide:</b><br>'
        'Strict small groups = mostly 3, rarely more.<br>'
        'Balanced RAUN mode = target 3, common 4, rare low exceptions.<br>'
        'High-pressure mode = target 4, used when participants are many and topics are limited.<br>'
        'Manual only = you set all capacities yourself.'
        '</div>',
        unsafe_allow_html=True,
    )

recommended = recommend_capacity_pattern(
    n_participants=len(responses_working),
    n_projects=int((normalize_projects_input(st.session_state.projects_working_df)["Active"] == True).sum()),
    strategy=capacity_strategy,
)

r1, r2, r3, r4 = st.columns(4)
r1.metric("Participants", len(responses_working))
r2.metric("Active topics", int((normalize_projects_input(st.session_state.projects_working_df)["Active"] == True).sum()))
r3.metric("Average per topic", recommended["avg"])
r4.metric("Preferred size", recommended["preferred"])

st.markdown(
    f"""
    <div class="soft-note">
    <b>Recommended planning pattern:</b><br>
    Preferred size: <b>{recommended['preferred']}</b><br>
    Comfort max: <b>{recommended['comfort_max']}</b><br>
    Rare low exception: <b>{recommended['rare_low']}</b><br>
    Absolute max: <b>{recommended['absolute_max']}</b><br>
    {recommended['mix_text']}
    </div>
    """,
    unsafe_allow_html=True,
)

apply_recommendation = st.checkbox(
    "Apply the recommended capacity pattern automatically to all active topics",
    value=(capacity_strategy != "Manual only"),
    help="You can still edit the project table afterward.",
)

projects_df = normalize_projects_input(st.session_state.projects_working_df)
if apply_recommendation and capacity_strategy != "Manual only":
    active_mask = projects_df["Active"] == True
    projects_df.loc[active_mask, "Target Capacity"] = recommended["preferred"]
    projects_df.loc[active_mask, "Max Capacity"] = recommended["comfort_max"]
    projects_df.loc[active_mask, "Min Capacity"] = recommended["rare_low"]

edited_projects = st.data_editor(
    projects_df,
    width="stretch",
    height=520,
    num_rows="fixed",
    column_config={
        "Active": st.column_config.CheckboxColumn("Active", help="Turn a project on or off for this round."),
        "Min Capacity": st.column_config.NumberColumn("Min Capacity", min_value=0, step=1),
        "Target Capacity": st.column_config.NumberColumn("Target Capacity", min_value=0, step=1),
        "Max Capacity": st.column_config.NumberColumn("Max Capacity", min_value=0, step=1),
        "Priority Weight": st.column_config.NumberColumn("Priority Weight", min_value=0.0, step=0.1),
        "Notes": st.column_config.TextColumn("Notes"),
    },
    key="project_editor",
)

st.session_state.projects_working_df = normalize_projects_input(edited_projects)
projects_df = normalize_projects_input(st.session_state.projects_working_df)
active_projects = projects_df[projects_df["Active"] == True].copy()

cap1, cap2, cap3, cap4 = st.columns(4)
cap1.metric("Active projects", len(active_projects))
cap2.metric("Total min capacity", int(active_projects["Min Capacity"].sum()) if not active_projects.empty else 0)
cap3.metric("Total target capacity", int(active_projects["Target Capacity"].sum()) if not active_projects.empty else 0)
cap4.metric("Total max capacity", int(active_projects["Max Capacity"].sum()) if not active_projects.empty else 0)

n_people = len(responses_working)
if not active_projects.empty:
    total_target = int(active_projects["Target Capacity"].sum())
    total_max = int(active_projects["Max Capacity"].sum())

    if total_max < n_people:
        st.markdown('<div class="warn-note"><b>Important:</b> Total max capacity is smaller than the number of participants. Increase capacity or reduce participants / active topics.</div>', unsafe_allow_html=True)
    elif total_target < n_people:
        st.markdown('<div class="warn-note"><b>Heads-up:</b> Target capacity is below the number of participants. That is okay if max capacity absorbs the overflow.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="good-note"><b>Good:</b> Current target capacity can absorb all participants.</div>', unsafe_allow_html=True)

st.markdown("### Topic demand studio")

demand_df = build_topic_demand_table(responses_working)

left_plot, right_plot = st.columns(2)
with left_plot:
    pref_fig = go.Figure()
    pref_fig.add_trace(go.Bar(x=demand_df["Project"], y=demand_df["1st Choice Count"], name="1st choice"))
    pref_fig.add_trace(go.Bar(x=demand_df["Project"], y=demand_df["2nd Choice Count"], name="2nd choice"))
    pref_fig.add_trace(go.Bar(x=demand_df["Project"], y=demand_df["3rd Choice Count"], name="3rd choice"))
    pref_fig.update_layout(
        barmode="stack",
        title="How many participants preferred each topic",
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="Project",
        yaxis_title="Participants",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    pref_fig.update_xaxes(showgrid=False)
    pref_fig.update_yaxes(gridcolor="rgba(148,163,184,0.18)")
    st.plotly_chart(pref_fig, width="stretch")

with right_plot:
    score_interest_fig = go.Figure()
    score_interest_fig.add_trace(
        go.Bar(
            x=demand_df["Project"],
            y=demand_df["Average Topic Score"],
            name="Average topic score",
            hovertemplate="<b>%{x}</b><br>Average score: %{y:.2f}<extra></extra>",
        )
    )
    score_interest_fig.update_layout(
        title="Average score participants gave each topic",
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="Project",
        yaxis_title="Average score (1-5)",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    score_interest_fig.update_xaxes(showgrid=False)
    score_interest_fig.update_yaxes(gridcolor="rgba(148,163,184,0.18)", range=[0, 5])
    st.plotly_chart(score_interest_fig, width="stretch")

with st.expander("Show topic demand table", expanded=False):
    st.dataframe(demand_df, width="stretch", height=340)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# STEP 4 — SETTINGS
# =========================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown("## Step 4 — Set the decision logic")
st.markdown('<div class="tiny">These sliders control how strongly the app values ranked choices, topic scores, and balancing across projects.</div>', unsafe_allow_html=True)

s1, s2, s3 = st.columns(3)
with s1:
    choice1_weight = st.slider("1st choice importance", min_value=0.0, max_value=40.0, value=20.0, step=1.0)
    choice2_weight = st.slider("2nd choice importance", min_value=0.0, max_value=30.0, value=12.0, step=1.0)

with s2:
    choice3_weight = st.slider("3rd choice importance", min_value=0.0, max_value=20.0, value=6.0, step=1.0)
    score_weight = st.slider("Topic score importance", min_value=0.0, max_value=15.0, value=8.0, step=0.5)

with s3:
    balance_penalty = st.slider("Balance pressure", min_value=0.0, max_value=12.0, value=4.0, step=0.5)
    outside_top3_allowed = st.checkbox("Allow allocation outside top 3 if needed", value=True)
    outside_top3_penalty = st.slider("Penalty for outside top 3", min_value=0.0, max_value=40.0, value=18.0, step=1.0)

priority_weight_factor = st.slider(
    "Project priority weight effect",
    min_value=0.0,
    max_value=10.0,
    value=3.0,
    step=0.5,
)

min_size_penalty = st.slider(
    "Pressure to keep groups above the minimum",
    min_value=0.0,
    max_value=12.0,
    value=5.0,
    step=0.5,
    help="Higher means the allocator will try harder to avoid leaving a project below its minimum size.",
)

cfg = AllocationConfig(
    choice1_weight=choice1_weight,
    choice2_weight=choice2_weight,
    choice3_weight=choice3_weight,
    score_weight=score_weight,
    priority_weight_factor=priority_weight_factor,
    target_balance_penalty=balance_penalty,
    min_size_penalty=min_size_penalty,
    outside_top3_allowed=outside_top3_allowed,
    outside_top3_penalty=outside_top3_penalty,
)

with st.expander("Explain these settings simply", expanded=False):
    st.write(
        "Think of the app as giving points to each project for each participant. "
        "It gives points if that project is the participant’s 1st, 2nd, or 3rd choice. "
        "It also gives points based on the score the participant gave that topic. "
        "Then it subtracts points if a project is becoming too full. "
        "It also tries not to leave projects below the minimum size. "
        "The project with the best final score wins."
    )

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# STEP 5 — RUN ALLOCATION
# =========================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown("## Step 5 — Generate the allocation")
st.markdown('<div class="tiny">When you are happy with the project table and the settings, run the allocation.</div>', unsafe_allow_html=True)

run_clicked = st.button("Generate project allocation", type="primary", width="stretch")

if run_clicked:
    alloc_df, project_stats_df, exceptions_df = allocate_participants(
        responses_df=responses_working,
        projects_df=projects_df,
        cfg=cfg,
    )
    st.session_state.alloc_df = alloc_df
    st.session_state.project_stats_df = project_stats_df
    st.session_state.exceptions_df = exceptions_df

if st.session_state.alloc_df is None:
    st.info("No allocation has been generated yet.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

alloc_df = st.session_state.alloc_df
project_stats_df = st.session_state.project_stats_df
exceptions_df = st.session_state.exceptions_df
summary = compute_allocation_summary(alloc_df, project_stats_df)

sm1, sm2, sm3, sm4, sm5, sm6 = st.columns(6)
sm1.metric("Participants", summary["Participants"])
sm2.metric("1st choice matched", summary["1st Choice"])
sm3.metric("2nd choice matched", summary["2nd Choice"])
sm4.metric("3rd choice matched", summary["3rd Choice"])
sm5.metric("Outside top 3", summary["Outside Top 3"])
sm6.metric("Avg assigned score", summary["Avg Assigned Score"])

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# RESULTS TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "Allocation Output",
        "Project Studio",
        "Participant Satisfaction",
        "Manual Review Studio",
        "Downloads",
    ]
)

with tab1:
    st.subheader("Allocation output")
    st.dataframe(alloc_df, width="stretch", height=460)

    st.markdown("### Manual review flags")
    if exceptions_df is not None and not exceptions_df.empty:
        st.dataframe(exceptions_df, width="stretch", height=260)
    else:
        st.success("No major manual review flags were produced by the current run.")

with tab2:
    st.subheader("Project studio")
    st.dataframe(project_stats_df, width="stretch", height=420)

    if not project_stats_df.empty:
        plot_df = project_stats_df.sort_values("Project ID")
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=plot_df["Project"],
                y=plot_df["Assigned"],
                name="Assigned",
                hovertemplate="<b>%{x}</b><br>Assigned: %{y}<extra></extra>",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=plot_df["Project"],
                y=plot_df["Target Capacity"],
                mode="lines+markers",
                name="Target capacity",
                line=dict(width=3, dash="dash"),
                hovertemplate="<b>%{x}</b><br>Target: %{y}<extra></extra>",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=plot_df["Project"],
                y=plot_df["Max Capacity"],
                mode="lines+markers",
                name="Max capacity",
                line=dict(width=3, dash="dot"),
                hovertemplate="<b>%{x}</b><br>Max: %{y}<extra></extra>",
            )
        )
        fig.update_layout(
            title="Project fill levels vs capacities",
            paper_bgcolor="white",
            plot_bgcolor="white",
            xaxis_title="Project",
            yaxis_title="Participants",
            margin=dict(l=20, r=20, t=60, b=20),
        )
        fig.update_xaxes(showgrid=False)
        fig.update_yaxes(gridcolor="rgba(148,163,184,0.18)")
        st.plotly_chart(fig, width="stretch")

with tab3:
    st.subheader("Participant satisfaction")

    ranks = alloc_df["Matched Preference Rank"].astype(str).value_counts().to_dict()
    pref_df = pd.DataFrame(
        {
            "Preference Outcome": ["1st choice", "2nd choice", "3rd choice", "Outside top 3"],
            "Count": [
                ranks.get("1", 0),
                ranks.get("2", 0),
                ranks.get("3", 0),
                ranks.get("Outside top 3", 0),
            ],
        }
    )

    sat_fig = go.Figure()
    sat_fig.add_trace(
        go.Bar(
            x=pref_df["Preference Outcome"],
            y=pref_df["Count"],
            hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>",
        )
    )
    sat_fig.update_layout(
        title="How well participant preferences were matched",
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="Outcome",
        yaxis_title="Participants",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    sat_fig.update_xaxes(showgrid=False)
    sat_fig.update_yaxes(gridcolor="rgba(148,163,184,0.18)")
    st.plotly_chart(sat_fig, width="stretch")

    score_by_project = alloc_df.groupby(["Allocated Project ID", "Allocated Project"])["Allocated Score"].mean().reset_index()
    score_by_project = score_by_project.sort_values("Allocated Project ID")
    score_fig = go.Figure()
    score_fig.add_trace(
        go.Bar(
            x=score_by_project["Allocated Project"],
            y=score_by_project["Allocated Score"],
            hovertemplate="<b>%{x}</b><br>Average assigned score: %{y:.2f}<extra></extra>",
        )
    )
    score_fig.update_layout(
        title="Average topic score of the allocated participants per project",
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="Project",
        yaxis_title="Average score (1-5)",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    score_fig.update_xaxes(showgrid=False)
    score_fig.update_yaxes(gridcolor="rgba(148,163,184,0.18)", range=[0, 5])
    st.plotly_chart(score_fig, width="stretch")

with tab4:
    st.subheader("Manual review studio")

    review_low = alloc_df[pd.to_numeric(alloc_df["Allocated Score"], errors="coerce").fillna(0) <= 2].copy()
    review_outside = alloc_df[alloc_df["Matched Preference Rank"].astype(str) == "Outside top 3"].copy()
    group_issues = project_stats_df[
        (project_stats_df["Assigned"] < project_stats_df["Min Capacity"])
        | (project_stats_df["Assigned"] > project_stats_df["Max Capacity"])
    ].copy()
    tiny_groups = project_stats_df[project_stats_df["Assigned"] <= 2].copy()

    mr1, mr2, mr3, mr4 = st.columns(4)
    mr1.metric("Low-interest allocations", len(review_low))
    mr2.metric("Outside top 3", len(review_outside))
    mr3.metric("Groups below min or above max", len(group_issues))
    mr4.metric("Tiny groups (≤2)", len(tiny_groups))

    with st.expander("Participants allocated to low-scored topics", expanded=True):
        if review_low.empty:
            st.success("No low-interest allocations under the current run.")
        else:
            st.dataframe(review_low, width="stretch", height=260)

    with st.expander("Participants allocated outside their top 3", expanded=True):
        if review_outside.empty:
            st.success("No outside-top-3 allocations under the current run.")
        else:
            st.dataframe(review_outside, width="stretch", height=260)

    with st.expander("Groups that break your size rules", expanded=True):
        if group_issues.empty:
            st.success("All active groups stay within the current min and max settings.")
        else:
            st.dataframe(group_issues, width="stretch", height=240)

    with st.expander("Small groups that may need manual judgement", expanded=False):
        if tiny_groups.empty:
            st.success("No tiny groups found.")
        else:
            st.dataframe(tiny_groups, width="stretch", height=220)

with tab5:
    st.subheader("Downloads")

    merged_export = pd.merge(
        responses_working,
        alloc_df[
            [
                "Participant ID",
                "Allocated Project ID",
                "Allocated Project",
                "Allocated Project Title",
                "Allocated Score",
                "Matched Preference Rank",
                "Utility Score",
                "Allocation Reason",
            ]
        ],
        on="Participant ID",
        how="left",
    )

    export_bytes = to_excel_bytes(
        {
            "Allocation": alloc_df,
            "Project Stats": project_stats_df,
            "Exceptions": exceptions_df,
            "Merged Working Data": merged_export,
            "Project Table": projects_df,
            "Topic Demand": demand_df,
        }
    )

    c1, c2 = st.columns(2)
    c1.download_button(
        "Download full Excel output",
        data=export_bytes,
        file_name="raun_project_allocation_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    c2.download_button(
        "Download allocation CSV",
        data=alloc_df.to_csv(index=False).encode("utf-8"),
        file_name="raun_project_allocation.csv",
        mime="text/csv",
        width="stretch",
    )

    st.markdown("### Suggested admin use")
    st.markdown(
        """
        1. Run the tool once with balanced settings.  
        2. Check who was placed outside their top 3.  
        3. Check who received projects they scored low.  
        4. Check tiny groups and overloaded groups.  
        5. Adjust capacities or weights if needed.  
        6. Finalize manually only for the few edge cases.
        """
    )