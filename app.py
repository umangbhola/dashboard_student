import os
from typing import List, Optional, Tuple, Dict

import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
from dateutil import parser as date_parser
import re
from collections import Counter

DEFAULT_XLSX = "Program Manager - Fortnightly school audit checklist - HUHT project, June-Dec 2025 (Responses).xlsx"
DEFAULT_SHEET = "Form Responses 1"

st.set_page_config(page_title="Program Manager - Fortnightly School Audit Dashboard", layout="wide")

SAT_DOMAIN = ["Not Satisfied", "Satisfied", "Very Satisfied"]
# Include common misspelling mapping when encountered
SAT_ALIASES: Dict[str, str] = {
    "not satisifed": "Not Satisfied",
    "not satisfied": "Not Satisfied",
    "satisfied": "Satisfied",
    "very satisfied": "Very Satisfied",
}
SAT_COLORS = ["#C00000", "#F5AE1B", "#007467"]

@st.cache_data(show_spinner=False)
def load_excel(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    try:
        df = pd.read_excel(path, sheet_name=sheet_name or 0, engine="openpyxl")
    except ValueError:
        # Sheet not found; fall back to first sheet
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    return df


def coerce_datetime(series: pd.Series) -> pd.Series:
    # Attempt to parse datetimes, leave untouched if fails
    try:
        parsed = pd.to_datetime(series, errors="coerce")
        if parsed.notna().sum() > 0:
            return parsed
    except Exception:
        pass
    return series


def infer_question_columns(df: pd.DataFrame) -> List[str]:
    # Heuristic: Exclude metadata columns like Timestamp, Email, Name, etc.
    metadata_like = {"timestamp", "time", "email", "name", "phone", "id", "score"}
    question_cols: List[str] = []
    for col in df.columns:
        lower = str(col).strip().lower()
        if any(token in lower for token in metadata_like):
            continue
        question_cols.append(col)
    return question_cols


def is_numeric(series: pd.Series) -> bool:
    if pd.api.types.is_numeric_dtype(series):
        return True
    # Try coercion
    coerced = pd.to_numeric(series, errors="coerce")
    return coerced.notna().mean() > 0.6


def find_classroom_indicator_columns(df: pd.DataFrame) -> List[str]:
    targets = [
        "Cleanliness",
        "Participation of ALL students",
        "Display of age-appropriate student-made materials",
        "Use and maintenance of TLMs",
        "Checking of corrections in exercise books",
        "Teaching methodology (Play-based, Experiential/Group activity)",
    ]
    cols: List[str] = []
    for col in df.columns:
        col_str = str(col)
        if "Classroom Indicators" in col_str and "Rate your observation" in col_str:
            for t in targets:
                if t in col_str:
                    cols.append(col)
                    break
    return cols


def find_corridors_columns(df: pd.DataFrame) -> List[str]:
    targets = [
        "Lesson plan conducive to display",
        "Dustbins installed and used",
        "Cleanliness",
        "Proper utilisation of the space",
        "Overall atmosphere of the school",
    ]
    cols: List[str] = []
    for col in df.columns:
        col_str = str(col)
        if "Corridors and Open Spaces" in col_str and "Rate your observation" in col_str:
            for t in targets:
                if t in col_str:
                    cols.append(col)
                    break
    return cols


def find_teacher_update_columns(df: pd.DataFrame) -> List[str]:
    targets = [
        "Update on ongoing modules/lesson plans",
        "Feedback on difficulties/special benefits of previous modules",
        "Class management (difficulties with students)",
        "Student development/improvement",
        "Takeaways from Group meetings/Training sessions",
        "Update on prescribed Manual reading",
    ]
    cols: List[str] = []
    for col in df.columns:
        col_str = str(col)
        if "Teacher Update" in col_str and "Rate your observation" in col_str:
            for t in targets:
                if t in col_str:
                    cols.append(col)
                    break
    return cols


def find_student_interaction_columns(df: pd.DataFrame) -> List[str]:
    targets = [
        "Feedback on the new curriculum",
        "Student-teacher relation",
        "Overall development/change",
    ]
    cols: List[str] = []
    for col in df.columns:
        col_str = str(col)
        if "Interaction with Students" in col_str and "Rate your observation" in col_str:
            for t in targets:
                if t in col_str:
                    cols.append(col)
                    break
    return cols


def find_curriculum_activity_columns(df: pd.DataFrame) -> List[str]:
    targets = [
        "Brotochari practice",
        "Yoga and Meditation",
        "Visual Art Integration within the Curriculum",
        "Pottery",
        "Hygiene and Organic Hygiene product preparation",
        "Gardening and agricultural resource utilization",
        "Music (Vocal/Flute)",
        "Textile (Eco-printing/Stitching/Weaving/Hand Charkha)",
        "Using computer",
        "Food Preparation",
        "Renewable energy (model making/recycling/reusing resources)",
    ]
    cols: List[str] = []
    for col in df.columns:
        col_str = str(col)
        if "Manual-based Curriculum Activities" in col_str and "Rate your observation" in col_str:
            for t in targets:
                if t in col_str:
                    cols.append(col)
                    break
    return cols


def shorten_indicator_label(full_col: str) -> str:
    # Extract content inside the last [...] if present; else return tail after known prefix
    if "]" in full_col and "[" in full_col and full_col.rfind("[") < full_col.rfind("]"):
        return full_col[full_col.rfind("[") + 1 : full_col.rfind("]")].strip()
    # Fallbacks based on known prefixes
    for prefix in [
        "Classroom Indicators – Rate your observation",
        "Corridors and Open Spaces-  Rate your observation",
        "Corridors and Open Spaces- Rate your observation",
        "Teacher Update – Rate your observation",
        "Interaction with Students – Rate your observation",
        "Manual-based Curriculum Activities – Rate your observation",
    ]:
        if prefix in full_col:
            return full_col.split(prefix, 1)[-1].strip(" -:")
    return full_col


# Custom label mappings by bracket content for specific sections
TEACHER_LABEL_MAP: Dict[str, str] = {
    "Update on ongoing modules/lesson plans": "Update on ongoing modules/lLsson Plans",
    "Feedback on difficulties/special benefits of previous modules": "Feedback on Difficulties or Benefits",
    "Class management (difficulties with students)": "Class management",
    "Student development/improvement": "Student Imporvement",
    "Takeaways from Group meetings/Training sessions": "Takeaways from Group Meetings",
    "Update on prescribed Manual reading": "Update on Manual Reading",
}

INTERACTION_LABEL_MAP: Dict[str, str] = {
    "Feedback on the new curriculum": "Curriculum Feedback",
    "Student-teacher relation": "Student-Teacher Relation",
    "Overall development/change": "Overall Development",
}

CURRICULUM_LABEL_MAP: Dict[str, str] = {
    "Brotochari practice": "Brotochari",
    "Yoga and Meditation": "Yoga and Meditation",
    "Visual Art Integration within the Curriculum": "Visual Art Integration",
    "Pottery": "Pottery",
    "Hygiene and Organic Hygiene product preparation": "Hygiene",
    "Gardening and agricultural resource utilization": "Gardening and Agriculture",
    "Music (Vocal/Flute)": "Vocal/Flute",
    "Textile (Eco-printing/Stitching/Weaving/Hand Charkha)": "Textile",
    "Using computer": "Computer Usage",
    "Food Preparation": "Food Prepration",
    "Renewable energy (model making/recycling/reusing resources)": "Renewable Energy",
}


def map_custom_label(full_col: str) -> str:
    label = shorten_indicator_label(full_col)
    if ("Teacher Update" in full_col and "Rate your observation" in full_col) and label in TEACHER_LABEL_MAP:
        return TEACHER_LABEL_MAP[label]
    if ("Interaction with Students" in full_col and "Rate your observation" in full_col) and label in INTERACTION_LABEL_MAP:
        return INTERACTION_LABEL_MAP[label]
    if ("Manual-based Curriculum Activities" in full_col and "Rate your observation" in full_col) and label in CURRICULUM_LABEL_MAP:
        return CURRICULUM_LABEL_MAP[label]
    return label


def wrap_label(text: str, width: int = 18) -> str:
    words = str(text).split()
    if not words:
        return str(text)
    lines: List[str] = []
    current = words[0]
    for w in words[1:]:
        if len(current) + 1 + len(w) <= width:
            current += " " + w
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return "\n".join(lines)


def normalize_satisfaction(value: str) -> str:
    v = str(value).strip()
    key = v.lower()
    return SAT_ALIASES.get(key, v)


def build_stacked_chart_vertical(df: pd.DataFrame, columns: List[str], title: Optional[str] = None) -> alt.Chart:
    if not columns:
        return alt.Chart(pd.DataFrame({"x": [], "y": []})).mark_bar()

    melted = (
        df[columns]
        .astype(object)
        .melt(var_name="IndicatorFull", value_name="Response")
    )

    melted["Indicator"] = melted["IndicatorFull"].apply(map_custom_label)
    melted["IndicatorWrapped"] = melted["Indicator"].apply(lambda s: wrap_label(s, width=22))
    melted["Response"] = melted["Response"].astype(str).replace({"nan": "(Missing)"}).map(normalize_satisfaction)

    counts = (
        melted.groupby(["Indicator", "IndicatorWrapped", "Response"])  # type: ignore
        .size()
        .reset_index(name="Count")
    )

    totals = counts.groupby(["Indicator", "IndicatorWrapped"])['Count'].sum().reset_index(name="Total")
    counts = counts.merge(totals, on=["Indicator", "IndicatorWrapped"], how="left")
    counts["Percent"] = (counts["Count"] / counts["Total"]).fillna(0.0)

    # Build domain for color scale based on SAT_DOMAIN intersection
    unique_responses = counts["Response"].unique().tolist()
    domain = [r for r in SAT_DOMAIN if r in unique_responses]

    x_enc = alt.X(
        "IndicatorWrapped:N",
        title="Indicator",
        axis=alt.Axis(labelAngle=-35, labelLimit=500, labelOverlap=False)
    )

    color_kwargs = {}
    if domain:
        color_kwargs = {"scale": alt.Scale(domain=domain, range=SAT_COLORS[: len(domain)])}

    base = alt.Chart(counts)
    if title:
        base = base.properties(title=title)

    bars = (
        base.mark_bar()
        .encode(
            x=x_enc,
            y=alt.Y("Percent:Q", stack="normalize", axis=alt.Axis(format="%"), title="Percent"),
            color=alt.Color("Response:N", legend=alt.Legend(title="Response"), **color_kwargs),
            order=alt.Order("Response:N", sort="ascending"),
            tooltip=["Indicator:N", "Response:N", alt.Tooltip("Percent:Q", format=".0%"), "Count:Q"],
        )
    )

    chart = bars.properties(height=460)
    return chart


def extract_observations_text(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    if column_name not in df.columns:
        return pd.DataFrame(columns=["term", "count"])  # empty

    texts = df[column_name].dropna().astype(str)
    if texts.empty:
        return pd.DataFrame(columns=["term", "count"])  # empty

    stopwords = set(
        [
            "the","a","an","and","or","but","to","of","in","on","for","with","without","at","by","from","as","is","are","was","were","be","been","being","it","its","this","that","these","those","we","they","you","i","he","she","him","her","his","hers","their","our","us","your",
            "there","here","also","very","much","so","can","could","may","might","should","would","will","shall","have","has","had","do","does","did","not","no","yes","if","then","than","because","due","since","into","out","over","under","more","most","less","least","few","many",
        ]
    )

    def tokenize(s: str) -> List[str]:
        s = s.lower()
        s = re.sub(r"[^a-z0-9\s]", " ", s)
        tokens = [t for t in s.split() if t and t not in stopwords and not t.isdigit() and len(t) > 2]
        return tokens

    unigram_counter: Counter = Counter()
    bigram_counter: Counter = Counter()

    for t in texts:
        toks = tokenize(t)
        unigram_counter.update(toks)
        bigrams = [f"{toks[i]} {toks[i+1]}" for i in range(len(toks)-1)]
        bigram_counter.update(bigrams)

    top_uni = unigram_counter.most_common(15)
    top_bi = bigram_counter.most_common(15)

    df_uni = pd.DataFrame(top_uni, columns=["term", "count"]) if top_uni else pd.DataFrame(columns=["term","count"])
    df_bi = pd.DataFrame(top_bi, columns=["term", "count"]) if top_bi else pd.DataFrame(columns=["term","count"])

    return df_uni, df_bi


def observations_charts(df: pd.DataFrame, column_name: str) -> Optional[alt.VConcatChart]:
    uni, bi = extract_observations_text(df, column_name)
    if uni.empty and bi.empty:
        return None

    charts: List[alt.Chart] = []
    if not uni.empty:
        c1 = (
            alt.Chart(uni)
            .mark_bar()
            .encode(
                x=alt.X("count:Q", title="Count"),
                y=alt.Y("term:N", sort="-x", title=""),
                tooltip=["term:N", "count:Q"],
            )
            .properties(height=max(200, 18 * len(uni)))
        )
        charts.append(c1)
    if not bi.empty:
        c2 = (
            alt.Chart(bi)
            .mark_bar()
            .encode(
                x=alt.X("count:Q", title="Count"),
                y=alt.Y("term:N", sort="-x", title=""),
                tooltip=["term:N", "count:Q"],
            )
            .properties(height=max(200, 18 * len(bi)))
        )
        charts.append(c2)

    return alt.vconcat(*charts)


def main() -> None:
    st.title("Program Manager - Fortnightly School Audit Dashboard")
    st.caption("Analyze responses from the Excel sheet 'Form Responses 1'.")

    # Load default file and sheet; no filters or inputs
    path = DEFAULT_XLSX
    sheet = DEFAULT_SHEET

    try:
        df = load_excel(path, sheet_name=sheet)
    except Exception as e:
        st.error(str(e))
        st.stop()

    # Coerce potential timestamp columns (safe)
    for col in df.columns:
        df[col] = coerce_datetime(df[col]) if df[col].dtype == object else df[col]

    filtered_df = df

    if filtered_df.empty:
        st.warning("No data available.")
        st.stop()

    # Performance of Classroom Indicators (vertical stacked)
    st.subheader("Performance of Classroom Indicators")
    classroom_cols = find_classroom_indicator_columns(filtered_df)
    if classroom_cols:
        chart = build_stacked_chart_vertical(filtered_df, classroom_cols)
        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("No 'Classroom Indicators – Rate your observation [...]' columns found.")

    # Performance of Corridors and Open Spaces (vertical stacked)
    st.subheader("Performance of Corridors and Open Spaces")
    corridor_cols = find_corridors_columns(filtered_df)
    if corridor_cols:
        chart2 = build_stacked_chart_vertical(filtered_df, corridor_cols)
        st.altair_chart(chart2, use_container_width=True)
    else:
        st.info("No 'Corridors and Open Spaces – Rate your observation [...]' columns found.")

    # Performance of Teacher Update
    st.subheader("Performance of Teacher Update")
    teacher_cols = find_teacher_update_columns(filtered_df)
    if teacher_cols:
        chart3 = build_stacked_chart_vertical(filtered_df, teacher_cols)
        st.altair_chart(chart3, use_container_width=True)
    else:
        st.info("No 'Teacher Update – Rate your observation [...]' columns found.")

    # Student Interaction performance
    st.subheader("Student Interaction performance")
    interaction_cols = find_student_interaction_columns(filtered_df)
    if interaction_cols:
        chart4 = build_stacked_chart_vertical(filtered_df, interaction_cols)
        st.altair_chart(chart4, use_container_width=True)
    else:
        st.info("No 'Interaction with Students – Rate your observation [...]' columns found.")

    # Performance of Curriculum-Based Activities
    st.subheader("Performance of Curriculum-Based Activities")
    curriculum_cols = find_curriculum_activity_columns(filtered_df)
    if curriculum_cols:
        chart5 = build_stacked_chart_vertical(filtered_df, curriculum_cols)
        st.altair_chart(chart5, use_container_width=True)
    else:
        st.info("No 'Manual-based Curriculum Activities – Rate your observation [...]' columns found.")

    # Optional raw data preview
    with st.expander("Preview data"):
        st.dataframe(filtered_df.head(200))


if __name__ == "__main__":
    main()
