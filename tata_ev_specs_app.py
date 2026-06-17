"""
Tata.ev Spec & Feature Extractor — Streamlit app
--------------------------------------------------
Pulls the embedded product spec/feature JSON out of a Tata.ev model
"specifications" page (a raw HTML file hosted on GitHub, the live
ev.tatamotors.com page, or a manually uploaded .html file) and renders
every spec/feature category as its own table, stacked one below another.

Run with:
    pip install streamlit pandas requests beautifulsoup4 openpyxl lxml
    streamlit run tata_ev_specs_app.py
"""

import html
import io
import json
import re

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="Tata.ev Spec Extractor", layout="wide")

# Replace with the raw URL of the HTML file you've committed to GitHub,
# e.g. https://raw.githubusercontent.com/<user>/<repo>/main/tiago-ev-specs.html
DEFAULT_GITHUB_RAW_URL = "https://raw.githubusercontent.com/<user>/<repo>/main/tiago-ev-specifications.html"
LIVE_URL = "https://ev.tatamotors.com/tiago/ev/specifications.html"

st.title("Tata.ev — Specifications & Features Extractor")

with st.sidebar:
    st.header("Data source")
    source_type = st.radio(
        "Fetch HTML from:",
        ["Paste HTML or JSON", "GitHub raw URL", "Live Tata.ev URL", "Upload HTML file"],
    )

    url = None
    uploaded = None
    pasted_text = None
    if source_type == "Paste HTML or JSON":
        pasted_text = st.text_area(
            "Paste the page's HTML (or just the productspecjson value) here",
            height=200,
            placeholder='Paste full HTML source, or just {"results": {"variantSpecFeature": [...]}} ...',
        )
    elif source_type == "GitHub raw URL":
        url = st.text_input("GitHub raw URL", value=DEFAULT_GITHUB_RAW_URL)
    elif source_type == "Live Tata.ev URL":
        url = st.text_input("Page URL", value=LIVE_URL)
    else:
        uploaded = st.file_uploader("Upload saved .html file", type=["html", "htm"])

    diff_only = st.checkbox("Show only rows that differ across variants", value=False)
    fetch_btn = st.button("Fetch & Parse", type="primary")


@st.cache_data(show_spinner=False)
def fetch_html(source_url: str) -> str:
    resp = requests.get(source_url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.text


def extract_spec_json(raw_text: str) -> dict:
    """Locate the data-productspecjson attribute and parse it as JSON.
    Also accepts a plain JSON blob (e.g. pasted directly, no HTML wrapper)."""
    stripped = raw_text.strip()
    if stripped.startswith("{"):
        return json.loads(stripped)

    soup = BeautifulSoup(raw_text, "html.parser")
    container = soup.find(attrs={"data-productspecjson": True})
    if container is not None:
        raw_json = container["data-productspecjson"]  # bs4 already unescapes entities
    else:
        match = re.search(
            r'data-productspecjson="(.*?)"\s+data-productRecommendation',
            raw_text,
            re.DOTALL,
        )
        if not match:
            raise ValueError("Could not find the productspecjson data attribute on this page.")
        raw_json = html.unescape(match.group(1))
    return json.loads(raw_json)


def build_tables(spec_json: dict):
    """Returns ({category_title: DataFrame}, [variant_labels])."""
    variants = spec_json["results"]["variantSpecFeature"]
    variant_labels = [v["variantLabel"] for v in variants]
    tables = {}

    price_row = {"Feature": "Starting Price"}
    for v in variants:
        price_row[v["variantLabel"]] = v.get("startingPrice", "-")
    tables["Overview"] = pd.DataFrame([price_row]).set_index("Feature")

    def collect(items_key, type_key, title_key, list_key, label_key, value_key):
        categories = {}
        for v in variants:
            for block in v.get(items_key, []):
                title = block[title_key]
                categories.setdefault(title, {})
                for item in block.get(list_key, []):
                    label = item.get(label_key, "")
                    value = item.get(value_key) or "-"
                    categories[title].setdefault(label, {})[v["variantLabel"]] = value
        for title, rows in categories.items():
            df = pd.DataFrame(rows).T.reindex(columns=variant_labels).fillna("-")
            tables[title] = df

    collect("productSpecifications", "specType", "specTypeTitle", "specList", "specLabel", "specValue")
    collect("productFeatures", "featureType", "featureTypeTitle", "featureList", "featureLabel", "featureValue")

    return tables, variant_labels


def filter_differences(df: pd.DataFrame) -> pd.DataFrame:
    return df[df.nunique(axis=1) > 1]


if fetch_btn or "tables" in st.session_state:
    try:
        if pasted_text:
            raw_html = pasted_text
        elif uploaded is not None:
            raw_html = uploaded.read().decode("utf-8", errors="ignore")
        elif url:
            raw_html = fetch_html(url)
        else:
            st.warning("Paste data, provide a URL, or upload an HTML file, then click Fetch & Parse.")
            st.stop()

        spec_json = extract_spec_json(raw_html)
        tables, variant_labels = build_tables(spec_json)
        st.session_state["tables"] = tables
        st.session_state["variant_labels"] = variant_labels
    except Exception as e:
        st.error(f"Failed to extract data: {e}")
        st.stop()

if "tables" in st.session_state:
    tables = st.session_state["tables"]
    variant_labels = st.session_state["variant_labels"]

    st.success(f"Loaded {len(variant_labels)} variants: {', '.join(variant_labels)}")
    st.divider()

    # All category tables stacked vertically (no tabs)
    for title, df in tables.items():
        st.subheader(title)
        show_df = filter_differences(df) if diff_only else df
        if show_df.empty:
            st.caption("No differences across variants.")
        else:
            st.dataframe(show_df, use_container_width=True)
        st.divider()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for title, df in tables.items():
            df.to_excel(writer, sheet_name=title[:31])
    st.download_button(
        "Download all tables as Excel",
        data=buf.getvalue(),
        file_name="tata_ev_specifications.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Set a data source in the sidebar and click **Fetch & Parse** to begin.")
