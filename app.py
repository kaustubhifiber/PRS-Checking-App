import streamlit as st
import pandas as pd


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PRS Search",
    page_icon="🔎",
    layout="centered"
)


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

.main {
    padding-top: 2rem;
}

.title {
    text-align: center;
    font-size: 32px;
    font-weight: 700;
    margin-bottom: 5px;
}

.subtitle {
    text-align: center;
    color: #666;
    margin-bottom: 30px;
}

.result-card {
    padding: 18px;
    border-radius: 12px;
    border: 1px solid #ddd;
    margin-bottom: 15px;
    background: white;
}

.label {
    font-size: 12px;
    color: #777;
    margin-bottom: 3px;
}

.value {
    font-size: 16px;
    font-weight: 600;
    margin-bottom: 12px;
    word-break: break-word;
}

@media (max-width: 600px) {

    .title {
        font-size: 26px;
    }

    .main {
        padding: 1rem;
    }

}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="title">PRS Search</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Search project site details</div>',
    unsafe_allow_html=True
)


# =========================================================
# LOAD EXCEL
# =========================================================

@st.cache_data
def load_excel():

    file_path = "Master Sheet Team-B.xlsx"

    df = pd.read_excel(
        file_path,
        sheet_name="Main Sheet"
    )

    df.columns = df.columns.astype(str).str.strip()

    return df


# =========================================================
# NORMALIZE PRS
# =========================================================

def normalize_prs(value):

    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    # Remove prs prefix
    value = value.replace("prs-", "")
    value = value.replace("prs", "")

    # Remove spaces
    value = value.replace(" ", "")

    return value


# =========================================================
# SEARCH INPUT
# =========================================================

prs_no = st.text_input(
    "Enter PRS No.",
    placeholder="Example: 00400 or PrS-00400"
)

search_button = st.button(
    "Search",
    use_container_width=True
)


# =========================================================
# SEARCH LOGIC
# =========================================================

if search_button:

    if not prs_no.strip():

        st.warning("Please enter a PRS No.")

    else:

        try:

            # Load Excel
            df = load_excel()

            # Required columns
            required_columns = [
                "Project Site No.",
                "Site Name",
                "Operator: Account Name",
                "Payment Status",
                "Ageing"
            ]

            missing_columns = [
                col
                for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:

                st.error(
                    "Missing columns: "
                    + ", ".join(missing_columns)
                )

                st.stop()

            # -------------------------------------------------
            # Normalize Excel PRS values
            # -------------------------------------------------

            df["_PRS_SEARCH"] = (
                df["Project Site No."]
                .apply(normalize_prs)
            )

            # Normalize user input
            search_value = normalize_prs(prs_no)

            # -------------------------------------------------
            # Find matching rows
            # -------------------------------------------------

            result = df[
                df["_PRS_SEARCH"] == search_value
            ]

            # =================================================
            # NO RESULT
            # =================================================

            if result.empty:

                st.error(
                    f"No record found for PRS No. {prs_no}"
                )

            # =================================================
            # RESULTS
            # =================================================

            else:

                st.success(
                    f"{len(result)} record(s) found"
                )

                for _, row in result.iterrows():

                    st.markdown(
                        '<div class="result-card">',
                        unsafe_allow_html=True
                    )

                    # -------------------------------------------------
                    # PRS NO
                    # -------------------------------------------------

                    st.markdown(
                        '<div class="label">PRS No.</div>',
                        unsafe_allow_html=True
                    )

                    prs_display = row["Project Site No."]

                    if pd.isna(prs_display):
                        prs_display = "-"

                    st.markdown(
                        f'<div class="value">{prs_display}</div>',
                        unsafe_allow_html=True
                    )

                    # -------------------------------------------------
                    # SITE NAME
                    # -------------------------------------------------

                    st.markdown(
                        '<div class="label">Site Name</div>',
                        unsafe_allow_html=True
                    )

                    site_name = row["Site Name"]

                    if pd.isna(site_name):
                        site_name = "-"

                    st.markdown(
                        f'<div class="value">{site_name}</div>',
                        unsafe_allow_html=True
                    )

                    # -------------------------------------------------
                    # OPERATOR NAME
                    # -------------------------------------------------

                    st.markdown(
                        '<div class="label">Operator Name</div>',
                        unsafe_allow_html=True
                    )

                    operator_name = row[
                        "Operator: Account Name"
                    ]

                    if pd.isna(operator_name):
                        operator_name = "-"

                    st.markdown(
                        f'<div class="value">{operator_name}</div>',
                        unsafe_allow_html=True
                    )

                    # -------------------------------------------------
                    # PAYMENT STATUS
                    # -------------------------------------------------

                    st.markdown(
                        '<div class="label">Payment Status</div>',
                        unsafe_allow_html=True
                    )

                    payment_status = row[
                        "Payment Status"
                    ]

                    if pd.isna(payment_status):
                        payment_status = "-"

                    st.markdown(
                        f'<div class="value">{payment_status}</div>',
                        unsafe_allow_html=True
                    )

                    # -------------------------------------------------
                    # AGEING
                    # -------------------------------------------------

                    st.markdown(
                        '<div class="label">Ageing</div>',
                        unsafe_allow_html=True
                    )

                    ageing = row["Ageing"]

                    if pd.isna(ageing):
                        ageing = "-"

                    st.markdown(
                        f'<div class="value">{ageing}</div>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        '</div>',
                        unsafe_allow_html=True
                    )

        except FileNotFoundError:

            st.error(
                "Master Sheet Team-B.xlsx not found. "
                "Keep the Excel file in the same folder as app.py."
            )

        except Exception as e:

            st.error(
                f"Error: {str(e)}"
            )
