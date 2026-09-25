import streamlit as st
import pandas as pd
import urllib.parse


# ============================================================
# CONFIG
# ============================================================

MASTER_FILE = "Master Sheet Sapphire Foods.xlsx"


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="PRS Search",
    page_icon="https://share.google/0lWz2leNhUYPDEkVl",
    layout="centered"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
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
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #ddd;
        margin-bottom: 18px;
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

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="title">PRS Search</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Search project site details</div>',
    unsafe_allow_html=True
)


# ============================================================
# LOAD MASTER EXCEL
# ============================================================

@st.cache_data
def load_excel():

    df = pd.read_excel(
        MASTER_FILE,
        sheet_name="Main Sheet"
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# NORMALIZE PRS
# ============================================================

def normalize_prs(value):

    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    value = value.replace("prs-", "")
    value = value.replace("prs", "")
    value = value.replace(" ", "")

    return value


# ============================================================
# SAFE VALUE
# ============================================================

def safe_value(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


# ============================================================
# SESSION STATE
# ============================================================

if "search_results" not in st.session_state:
    st.session_state.search_results = None

if "search_text" not in st.session_state:
    st.session_state.search_text = ""


# ============================================================
# SEARCH INPUT
# ============================================================

prs_no = st.text_input(
    "Enter PRS No.",
    value=st.session_state.search_text,
    placeholder="Example: 00400 or PrS-00400"
)


# ============================================================
# SEARCH BUTTON
# ============================================================

search_button = st.button(
    "Search",
    use_container_width=True
)


# ============================================================
# SEARCH LOGIC
# ============================================================

if search_button:

    if not prs_no.strip():

        st.warning("Please enter a PRS No.")
        st.session_state.search_results = None

    else:

        try:

            df = load_excel()

            required_columns = [
                "Project Site No.",
                "Project Name",
                "Installation Address (Street)",
                "Site Name",
                "Operator: Account Name",
                "Payment Status",
                "Ageing"
            ]

            missing_columns = [
                column
                for column in required_columns
                if column not in df.columns
            ]

            if missing_columns:

                st.error(
                    "Missing columns in Master Sheet:\n"
                    + ", ".join(missing_columns)
                )

                st.session_state.search_results = None

            else:

                df["_PRS_SEARCH"] = (
                    df["Project Site No."]
                    .apply(normalize_prs)
                )

                search_value = normalize_prs(prs_no)

                result = df[
                    df["_PRS_SEARCH"] == search_value
                ].copy()

                if result.empty:

                    st.session_state.search_results = None

                    st.error(
                        f"No record found for PRS No. {prs_no}"
                    )

                else:

                    st.session_state.search_results = result
                    st.session_state.search_text = prs_no

        except FileNotFoundError:

            st.error(
                f"Master Excel file not found: {MASTER_FILE}"
            )

        except Exception as e:

            st.error(
                f"Search error:\n\n{e}"
            )


# ============================================================
# DISPLAY RESULTS
# ============================================================

if st.session_state.search_results is not None:

    result = st.session_state.search_results

    st.success(
        f"{len(result)} record(s) found"
    )

    for index, (_, row) in enumerate(
        result.iterrows()
    ):

        # ----------------------------------------------------
        # GET VALUES
        # ----------------------------------------------------

        project_site_no = safe_value(
            row["Project Site No."]
        )

        project_name = safe_value(
            row["Project Name"]
        )

        installation_address = safe_value(
            row["Installation Address (Street)"]
        )

        site_name = safe_value(
            row["Site Name"]
        )

        operator_name = safe_value(
            row["Operator: Account Name"]
        )

        payment_status = safe_value(
            row["Payment Status"]
        )

        ageing = safe_value(
            row["Ageing"]
        )


        # ====================================================
        # CARD
        # ====================================================

        st.markdown(
            '<div class="result-card">',
            unsafe_allow_html=True
        )


        # ----------------------------------------------------
        # DETAILS
        # ----------------------------------------------------

        st.markdown(
            '<div class="label">PRS No.</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{project_site_no}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Project Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{project_name}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Installation Address</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{installation_address}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Site Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{site_name}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Operator Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{operator_name}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Payment Status</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{payment_status}</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div class="label">Ageing</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{ageing}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # WHATSAPP MESSAGE
        # ====================================================

        whatsapp_message = (
            f"{project_site_no}\n"
            f"{project_name}\n"
            f"{site_name}\n"
            f"{operator_name}\n"
            f"{ageing}\n\n"
            "@7208930912Please check and update the payment status."
        )


        # ====================================================
        # WHATSAPP URL
        # ====================================================

        whatsapp_url = (
            "https://web.whatsapp.com/send"
            "?text="
            + urllib.parse.quote(whatsapp_message)
        )


        # ====================================================
        # WHATSAPP BUTTON ONLY
        # ====================================================

        st.link_button(
            "Send on WhatsApp",
            whatsapp_url,
            use_container_width=True
        )


        # ====================================================
        # CARD END
        # ====================================================

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )
