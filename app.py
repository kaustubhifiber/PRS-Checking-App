import streamlit as st
import pandas as pd
import urllib.parse
import os
import time
import pythoncom

from datetime import datetime
import win32com.client as win32


# ============================================================
# FILE PATHS
# ============================================================

MASTER_FILE = "Master Sheet Team-B.xlsx"

PAYMENT_TRACKER = (
    r"C:\Users\DELL\OneDrive - Innovative Fiber Solutions Pvt Ltd"
    r"\Payment Tracker.xlsx"
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="PRS Search",
    page_icon="🔎",
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
# GET RUNNING EXCEL
# ============================================================

def get_open_excel():

    try:

        excel = win32.GetActiveObject(
            "Excel.Application"
        )

        return excel

    except Exception:

        return None


# ============================================================
# FIND PAYMENT TRACKER IN OPEN EXCEL
# ============================================================

def find_open_payment_tracker(excel):

    if excel is None:
        return None

    target_path = os.path.abspath(
        PAYMENT_TRACKER
    ).lower()

    for workbook in excel.Workbooks:

        try:

            workbook_path = os.path.abspath(
                workbook.FullName
            ).lower()

            if workbook_path == target_path:

                return workbook

        except Exception:

            continue

    return None


# ============================================================
# GET EXISTING / OPEN PAYMENT TRACKER
# ============================================================

def get_payment_tracker():

    # --------------------------------------------------------
    # Check whether Excel is already running
    # --------------------------------------------------------

    excel = get_open_excel()

    # --------------------------------------------------------
    # If workbook is already open
    # --------------------------------------------------------

    if excel is not None:

        workbook = find_open_payment_tracker(
            excel
        )

        if workbook is not None:

            return (
                excel,
                workbook,
                False
            )

    # --------------------------------------------------------
    # Excel not running OR workbook not open
    # --------------------------------------------------------

    excel = win32.Dispatch(
        "Excel.Application"
    )

    excel.Visible = True

    workbook = excel.Workbooks.Open(
        os.path.abspath(
            PAYMENT_TRACKER
        )
    )

    return (
        excel,
        workbook,
        True
    )


# ============================================================
# FIND TRACKER SHEET + HEADERS
# ============================================================

def find_tracker_sheet(workbook):

    required_headers = [
        "Date",
        "Project Site No.",
        "Project Name",
        "Installation Address (Street)",
        "Site Name",
        "Operator: Account Name",
        "Payment Status",
        "Accounts Remark"
    ]

    # --------------------------------------------------------
    # Check every worksheet
    # --------------------------------------------------------

    for sheet in workbook.Worksheets:

        used_range = sheet.UsedRange

        first_row = used_range.Row

        last_row = (
            used_range.Row
            + used_range.Rows.Count
            - 1
        )

        first_col = used_range.Column

        last_col = (
            used_range.Column
            + used_range.Columns.Count
            - 1
        )

        # Search first 30 rows for headers
        scan_last_row = min(
            last_row,
            first_row + 30
        )

        for row_number in range(
            first_row,
            scan_last_row + 1
        ):

            headers = {}

            for col_number in range(
                first_col,
                last_col + 1
            ):

                value = sheet.Cells(
                    row_number,
                    col_number
                ).Value

                if value is not None:

                    header = str(
                        value
                    ).strip()

                    headers[header] = (
                        col_number
                    )

            if all(
                header in headers
                for header in required_headers
            ):

                return (
                    sheet,
                    row_number,
                    headers
                )

    raise ValueError(
        "Could not find the Payment Tracker sheet "
        "with the required headers."
    )


# ============================================================
# ENSURE AGEING COLUMN
# ============================================================

def ensure_ageing_column(
    sheet,
    header_row,
    header_map
):

    # --------------------------------------------------------
    # Ageing already exists
    # --------------------------------------------------------

    if "Ageing" in header_map:

        return header_map["Ageing"]

    # --------------------------------------------------------
    # Ageing doesn't exist
    # Create it after last header
    # --------------------------------------------------------

    last_column = max(
        header_map.values()
    )

    ageing_column = (
        last_column + 1
    )

    sheet.Cells(
        header_row,
        ageing_column
    ).Value = "Ageing"

    return ageing_column


# ============================================================
# FIND LAST DATA ROW
# ============================================================

def find_next_row(
    sheet,
    header_row,
    header_map
):

    # --------------------------------------------------------
    # Use important tracker columns to find last real row
    # --------------------------------------------------------

    columns_to_check = [
        header_map["Date"],
        header_map["Project Site No."],
        header_map["Project Name"],
        header_map["Site Name"]
    ]

    last_data_row = header_row

    for column_number in columns_to_check:

        try:

            last_cell = sheet.Cells(
                sheet.Rows.Count,
                column_number
            )

            last_row = last_cell.End(
                -4162
            ).Row
            # -4162 = xlUp

            if last_row > last_data_row:

                last_data_row = last_row

        except Exception:

            continue

    # --------------------------------------------------------
    # ALWAYS NEXT ROW
    # --------------------------------------------------------

    next_row = (
        last_data_row + 1
    )

    if next_row <= header_row:

        next_row = (
            header_row + 1
        )

    return next_row


# ============================================================
# COPY PREVIOUS ROW FORMATTING
# ============================================================

def copy_previous_row_format(
    sheet,
    header_row,
    next_row,
    header_map,
    ageing_column
):

    # No previous data row
    if next_row <= header_row + 1:
        return

    previous_row = (
        next_row - 1
    )

    last_column = max(
        max(header_map.values()),
        ageing_column
    )

    try:

        source_range = sheet.Range(
            sheet.Cells(
                previous_row,
                1
            ),
            sheet.Cells(
                previous_row,
                last_column
            )
        )

        target_range = sheet.Range(
            sheet.Cells(
                next_row,
                1
            ),
            sheet.Cells(
                next_row,
                last_column
            )
        )

        source_range.Copy(
            target_range
        )

    except Exception:

        pass


# ============================================================
# EXTEND EXCEL TABLE
# ============================================================

def extend_excel_table(
    sheet,
    header_row,
    next_row
):

    try:

        tables = sheet.ListObjects

        for i in range(
            1,
            tables.Count + 1
        ):

            table = tables.Item(i)

            header_range = (
                table.HeaderRowRange
            )

            table_header_row = (
                header_range.Row
            )

            if (
                table_header_row
                == header_row
            ):

                current_range = (
                    table.Range
                )

                first_row = (
                    current_range.Row
                )

                first_col = (
                    current_range.Column
                )

                total_columns = (
                    current_range.Columns.Count
                )

                new_range = sheet.Range(
                    sheet.Cells(
                        first_row,
                        first_col
                    ),
                    sheet.Cells(
                        next_row,
                        first_col
                        + total_columns
                        - 1
                    )
                )

                table.Resize(
                    new_range
                )

                break

    except Exception:

        # Table resize should not stop
        # the actual Excel update.
        pass


# ============================================================
# APPEND TO PAYMENT TRACKER
# ============================================================

def append_to_payment_tracker(
    project_site_no,
    project_name,
    installation_address,
    site_name,
    operator_name,
    ageing
):

    # --------------------------------------------------------
    # CHECK FILE
    # --------------------------------------------------------

    if not os.path.exists(
        PAYMENT_TRACKER
    ):

        raise FileNotFoundError(
            "Payment Tracker.xlsx not found.\n\n"
            f"Expected location:\n"
            f"{PAYMENT_TRACKER}"
        )


    # --------------------------------------------------------
    # INITIALIZE COM
    # --------------------------------------------------------

    pythoncom.CoInitialize()


    excel = None
    workbook = None
    opened_by_app = False


    try:

        # ====================================================
        # GET / OPEN EXCEL
        # ====================================================

        (
            excel,
            workbook,
            opened_by_app
        ) = get_payment_tracker()


        # ====================================================
        # FIND TRACKER SHEET
        # ====================================================

        (
            sheet,
            header_row,
            header_map
        ) = find_tracker_sheet(
            workbook
        )


        # ====================================================
        # AGEING COLUMN
        # ====================================================

        ageing_column = ensure_ageing_column(
            sheet,
            header_row,
            header_map
        )


        # ====================================================
        # FIND NEXT NEW ROW
        # ====================================================

        next_row = find_next_row(
            sheet,
            header_row,
            header_map
        )


        # ====================================================
        # COPY FORMATTING
        # ====================================================

        copy_previous_row_format(
            sheet,
            header_row,
            next_row,
            header_map,
            ageing_column
        )


        # ====================================================
        # TODAY
        # ====================================================

        today = datetime.now().strftime(
            "%d-%m-%Y"
        )


        # ====================================================
        # WRITE DATE
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Date"]
        ).Value = today


        # ====================================================
        # WRITE PRS
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Project Site No."]
        ).Value = project_site_no


        # ====================================================
        # WRITE PROJECT NAME
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Project Name"]
        ).Value = project_name


        # ====================================================
        # WRITE ADDRESS
        # ====================================================

        sheet.Cells(
            next_row,
            header_map[
                "Installation Address (Street)"
            ]
        ).Value = installation_address


        # ====================================================
        # WRITE SITE NAME
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Site Name"]
        ).Value = site_name


        # ====================================================
        # WRITE OPERATOR
        # ====================================================

        sheet.Cells(
            next_row,
            header_map[
                "Operator: Account Name"
            ]
        ).Value = operator_name


        # ====================================================
        # PAYMENT STATUS
        # ALWAYS PENDING
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Payment Status"]
        ).Value = "Pending"


        # ====================================================
        # ACCOUNTS REMARK
        # BLANK
        # ====================================================

        sheet.Cells(
            next_row,
            header_map["Accounts Remark"]
        ).Value = ""


        # ====================================================
        # AGEING
        # ====================================================

        sheet.Cells(
            next_row,
            ageing_column
        ).Value = ageing


        # ====================================================
        # EXTEND TABLE
        # ====================================================

        extend_excel_table(
            sheet,
            header_row,
            next_row
        )


        # ====================================================
        # SAVE WORKBOOK
        # ====================================================

        workbook.Save()


        # ====================================================
        # WAIT
        # ====================================================

        time.sleep(1)


        # ====================================================
        # VERIFY DATA
        # ====================================================

        saved_prs = sheet.Cells(
            next_row,
            header_map[
                "Project Site No."
            ]
        ).Value


        if (
            str(saved_prs).strip()
            != str(project_site_no).strip()
        ):

            raise RuntimeError(
                "Excel save verification failed. "
                "The row was not saved correctly."
            )


        # ====================================================
        # RETURN SUCCESS
        # ====================================================

        return {
            "date": today,
            "row": next_row,
            "sheet": sheet.Name,
            "path": PAYMENT_TRACKER
        }


    finally:

        # ====================================================
        # CLOSE ONLY IF APP OPENED EXCEL
        # ====================================================

        try:

            if (
                workbook is not None
                and opened_by_app
            ):

                workbook.Save()

                workbook.Close(
                    SaveChanges=True
                )

        except Exception:

            pass


        # ====================================================
        # QUIT ONLY IF APP OPENED EXCEL
        # ====================================================

        try:

            if (
                excel is not None
                and opened_by_app
            ):

                excel.Quit()

        except Exception:

            pass


        # ====================================================
        # COM CLEANUP
        # ====================================================

        try:

            pythoncom.CoUninitialize()

        except Exception:

            pass


# ============================================================
# SESSION STATE
# ============================================================

if (
    "search_results"
    not in st.session_state
):

    st.session_state.search_results = None


if (
    "search_text"
    not in st.session_state
):

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

        st.warning(
            "Please enter a PRS No."
        )

        st.session_state.search_results = None


    else:

        try:

            # =================================================
            # LOAD MASTER
            # =================================================

            df = load_excel()


            # =================================================
            # REQUIRED COLUMNS
            # =================================================

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
                    + ", ".join(
                        missing_columns
                    )
                )

                st.session_state.search_results = None


            else:

                # =============================================
                # NORMALIZE MASTER PRS
                # =============================================

                df["_PRS_SEARCH"] = (
                    df["Project Site No."]
                    .apply(normalize_prs)
                )


                # =============================================
                # NORMALIZE USER INPUT
                # =============================================

                search_value = normalize_prs(
                    prs_no
                )


                # =============================================
                # SEARCH
                # =============================================

                result = df[
                    df["_PRS_SEARCH"]
                    == search_value
                ].copy()


                # =============================================
                # SAVE RESULTS
                # =============================================

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
                f"Master Excel file not found:\n"
                f"{MASTER_FILE}"
            )


        except Exception as e:

            st.error(
                f"Search error:\n\n{e}"
            )


# ============================================================
# DISPLAY RESULTS
# ============================================================

if (
    st.session_state.search_results
    is not None
):

    result = (
        st.session_state.search_results
    )


    st.success(
        f"{len(result)} record(s) found"
    )


    # ========================================================
    # EACH RESULT
    # ========================================================

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
            row[
                "Installation Address (Street)"
            ]
        )

        site_name = safe_value(
            row["Site Name"]
        )

        operator_name = safe_value(
            row[
                "Operator: Account Name"
            ]
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


        # ====================================================
        # PRS
        # ====================================================

        st.markdown(
            '<div class="label">PRS No.</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{project_site_no}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # PROJECT NAME
        # ====================================================

        st.markdown(
            '<div class="label">Project Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{project_name}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # ADDRESS
        # ====================================================

        st.markdown(
            '<div class="label">Installation Address</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{installation_address}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # SITE NAME
        # ====================================================

        st.markdown(
            '<div class="label">Site Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{site_name}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # OPERATOR
        # ====================================================

        st.markdown(
            '<div class="label">Operator Name</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{operator_name}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # PAYMENT STATUS
        # ====================================================

        st.markdown(
            '<div class="label">Payment Status</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            f'<div class="value">{payment_status}</div>',
            unsafe_allow_html=True
        )


        # ====================================================
        # AGEING
        # ====================================================

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
            "*PRS PAYMENT UPDATE*\n\n"
            f"PRS No: {project_site_no}\n"
            f"Project Name: {project_name}\n"
            f"Site Name: {site_name}\n"
            f"Operator: {operator_name}\n"
            f"Payment Status: {payment_status}\n"
            f"Ageing: {ageing}\n\n"
            "Please check and update the payment status."
        )


        # ====================================================
        # WHATSAPP URL
        # ====================================================

        whatsapp_url = (
            "https://web.whatsapp.com/send"
            "?text="
            + urllib.parse.quote(
                whatsapp_message
            )
        )


        # ====================================================
        # BUTTONS
        # ====================================================

        col1, col2 = st.columns(2)


        # ====================================================
        # WHATSAPP BUTTON
        # ====================================================

        with col1:

            st.link_button(
                "Send on WhatsApp",
                whatsapp_url,
                use_container_width=True
            )


        # ====================================================
        # MAIL BUTTON
        # ====================================================

        with col2:

            mail_clicked = st.button(
                "Send on Mail",
                key=f"mail_button_{index}",
                use_container_width=True
            )


        # ====================================================
        # MAIL ACTION
        # ====================================================

        if mail_clicked:

            try:

                tracker_result = (
                    append_to_payment_tracker(
                        project_site_no=project_site_no,
                        project_name=project_name,
                        installation_address=installation_address,
                        site_name=site_name,
                        operator_name=operator_name,
                        ageing=ageing
                    )
                )


                st.success(
                    "✓ Payment Tracker updated successfully\n\n"
                    f"Sheet: {tracker_result['sheet']} | "
                    f"Row: {tracker_result['row']} | "
                    f"Date: {tracker_result['date']}"
                )


            except FileNotFoundError as e:

                st.error(
                    str(e)
                )


            except PermissionError:

                st.error(
                    "Excel file access denied. "
                    "Please check the Payment Tracker file."
                )


            except RuntimeError as e:

                st.error(
                    str(e)
                )


            except Exception as e:

                st.error(
                    "Payment Tracker update failed:\n\n"
                    f"{e}"
                )


        # ====================================================
        # CARD END
        # ====================================================

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )
