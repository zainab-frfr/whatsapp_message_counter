import streamlit as st
import zipfile
import re
import os
import tempfile
from datetime import datetime
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="WhatsApp Message Counter",
    page_icon="💬",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("💬 WhatsApp Message Counter")

st.write(
    "Upload a WhatsApp chat export ZIP file, select a date range, "
    "and generate an Excel file showing message frequency by "
    "number and date."
)


# ============================================================
# PARSE WHATSAPP CHAT
# ============================================================

def parse_whatsapp_chat(uploaded_file):

    # --------------------------------------------------------
    # Create temporary folder
    # --------------------------------------------------------

    temp_dir = tempfile.mkdtemp()

    zip_path = os.path.join(
        temp_dir,
        "whatsapp_chat.zip"
    )


    # --------------------------------------------------------
    # Save uploaded ZIP
    # --------------------------------------------------------

    with open(zip_path, "wb") as f:

        f.write(
            uploaded_file.getbuffer()
        )


    # --------------------------------------------------------
    # Extract ZIP
    # --------------------------------------------------------

    with zipfile.ZipFile(
        zip_path,
        "r"
    ) as zip_ref:

        zip_ref.extractall(
            temp_dir
        )


    # --------------------------------------------------------
    # Find TXT file
    # --------------------------------------------------------

    txt_files = []

    for root, dirs, files in os.walk(temp_dir):

        for filename in files:

            if filename.lower().endswith(".txt"):

                txt_files.append(
                    os.path.join(
                        root,
                        filename
                    )
                )


    if not txt_files:

        raise ValueError(
            "No .txt file was found inside the ZIP file."
        )


    # --------------------------------------------------------
    # If multiple TXT files exist
    # --------------------------------------------------------

    if len(txt_files) > 1:

        # Ignore WhatsApp's common metadata files if present

        preferred_files = [
            f for f in txt_files
            if "_chat.txt" in os.path.basename(f).lower()
        ]

        if preferred_files:

            chat_file = preferred_files[0]

        else:

            chat_file = txt_files[0]

    else:

        chat_file = txt_files[0]


    # --------------------------------------------------------
    # Read TXT file
    # --------------------------------------------------------

    try:

        with open(
            chat_file,
            "r",
            encoding="utf-8"
        ) as f:

            lines = f.readlines()

    except UnicodeDecodeError:

        with open(
            chat_file,
            "r",
            encoding="utf-8-sig"
        ) as f:

            lines = f.readlines()


    # --------------------------------------------------------
    # WhatsApp message pattern
    # --------------------------------------------------------

    message_pattern = re.compile(
        r"^(\d{1,2}/\d{1,2}/\d{4}),\s+"
        r"(\d{1,2}:\d{2})\s+-\s+"
        r"(.*)$"
    )


    # --------------------------------------------------------
    # Parse messages
    # --------------------------------------------------------

    messages = []

    current_message = None


    for raw_line in lines:

        line = raw_line.rstrip(
            "\n\r"
        )


        match = message_pattern.match(
            line
        )


        # ----------------------------------------------------
        # New message
        # ----------------------------------------------------

        if match:

            # Save previous message
            if current_message is not None:

                messages.append(
                    current_message
                )


            date_str = match.group(1)
            time_str = match.group(2)
            content = match.group(3)


            # Convert date/time
            try:

                message_datetime = datetime.strptime(
                    f"{date_str}, {time_str}",
                    "%d/%m/%Y, %H:%M"
                )

            except ValueError:

                current_message = None
                continue


            # ------------------------------------------------
            # Separate sender from message
            # ------------------------------------------------

            if ": " in content:

                sender, message_text = content.split(
                    ": ",
                    1
                )


                current_message = {

                    "datetime": message_datetime,

                    "date": message_datetime.date(),

                    "sender": sender.strip(),

                    "message": message_text

                }


            else:

                # WhatsApp system message

                current_message = {

                    "datetime": message_datetime,

                    "date": message_datetime.date(),

                    "sender": None,

                    "message": content

                }


        # ----------------------------------------------------
        # Multiline continuation
        # ----------------------------------------------------

        else:

            if current_message is not None:

                current_message["message"] += (
                    "\n" + line
                )


    # --------------------------------------------------------
    # Save final message
    # --------------------------------------------------------

    if current_message is not None:

        messages.append(
            current_message
        )


    if not messages:

        raise ValueError(
            "No WhatsApp messages could be identified."
        )


    # --------------------------------------------------------
    # Create dataframe
    # --------------------------------------------------------

    df = pd.DataFrame(
        messages
    )


    # --------------------------------------------------------
    # Remove WhatsApp system messages
    # --------------------------------------------------------

    df = df[
        df["sender"].notna()
    ].copy()


    if df.empty:

        raise ValueError(
            "No user messages were found in the chat."
        )


    return df


# ============================================================
# CREATE EXCEL FILE
# ============================================================

def create_excel(
    filtered_df,
    start_date,
    end_date
):

    # ========================================================
    # SHEET 1
    # NUMBER × DATE FREQUENCY MATRIX
    # ========================================================

    result = pd.crosstab(
        filtered_df["sender"],
        filtered_df["date"]
    )


    # Name index
    result.index.name = "Number"


    # Convert date columns to datetime
    result.columns = pd.to_datetime(
        result.columns
    )


    # Sort dates chronologically
    result = result.sort_index(
        axis=1
    )


    # Sort senders alphabetically
    result = result.sort_index()


    # Fill missing combinations with zero
    result = result.fillna(0)


    # Make message counts integers
    result = result.astype(int)


    # Put Number back as a column
    result = result.reset_index()


    # ========================================================
    # SHEET 2
    # SUMMARY BY NUMBER
    # ========================================================

    summary = (

        filtered_df

        .groupby(
            "sender"
        )

        .size()

        .reset_index(
            name="Message Count"
        )

    )


    summary = summary.rename(
        columns={
            "sender": "Number"
        }
    )


    summary = summary.sort_values(
        by="Message Count",
        ascending=False
    ).reset_index(
        drop=True
    )


    # ========================================================
    # SHEET 3
    # DAILY SUMMARY
    # ========================================================

    daily_summary = (

        filtered_df

        .groupby(
            "date"
        )

        .size()

        .reset_index(
            name="Message Count"
        )

    )


    daily_summary = daily_summary.rename(
        columns={
            "date": "Date"
        }
    )


    daily_summary["Date"] = pd.to_datetime(
        daily_summary["Date"]
    )


    daily_summary = daily_summary.sort_values(
        by="Date"
    )


    # ========================================================
    # CREATE EXCEL IN MEMORY
    # ========================================================

    output = BytesIO()


    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        result.to_excel(
            writer,
            sheet_name="Date x Number",
            index=False
        )


        summary.to_excel(
            writer,
            sheet_name="Summary by Number",
            index=False
        )


        daily_summary.to_excel(
            writer,
            sheet_name="Daily Summary",
            index=False
        )


    # ========================================================
    # FORMAT EXCEL
    # ========================================================

    output.seek(0)


    wb = load_workbook(
        output
    )


    # --------------------------------------------------------
    # General formatting
    # --------------------------------------------------------

    for ws in wb.worksheets:

        # Freeze header row
        ws.freeze_panes = "A2"


        # Bold headers
        for cell in ws[1]:

            cell.font = Font(
                bold=True
            )


        # Center headers
        for cell in ws[1]:

            cell.alignment = Alignment(
                horizontal="center"
            )


        # Auto-size columns
        for column in ws.columns:

            max_length = 0

            column_letter = (
                column[0].column_letter
            )


            for cell in column:

                try:

                    cell_length = len(
                        str(cell.value)
                    )

                    if cell_length > max_length:

                        max_length = cell_length

                except Exception:

                    pass


            ws.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                40
            )


    # --------------------------------------------------------
    # Format dates in first sheet
    # --------------------------------------------------------

    ws = wb[
        "Date x Number"
    ]


    for cell in ws[1][1:]:

        cell.number_format = "dd/mm/yyyy"


    # --------------------------------------------------------
    # Format dates in Daily Summary
    # --------------------------------------------------------

    ws = wb[
        "Daily Summary"
    ]


    for cell in ws["A"][1:]:

        cell.number_format = "dd/mm/yyyy"


    # --------------------------------------------------------
    # Save formatted workbook back to memory
    # --------------------------------------------------------

    final_output = BytesIO()


    wb.save(
        final_output
    )


    final_output.seek(0)


    return final_output


# ============================================================
# STREAMLIT APP
# ============================================================

uploaded_file = st.file_uploader(
    "Upload WhatsApp Chat ZIP",
    type=["zip"],
    help="Upload the ZIP file created when exporting the WhatsApp chat."
)


# ============================================================
# AFTER FILE UPLOAD
# ============================================================

if uploaded_file is not None:

    # --------------------------------------------------------
    # Parse chat
    # --------------------------------------------------------

    with st.spinner(
        "Reading WhatsApp chat..."
    ):

        try:

            df = parse_whatsapp_chat(
                uploaded_file
            )

        except Exception as e:

            st.error(
                f"Could not process the chat: {e}"
            )

            st.stop()


    # --------------------------------------------------------
    # Determine chat date range
    # --------------------------------------------------------

    chat_min_date = df[
        "date"
    ].min()


    chat_max_date = df[
        "date"
    ].max()


    # --------------------------------------------------------
    # Success message
    # --------------------------------------------------------

    st.success(
        f"Chat loaded successfully — "
        f"{len(df):,} messages found."
    )


    st.caption(
        f"Available chat dates: "
        f"{chat_min_date.strftime('%d/%m/%Y')} "
        f"to "
        f"{chat_max_date.strftime('%d/%m/%Y')}"
    )


    # ========================================================
    # DATE RANGE
    # ========================================================

    st.subheader(
        "Select Date Range"
    )


    col1, col2 = st.columns(2)


    with col1:

        start_date = st.date_input(
            "Start Date",
            value=chat_min_date,
            min_value=chat_min_date,
            max_value=chat_max_date
        )


    with col2:

        end_date = st.date_input(
            "End Date",
            value=chat_max_date,
            min_value=chat_min_date,
            max_value=chat_max_date
        )


    # ========================================================
    # VALIDATE DATES
    # ========================================================

    if start_date > end_date:

        st.error(
            "Start Date cannot be after End Date."
        )

        st.stop()


    # ========================================================
    # PROCESS BUTTON
    # ========================================================

    if st.button(
        "📊 Process Chat",
        type="primary",
        use_container_width=True
    ):

        # ----------------------------------------------------
        # Filter selected dates
        # ----------------------------------------------------

        filtered_df = df[
            (df["date"] >= start_date) &
            (df["date"] <= end_date)
        ].copy()


        # ----------------------------------------------------
        # Check if messages exist
        # ----------------------------------------------------

        if filtered_df.empty:

            st.warning(
                "No messages were found within "
                "the selected date range."
            )

            st.stop()


        # ====================================================
        # RESULTS
        # ====================================================

        st.subheader(
            "Results"
        )


        metric1, metric2, metric3 = st.columns(3)


        with metric1:

            st.metric(
                "Messages",
                f"{len(filtered_df):,}"
            )


        with metric2:

            st.metric(
                "Senders",
                f"{filtered_df['sender'].nunique():,}"
            )


        with metric3:

            st.metric(
                "Days",
                f"{filtered_df['date'].nunique():,}"
            )


        # ====================================================
        # PREVIEW TABLE
        # ====================================================

        st.subheader(
            "Message Frequency by Date & Number"
        )


        preview = pd.crosstab(
            filtered_df["sender"],
            filtered_df["date"]
        )


        preview.index.name = "Number"


        preview.columns = [
            d.strftime("%d/%m/%Y")
            for d in preview.columns
        ]


        preview = preview.fillna(
            0
        ).astype(int)


        preview = preview.sort_index()


        st.dataframe(
            preview,
            use_container_width=True
        )


        # ====================================================
        # CREATE EXCEL
        # ====================================================

        with st.spinner(
            "Creating Excel file..."
        ):

            excel_file = create_excel(
                filtered_df,
                start_date,
                end_date
            )


        # ====================================================
        # FILE NAME
        # ====================================================

        filename = (
            "WhatsApp_Message_Count_"
            f"{start_date.strftime('%Y%m%d')}_"
            f"to_"
            f"{end_date.strftime('%Y%m%d')}"
            ".xlsx"
        )


        # ====================================================
        # DOWNLOAD
        # ====================================================

        st.download_button(

            label="⬇️ Download Excel",

            data=excel_file,

            file_name=filename,

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True

        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "WhatsApp Message Counter"
)