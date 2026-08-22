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
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="WhatsApp Message Counter",
    page_icon="💬",
    layout="centered"
)


# ============================================================
# HEADING
# ============================================================

st.title("WhatsApp Message Counter")


# ============================================================
# PARSE WHATSAPP CHAT
# ============================================================

def parse_whatsapp_chat(uploaded_file):

    temp_dir = tempfile.mkdtemp()

    zip_path = os.path.join(
        temp_dir,
        "whatsapp_chat.zip"
    )

    # Save ZIP
    with open(zip_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extract ZIP
    with zipfile.ZipFile(
        zip_path,
        "r"
    ) as zip_ref:

        zip_ref.extractall(temp_dir)

    # Find TXT files
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
            "No WhatsApp chat text file found."
        )

    # Prefer WhatsApp's normal *_chat.txt file
    preferred_files = [
        f for f in txt_files
        if "_chat.txt" in os.path.basename(f).lower()
    ]

    if preferred_files:
        chat_file = preferred_files[0]
    else:
        chat_file = txt_files[0]

    # Read file
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

    # WhatsApp message pattern
    message_pattern = re.compile(
        r"^(\d{1,2}/\d{1,2}/\d{4}),\s+"
        r"(\d{1,2}:\d{2})\s+-\s+"
        r"(.*)$"
    )

    messages = []

    current_message = None

    for raw_line in lines:

        line = raw_line.rstrip(
            "\n\r"
        )

        match = message_pattern.match(
            line
        )

        if match:

            # Save previous message
            if current_message is not None:

                messages.append(
                    current_message
                )

            date_str = match.group(1)
            time_str = match.group(2)
            content = match.group(3)

            try:

                message_datetime = datetime.strptime(
                    f"{date_str}, {time_str}",
                    "%d/%m/%Y, %H:%M"
                )

            except ValueError:

                current_message = None
                continue

            # Separate sender and message
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

        else:

            # Multiline message
            if current_message is not None:

                current_message["message"] += (
                    "\n" + line
                )

    # Save final message
    if current_message is not None:

        messages.append(
            current_message
        )

    if not messages:

        raise ValueError(
            "No messages found."
        )

    df = pd.DataFrame(
        messages
    )

    # Remove WhatsApp system messages
    df = df[
        df["sender"].notna()
    ].copy()

    if df.empty:

        raise ValueError(
            "No user messages found."
        )

    return df


# ============================================================
# UPLOAD ZIP
# ============================================================

uploaded_file = st.file_uploader(
    "Upload ZIP file",
    type=["zip"],
    label_visibility="visible"
)


# ============================================================
# AFTER ZIP UPLOAD
# ============================================================

if uploaded_file is not None:

    # Parse chat
    try:

        df = parse_whatsapp_chat(
            uploaded_file
        )

    except Exception:

        st.error(
            "Unable to read the WhatsApp chat file."
        )

        st.stop()

    # Get available dates
    chat_min_date = df["date"].min()
    chat_max_date = df["date"].max()

    # ========================================================
    # DATE RANGE
    # ========================================================

    st.subheader("Select Date Range")

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
    # PROCESS
    # ========================================================

    if st.button(
        "Process",
        type="primary",
        use_container_width=True
    ):

        if start_date > end_date:

            st.error(
                "Start Date cannot be after End Date."
            )

            st.stop()

        # Filter messages
        filtered_df = df[
            (df["date"] >= start_date) &
            (df["date"] <= end_date)
        ].copy()

        if filtered_df.empty:

            st.error(
                "No messages found for the selected date range."
            )

            st.stop()

        # ====================================================
        # SHEET 1
        # NUMBER × DATE FREQUENCY MATRIX
        # ====================================================

        result = pd.crosstab(
            filtered_df["sender"],
            filtered_df["date"]
        )

        result.index.name = "Number"

        result.columns = pd.to_datetime(
            result.columns
        )

        result = result.sort_index(
            axis=1
        )

        result = result.sort_index()

        result = result.fillna(
            0
        ).astype(int)

        result = result.reset_index()

        # ====================================================
        # SHEET 2
        # SUMMARY BY NUMBER
        # ====================================================

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

        # ====================================================
        # SHEET 3
        # DAILY SUMMARY
        # ====================================================

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

        # ====================================================
        # CREATE EXCEL
        # ====================================================

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

        output.seek(0)

        # ====================================================
        # FORMAT EXCEL
        # ====================================================

        wb = load_workbook(
            output
        )

        for ws in wb.worksheets:

            # Freeze header
            ws.freeze_panes = "A2"

            # Bold headers
            for cell in ws[1]:

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

            # Auto-width
            for column in ws.columns:

                max_length = 0

                column_letter = (
                    column[0].column_letter
                )

                for cell in column:

                    try:

                        length = len(
                            str(cell.value)
                        )

                        max_length = max(
                            max_length,
                            length
                        )

                    except Exception:

                        pass

                ws.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 2,
                    40
                )

        # Format dates in first sheet
        ws = wb["Date x Number"]

        for cell in ws[1][1:]:

            cell.number_format = "dd/mm/yyyy"

        # Format dates in daily summary
        ws = wb["Daily Summary"]

        for cell in ws["A"][1:]:

            cell.number_format = "dd/mm/yyyy"

        # Save final Excel
        final_output = BytesIO()

        wb.save(
            final_output
        )

        final_output.seek(0)

        # ====================================================
        # DOWNLOAD
        # ====================================================

        filename = (
            "WhatsApp_Message_Count_"
            f"{start_date.strftime('%Y%m%d')}_"
            f"to_"
            f"{end_date.strftime('%Y%m%d')}.xlsx"
        )

        st.download_button(
            label="Download Excel",
            data=final_output,
            file_name=filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )
