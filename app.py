# File: app.py

import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------------
# 1) Processing Logic
# ------------------------------
def process_profit_summary(df, days):
    """
    Cleans and formats the profit summary data for a given # of days,
    using 'Items Cost' (Column X) to calculate ROI.
    """
    selected_columns = [
        "SKU",
        "Product Name",
        "ASIN",
        "Qty Sold",
        "Total Fees",
        "Items Cost",            # Use Column X for cost
        "Accrual Profit",        # Will rename -> 'Total Profit'
        "Accrual Profit Margin"  # Will rename -> 'Margin%'
    ]
    df_selected = df[selected_columns].copy()

    # Calculate ROI from 'Items Cost'
    df_selected["ROI"] = (
        df_selected["Accrual Profit"] / df_selected["Items Cost"]
    ) * 100

    # Profit per Item
    df_selected["Profit per Item"] = (
        df_selected["Accrual Profit"] / df_selected["Qty Sold"]
    )

    # Sales Velocity
    df_selected["Sales Velocity"] = df_selected["Qty Sold"] / days

    # Sort by Qty Sold descending
    df_selected.sort_values(by="Qty Sold", ascending=False, inplace=True)
    return df_selected


def run_full_processing(df_30, df_7, df_3):
    """
    Takes the three raw DataFrames (30/7/3-day),
    applies merges, color formatting, extra sheets,
    and returns a BytesIO containing the final Excel.
    """

    # ---------------------------
    # A) Build df_30_cleaned
    # ---------------------------
    df_30_cleaned = process_profit_summary(df_30, 30)
    df_7_cleaned  = process_profit_summary(df_7,  7)
    df_3_cleaned  = process_profit_summary(df_3,  3)

    # Merge ROI from 7/3-day
    df_30_cleaned = df_30_cleaned.merge(
        df_7_cleaned[["SKU", "ROI"]].rename(columns={"ROI": "ROI_7_Days"}),
        on="SKU",
        how="left"
    )
    df_30_cleaned = df_30_cleaned.merge(
        df_3_cleaned[["SKU", "ROI"]].rename(columns={"ROI": "ROI_3_Days"}),
        on="SKU",
        how="left"
    )

    # ROI Trend Changes
    df_30_cleaned["ROI Change (30D-7D)"] = (
        df_30_cleaned["ROI_7_Days"] - df_30_cleaned["ROI"]
    ).round(2)
    df_30_cleaned["ROI Change (7D-3D)"] = (
        df_30_cleaned["ROI_3_Days"] - df_30_cleaned["ROI_7_Days"]
    ).round(2)

    # Trend Direction
    df_30_cleaned["Trend Direction"] = df_30_cleaned.apply(
        lambda row: "📈 Up" if row["ROI Change (7D-3D)"] > 0
                    else "📉 Down" if row["ROI Change (7D-3D)"] < 0
                    else "➖ Stable",
        axis=1
    )

    # Merge velocity from 7/3-day
    df_30_cleaned = df_30_cleaned.merge(
        df_7_cleaned[["SKU", "Sales Velocity"]].rename(columns={"Sales Velocity": "Velocity_7_Days"}),
        on="SKU", how="left"
    )
    df_30_cleaned = df_30_cleaned.merge(
        df_3_cleaned[["SKU", "Sales Velocity"]].rename(columns={"Sales Velocity": "Velocity_3_Days"}),
        on="SKU", how="left"
    )
    df_30_cleaned.rename(columns={"Sales Velocity": "Velocity_30_Days"}, inplace=True)

    # Rename G-> 'Total Profit', H-> 'Margin%'
    df_30_cleaned.rename(
        columns={
            "Accrual Profit": "Total Profit",
            "Accrual Profit Margin": "Margin%"
        },
        inplace=True
    )

    # Reorder columns
    final_columns = [
        "SKU",
        "Product Name",
        "ASIN",
        "Qty Sold",
        "Total Fees",
        "Items Cost",      # from Column X
        "Total Profit",    # was Accrual Profit
        "Margin%",         # was Accrual Profit Margin
        "ROI",
        "Profit per Item",
        "ROI_7_Days",
        "ROI_3_Days",
        "ROI Change (30D-7D)",
        "ROI Change (7D-3D)",
        "Trend Direction",
        "Velocity_30_Days",
        "Velocity_7_Days",
        "Velocity_3_Days"
    ]
    df_30_cleaned = df_30_cleaned[final_columns]

    # ---------------------------
    # B) Unprofitable Sheet
    # ---------------------------
    df_unprofitable = df_30_cleaned[df_30_cleaned["ROI"] < 5].copy()
    df_unprofitable.sort_values(by="ROI", ascending=False, inplace=True)
    df_unprofitable["Reviewed?"]    = ""
    df_unprofitable["Discontinue?"] = ""

    # ---------------------------
    # C) Downward Trending Velocity
    # ---------------------------
    df_downward_velocity = df_30_cleaned[
        (df_30_cleaned["Velocity_7_Days"] < df_30_cleaned["Velocity_30_Days"]) &
        (df_30_cleaned["Velocity_3_Days"] < df_30_cleaned["Velocity_30_Days"])
    ].copy()
    df_downward_velocity["Velocity Drop (30D->7D)"] = (
        (1 - (df_downward_velocity["Velocity_7_Days"] / df_downward_velocity["Velocity_30_Days"])) * 100
    ).round(2)
    df_downward_velocity.sort_values(by="Velocity Drop (30D->7D)", ascending=False, inplace=True)

    # ---------------------------
    # D) Discontinued (blank)
    # ---------------------------
    df_discontinued = pd.DataFrame(columns=["ProductID", "IsEndOfLife", "Notes"])

    # ---------------------------
    # E) Write All to In-Memory Excel
    # ---------------------------
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
        # 1) Main
        df_30_cleaned.to_excel(writer, sheet_name="Optimized Summary (30 Days)", index=False)
        # 2) Unprofitable
        df_unprofitable.to_excel(writer, sheet_name="Unprofitable - Need to Review", index=False)
        # 3) Downward
        df_downward_velocity.to_excel(writer, sheet_name="Downward Trending Velocity", index=False)
        # 4) Discontinued
        df_discontinued.to_excel(writer, sheet_name="Discontinued Sheet", index=False)

        workbook = writer.book
        worksheet_main = writer.sheets["Optimized Summary (30 Days)"]
        worksheet_unprofitable = writer.sheets["Unprofitable - Need to Review"]
        worksheet_downward = writer.sheets["Downward Trending Velocity"]
        worksheet_discontinued = writer.sheets["Discontinued Sheet"]

        # Autofilter + Auto-Size Column A & B exactly to max data length, ignoring header
        def auto_filter_and_resize(df, ws):
            rows, cols = df.shape
            # Apply autofilter
            ws.autofilter(0, 0, rows, cols - 1)

            if rows > 0:
                # We create text series for each column ignoring the header
                # and strip whitespace.
                col_a_idx = 0
                col_b_idx = 1

                # Column A
                data_col_a = df.iloc[:, col_a_idx].astype(str).str.strip()
                max_len_a = 0 if data_col_a.empty else data_col_a.str.len().max()
                # No +2 => EXACT
                ws.set_column(col_a_idx, col_a_idx, max_len_a)

                # Column B
                if cols > 1:
                    data_col_b = df.iloc[:, col_b_idx].astype(str).str.strip()
                    max_len_b = 0 if data_col_b.empty else data_col_b.str.len().max()
                    ws.set_column(col_b_idx, col_b_idx, max_len_b)

        auto_filter_and_resize(df_30_cleaned, worksheet_main)
        auto_filter_and_resize(df_unprofitable, worksheet_unprofitable)
        auto_filter_and_resize(df_downward_velocity, worksheet_downward)
        auto_filter_and_resize(df_discontinued, worksheet_discontinued)

        # Data Validation for "Reviewed?" & "Discontinue?" in unprofitable
        u_rows, u_cols = df_unprofitable.shape
        reviewed_idx = u_cols - 2
        disc_idx     = u_cols - 1

        worksheet_unprofitable.data_validation(
            1, reviewed_idx, u_rows, reviewed_idx,
            {"validate": "list", "source": ["Yes", "No"]}
        )
        worksheet_unprofitable.data_validation(
            1, disc_idx, u_rows, disc_idx,
            {"validate": "list", "source": ["Yes", "No"]}
        )

        # Conditional Formatting
        green_fill = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red_fill   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

        # 3-color scale for 30-day ROI (Column I)
        worksheet_main.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 10, 'min_type': 'num', 'min_color': '#FF6666',
            'mid_value': 15, 'mid_type': 'num', 'mid_color': '#FFD580',
            'max_value': 20, 'max_type': 'num', 'max_color': '#90EE90'
        })
        worksheet_unprofitable.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 0, 'min_type': 'num', 'min_color': '#FF0000',
            'mid_value': 2.5, 'mid_type': 'num', 'mid_color': '#FFA07A',
            'max_value': 5, 'max_type': 'num', 'max_color': '#FFFF99'
        })
        worksheet_downward.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 10, 'min_type': 'num', 'min_color': '#FF6666',
            'mid_value': 15, 'mid_type': 'num', 'mid_color': '#FFD580',
            'max_value': 20, 'max_type': 'num', 'max_color': '#90EE90'
        })

        # ROI_7_Days (K) & ROI_3_Days (L) vs. ROI (I)
        def apply_roi_cmp(ws):
            ws.conditional_format('K2:K1000', {
                'type': 'formula',
                'criteria': '=$K2>$I2',
                'format': green_fill
            })
            ws.conditional_format('K2:K1000', {
                'type': 'formula',
                'criteria': '=$K2<$I2',
                'format': red_fill
            })
            ws.conditional_format('L2:L1000', {
                'type': 'formula',
                'criteria': '=$L2>$I2',
                'format': green_fill
            })
            ws.conditional_format('L2:L1000', {
                'type': 'formula',
                'criteria': '=$L2<$I2',
                'format': red_fill
            })

        apply_roi_cmp(worksheet_main)
        apply_roi_cmp(worksheet_unprofitable)
        apply_roi_cmp(worksheet_downward)

        # Velocity_7_Days (Q) & Velocity_3_Days (R) vs. Velocity_30_Days (P)
        def apply_vel_cmp(ws):
            ws.conditional_format('Q2:Q1000', {
                'type': 'formula',
                'criteria': '=$Q2>$P2',
                'format': green_fill
            })
            ws.conditional_format('Q2:Q1000', {
                'type': 'formula',
                'criteria': '=$Q2<$P2',
                'format': red_fill
            })
            ws.conditional_format('R2:R1000', {
                'type': 'formula',
                'criteria': '=$R2>$P2',
                'format': green_fill
            })
            ws.conditional_format('R2:R1000', {
                'type': 'formula',
                'criteria': '=$R2<$P2',
                'format': red_fill
            })

        apply_vel_cmp(worksheet_main)
        apply_vel_cmp(worksheet_unprofitable)
        apply_vel_cmp(worksheet_downward)

        # Highlight row if "Reviewed?" & "Discontinue?" = Yes
        highlight_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        def col_idx_to_excel_name(cidx):
            name = ""
            while cidx >= 0:
                name = chr(cidx % 26 + ord('A')) + name
                cidx = cidx // 26 - 1
            return name

        rev_letter  = col_idx_to_excel_name(reviewed_idx)
        disc_letter = col_idx_to_excel_name(disc_idx)
        data_area   = (1, 0, u_rows, u_cols - 1)
        formula     = f'=AND(${rev_letter}2="Yes",${disc_letter}2="Yes")'
        worksheet_unprofitable.conditional_format(
            data_area[0], data_area[1], data_area[2], data_area[3],
            {
                'type': 'formula',
                'criteria': formula,
                'format': highlight_format
            }
        )

    # Rewind buffer
    output_buffer.seek(0)
    return output_buffer


# ------------------------------
# 2) Streamlit UI
# ------------------------------
st.title("Excel Profit Summary App (Using Items Cost)")

file_30 = st.file_uploader("Upload 30-day Excel", type=["xlsx"])
file_7  = st.file_uploader("Upload 7-day Excel",  type=["xlsx"])
file_3  = st.file_uploader("Upload 3-day Excel",  type=["xlsx"])

if st.button("Process Files"):
    if not file_30 or not file_7 or not file_3:
        st.error("Please upload all three Excel files first!")
    else:
        try:
            # Read each file from user upload
            df_30 = pd.read_excel(file_30, sheet_name="Table")
            df_7  = pd.read_excel(file_7,  sheet_name="Table")
            df_3  = pd.read_excel(file_3,  sheet_name="Table")

            # Run the full logic
            final_excel = run_full_processing(df_30, df_7, df_3)

            st.success("Excel processing complete! Use the download button below.")

            # Provide a download button
            st.download_button(
                label="Download Processed Excel",
                data=final_excel.getvalue(),
                file_name="Product_Profit_Summary_With_Trends.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing files: {e}")
