import streamlit as st
import pandas as pd
from io import BytesIO

def process_profit_summary(df, days):
    selected_columns = [
        "SKU",
        "Product Name",
        "ASIN",
        "Qty Sold",
        "Total Fees",
        "Items Cost",           # Using 'Items Cost' from Column X
        "Accrual Profit",       # Will rename -> 'Total Profit'
        "Accrual Profit Margin" # Will rename -> 'Margin%'
    ]
    df_selected = df[selected_columns].copy()

    # ROI from Items Cost
    df_selected["ROI"] = (df_selected["Accrual Profit"] / df_selected["Items Cost"]) * 100
    df_selected["Profit per Item"] = df_selected["Accrual Profit"] / df_selected["Qty Sold"]
    df_selected["Sales Velocity"] = df_selected["Qty Sold"] / days

    df_selected.sort_values(by="Qty Sold", ascending=False, inplace=True)
    return df_selected

def run_full_processing(df_30, df_7, df_3):

    # 1) Build df_30_cleaned
    df_30_cleaned = process_profit_summary(df_30, 30)
    df_7_cleaned  = process_profit_summary(df_7,  7)
    df_3_cleaned  = process_profit_summary(df_3,  3)

    # Merge ROI from 7/3-day
    df_30_cleaned = df_30_cleaned.merge(
        df_7_cleaned[["SKU", "ROI"]].rename(columns={"ROI": "ROI_7_Days"}),
        on="SKU", how="left"
    )
    df_30_cleaned = df_30_cleaned.merge(
        df_3_cleaned[["SKU", "ROI"]].rename(columns={"ROI": "ROI_3_Days"}),
        on="SKU", how="left"
    )

    # ROI Change
    df_30_cleaned["ROI Change (30D-7D)"] = (df_30_cleaned["ROI_7_Days"] - df_30_cleaned["ROI"]).round(2)
    df_30_cleaned["ROI Change (7D-3D)"]  = (df_30_cleaned["ROI_3_Days"] - df_30_cleaned["ROI_7_Days"]).round(2)

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
        "Items Cost",
        "Total Profit",
        "Margin%",
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

    # 2) Unprofitable
    df_unprofitable = df_30_cleaned[df_30_cleaned["ROI"] < 5].copy()
    df_unprofitable.sort_values(by="ROI", ascending=False, inplace=True)
    df_unprofitable["Reviewed?"]    = ""
    df_unprofitable["Discontinue?"] = ""

    # 3) Downward Trending Velocity
    df_downward_velocity = df_30_cleaned[
        (df_30_cleaned["Velocity_7_Days"] < df_30_cleaned["Velocity_30_Days"]) &
        (df_30_cleaned["Velocity_3_Days"] < df_30_cleaned["Velocity_30_Days"])
    ].copy()
    df_downward_velocity["Velocity Drop (30D->7D)"] = (
        (1 - df_downward_velocity["Velocity_7_Days"] / df_downward_velocity["Velocity_30_Days"]) * 100
    ).round(2)
    df_downward_velocity.sort_values(by="Velocity Drop (30D->7D)", ascending=False, inplace=True)

    # 4) Discontinued (blank)
    df_discontinued = pd.DataFrame(columns=["ProductID", "IsEndOfLife", "Notes"])

    # 5) Round all numeric columns to 2 decimals in the DataFrame
    for df_ in [df_30_cleaned, df_unprofitable, df_downward_velocity, df_discontinued]:
        numeric_cols = df_.select_dtypes(include=["float", "int"]).columns
        df_[numeric_cols] = df_[numeric_cols].round(2)

    # 6) Write to Excel with a numeric format in xlsxwriter
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
        df_30_cleaned.to_excel(writer, "Optimized Summary (30 Days)", index=False)
        df_unprofitable.to_excel(writer, "Unprofitable - Need to Review", index=False)
        df_downward_velocity.to_excel(writer, "Downward Trending Velocity", index=False)
        df_discontinued.to_excel(writer, "Discontinued Sheet", index=False)

        workbook = writer.book
        ws_main = writer.sheets["Optimized Summary (30 Days)"]
        ws_unp  = writer.sheets["Unprofitable - Need to Review"]
        ws_down = writer.sheets["Downward Trending Velocity"]
        ws_disc = writer.sheets["Discontinued Sheet"]

        # Numeric format => 2 decimals
        two_dec_format = workbook.add_format({"num_format": "0.00"})

        def auto_filter_and_resize(df, ws):
            rows, cols = df.shape
            ws.autofilter(0, 0, rows, cols - 1)

            if rows > 0:
                # Apply the 2-decimal format to each numeric column
                numeric_cols = df.select_dtypes(include=["float","int"]).columns
                for col_name in numeric_cols:
                    col_idx = df.columns.get_loc(col_name)
                    # We won't set a specific width here, just the format
                    # If you want exact width, do `ws.set_column(col_idx, col_idx, width, two_dec_format)`
                    ws.set_column(col_idx, col_idx, None, two_dec_format)

                # Auto-size col A (SKU) ignoring header
                col_a_idx = 0
                col_b_idx = 1
                data_col_a = df.iloc[:, col_a_idx].astype(str).str.strip()
                max_len_a = 0 if data_col_a.empty else data_col_a.str.len().max()
                ws.set_column(col_a_idx, col_a_idx, max_len_a)

                # Auto-size col B (Product Name) ignoring header
                if cols > 1:
                    data_col_b = df.iloc[:, col_b_idx].astype(str).str.strip()
                    max_len_b = 0 if data_col_b.empty else data_col_b.str.len().max()
                    ws.set_column(col_b_idx, col_b_idx, max_len_b)

        auto_filter_and_resize(df_30_cleaned, ws_main)
        auto_filter_and_resize(df_unprofitable, ws_unp)
        auto_filter_and_resize(df_downward_velocity, ws_down)
        auto_filter_and_resize(df_discontinued, ws_disc)

        # Data Validation for unprofitable
        u_rows, u_cols = df_unprofitable.shape
        reviewed_idx = u_cols - 2
        disc_idx     = u_cols - 1
        ws_unp.data_validation(
            1, reviewed_idx, u_rows, reviewed_idx,
            {"validate": "list", "source": ["Yes", "No"]}
        )
        ws_unp.data_validation(
            1, disc_idx, u_rows, disc_idx,
            {"validate": "list", "source": ["Yes", "No"]}
        )

        # Color Formats
        green_fill = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red_fill   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

        # 3-color scale for ROI (I2:I1000) in each sheet
        ws_main.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 10, 'min_type': 'num', 'min_color': '#FF6666',
            'mid_value': 15, 'mid_type': 'num', 'mid_color': '#FFD580',
            'max_value': 20, 'max_type': 'num', 'max_color': '#90EE90'
        })
        ws_unp.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 0, 'min_type': 'num', 'min_color': '#FF0000',
            'mid_value': 2.5, 'mid_type': 'num', 'mid_color': '#FFA07A',
            'max_value': 5, 'max_type': 'num', 'max_color': '#FFFF99'
        })
        ws_down.conditional_format('I2:I1000', {
            'type': '3_color_scale',
            'min_value': 10, 'min_type': 'num', 'min_color': '#FF6666',
            'mid_value': 15, 'mid_type': 'num', 'mid_color': '#FFD580',
            'max_value': 20, 'max_type': 'num', 'max_color': '#90EE90'
        })

        # ROI_7_Days & ROI_3_Days vs ROI
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

        apply_roi_cmp(ws_main)
        apply_roi_cmp(ws_unp)
        apply_roi_cmp(ws_down)

        # Velocity_7_Days & Velocity_3_Days vs Velocity_30_Days
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

        apply_vel_cmp(ws_main)
        apply_vel_cmp(ws_unp)
        apply_vel_cmp(ws_down)

        # Highlight row if "Reviewed?" & "Discontinue?" = Yes
        highlight_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        def col_idx_to_excel_name(c):
            name = ""
            while c >= 0:
                name = chr(c % 26 + ord('A')) + name
                c = c // 26 - 1
            return name

        formula = ""
        if not df_unprofitable.empty:
            rev_letter  = col_idx_to_excel_name(reviewed_idx)
            disc_letter = col_idx_to_excel_name(disc_idx)
            data_area   = (1, 0, u_rows, u_cols - 1)
            formula     = f'=AND(${rev_letter}2="Yes",${disc_letter}2="Yes")'
            ws_unp.conditional_format(
                data_area[0], data_area[1], data_area[2], data_area[3],
                {
                    'type': 'formula',
                    'criteria': formula,
                    'format': highlight_format
                }
            )

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
            df_30 = pd.read_excel(file_30, sheet_name="Table")
            df_7  = pd.read_excel(file_7,  sheet_name="Table")
            df_3  = pd.read_excel(file_3,  sheet_name="Table")

            final_excel = run_full_processing(df_30, df_7, df_3)

            st.success("Excel processing complete! Use the download button below.")
            st.download_button(
                label="Download Processed Excel",
                data=final_excel.getvalue(),
                file_name="Product_Profit_Summary_With_Trends.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing files: {e}")
