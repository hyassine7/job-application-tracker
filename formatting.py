import os
import pandas as pd

def save_with_formatting(df: pd.DataFrame, path: str):
    """
    Writes `df` to `path` as an .xlsx, applying header styles,
    auto-column widths, and any conditional formats you like.
    """
    # Ensure parent folder exists
    os.makedirs(os.path.dirname(path), exist_ok=True)

    # Use XlsxWriter for rich formatting
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Applications")
        workbook  = writer.book
        worksheet = writer.sheets["Applications"]

        # 1) Header formatting
        header_fmt = workbook.add_format({
            "bold":      True,
            "text_wrap": True,
            "valign":    "top",
            "fg_color":  "#DCE6F1",
            "border":    1
        })
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)

        # 2) Auto-fit columns
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)

        # 3) Example conditional format: highlight non-blank Status
        if "Status" in df.columns:
            status_col = df.columns.get_loc("Status")
            worksheet.conditional_format(
                1, status_col, len(df), status_col,
                {
                    "type":     "cell",
                    "criteria": "not equal to",
                    "value":    '""',
                    "format":   workbook.add_format({
                                   "bg_color":   "#C6EFCE",
                                   "font_color": "#006100"
                               })
                }
            )
    # no return; file is written when the context manager exits
