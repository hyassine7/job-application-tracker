# formatting.py

import pandas as pd

def save_with_formatting(df: pd.DataFrame, path: str):
    """
    Writes the DataFrame to `path` with basic Excel formatting
    (autofit columns, header bold, etc.). Uses xlsxwriter.
    """
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        wb  = writer.book
        ws  = writer.sheets["Sheet1"]

        # Bold headers
        hdr_fmt = wb.add_format({"bold": True})
        for col_idx, col in enumerate(df.columns):
            ws.write(0, col_idx, col, hdr_fmt)
            # auto-fit: measure max length in column
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            ws.set_column(col_idx, col_idx, max_len)
