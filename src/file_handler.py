import os
import pandas as pd


def id_extract(file_path="input.txt"):
    ids = []
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                stripped = line.strip()
                if stripped:
                    ids.append(stripped)
        return ids
    return None


def write_excel(df, file_name="output.xlsx"):
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        red_format = workbook.add_format({'font_color': '#FF0000'})

        for i, row in df.iterrows():
            if any(pd.isnull(x) for x in row.iloc[1:]):
                worksheet.write(i + 1, 0, row.iloc[0], red_format)
                for j in range(1, len(row)):
                    worksheet.write(i + 1, j, '')
            else:
                for j, value in enumerate(row):
                    worksheet.write(i + 1, j, value)

        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
            worksheet.set_column(idx, idx, max_len)
