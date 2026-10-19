# -*- coding: utf-8 -*-
import openpyxl
path = r"D:\opencode\project\oracle_xrisk\20261013_QDII额度\QDII额度.xlsx"
wb = openpyxl.load_workbook(path, data_only=True)
for sn in wb.sheetnames:
    print("=== Sheet:", sn, "===")
    ws = wb[sn]
    print("dims:", ws.dimensions, "max_row:", ws.max_row, "max_col:", ws.max_column)
    # Print first 10 rows
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        print(i, row)
        if i >= 20:
            break
