"""Convierte el Excel original de UCI a CSV (paso 0)."""
import csv, sys, openpyxl
src, dst = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src, read_only=True)
ws = wb.active
with open(dst, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    for row in ws.iter_rows(values_only=True):
        w.writerow(row)
