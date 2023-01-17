import xlwings as xw

tgt, src = "data/tables.tsv", "data/item analysis.xls"
tgt_sheet, src_sheet = 0, 0
rng = "A:R"
out = "tables.xlsx"

# Load sheets
target = xw.Book(tgt)
style = xw.Book(src)
src_sheet = style.sheets[src_sheet]
tgt_sheet = target.sheets[tgt_sheet]

# Copy style from src to tgt
src = src_sheet.range(rng)
src.copy()
tgt_sheet.range(rng).paste(paste="formats")

# Save new workbook
target.save(path=out)
target.close()
style.close()
