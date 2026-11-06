import sys, json, datetime
from openpyxl import load_workbook

wb = load_workbook(sys.argv[1], read_only=False, data_only=True)
ws = wb.worksheets[0]
epoch = datetime.datetime(1970, 1, 1)
out = []
for row in ws.iter_rows(values_only=True):
    r = []
    for v in row:
        if isinstance(v, datetime.datetime):
            v = round((v - epoch).total_seconds() * 1000)
        r.append(v)
    out.append(r)
json.dump(out, open(sys.argv[2], 'w'))
