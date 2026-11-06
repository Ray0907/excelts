import sys, json, datetime
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

headers = ['id', 'user', 'uniq', 'amount', 'at', 'flag', 'qty', 'cat', 'price', 'note']
wb = Workbook(write_only=True)
ws = wb.create_sheet('Sheet1')
ws.append(headers)
epoch = datetime.datetime(1970, 1, 1)
for line in open(sys.argv[1]):
    r = json.loads(line)
    row = []
    for j, v in enumerate(r):
        if isinstance(v, dict):
            c = WriteOnlyCell(ws, value=epoch + datetime.timedelta(milliseconds=v['d']))
            c.number_format = 'yyyy-mm-dd hh:mm:ss'
            row.append(c)
        else:
            row.append(v)
    ws.append(row)
wb.save(sys.argv[2])
