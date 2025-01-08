import adif_io
import openpyxl

ADIF_NAME = "test.adi"
XLS_NAME = "test.xlsx"

FIELDS = { "id":"序号",
           "qso_date":"日期",
           "time_on":"UTC时间", 
           "band":"波段",
           "freq":"频率",
           "mode":"模式",
           "call":"对方呼号", 
           "rst_sent":"信号报告-发出",
           "rst_rcvd":"信号报告-收到", 
           "gridsquare":"对方网格",
           "my_gridsquare":"我方网格",
           "comment":"备注"}

if __name__ == "__main__":
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "通联日志"
    headerline = []
    for v in FIELDS.values():
        headerline.append(v)
    ws.append(headerline)

    qsos, header = adif_io.read_from_string(open(ADIF_NAME, encoding="utf-8").read())
    ff = FIELDS.copy()
    if "id" in ff.keys():
        del ff["id"]
    fields = ff.keys()
    id = 1
    for qso in qsos:
        row = [id]
        for field in fields:
            row.append(qso.get(field, "/"))
        ws.append(row)
        id += 1

    wb.save(XLS_NAME)