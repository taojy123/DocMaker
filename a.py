
import sys
import os
import urllib2
import traceback
import re
import xlrd
import xlutils.copy
import win32com
from win32com.client import Dispatch, constants



def repword(w, oldstr, newstr):
    try:
        newstr = str(newstr)
    except:
        pass
    while "  " in newstr:
        newstr = newstr.replace("  ", " ")
    w.Selection.Find.Execute(oldstr, False, False, False, False, False, True, 1, True, newstr, 1)


try:

    open("error.txt", "w")
    open("remarks.txt", "a")
    remarks = open("remarks.txt").read().strip().replace("\r", " ").replace("\n", " ")

    w = win32com.client.DispatchEx('Word.Application')
    w.Visible = 1
    w.DisplayAlerts = 0
    doc = w.Documents.Add(os.getcwd() + "\\mb.doc")

    w.Selection.WholeStory()
    w.Selection.Copy()


    fs = os.listdir("./data")
    for f in fs:
        if f[-3:] == "xls":
            filename = f
    print filename
    wb = xlrd.open_workbook("./data/" + filename, formatting_info=True)
    wb_w = xlutils.copy.copy(wb)   

    sheets = wb.sheets()
               
    ld_total = 0

    count = 1
    for h in range(len(sheets)):
        sheet = sheets[h]
        sheet_w = wb_w.get_sheet(h)

        try:
            date = sheet.cell(1,0).value
            date = re.findall(r".*?(\w+).*?(\w+).*?(\w+)", date)[0]
            year, month, day = date
        except:
            break

        i = 3
        while sheet.cell(i,6).value:

            count += 1

            num = int(sheet.cell(i, 6).value)
            name = sheet.cell(i, 0).value
            quantity = int(sheet.cell(i, 1).value)
            weight = int(sheet.cell(i, 2).value)
            address = sheet.cell(i, 3).value
            price1 = int(sheet.cell(i, 4).value)
            price2 = int(sheet.cell(i, 5).value)
            price = price1 + price2
            ld = int(weight * 0.8 + 0.999)
            if ld < 45:
                ld = 45
            total = price + ld
            ld_total += ld

            sheet_w.write(i, 7, ld)

            w.Selection.PasteAndFormat(16)

            repword(w, "{year}", year)
            repword(w, "{month}", month)
            repword(w, "{day}", day)
            repword(w, "{num}", num)

            repword(w, "{year}", year)
            repword(w, "{month}", month)
            repword(w, "{day}", day)
            repword(w, "{num}", num)

            repword(w, "{name}", name)
            repword(w, "{quantity}", quantity)
            repword(w, "{address}", address)
            repword(w, "{weight}", str(weight) + ".00")
            repword(w, "{price}", str(price) + ".00")
            repword(w, "{ld}", str(ld) + ".00")
            repword(w, "{total}", str(total) + ".00")

            repword(w, "{name}", name)
            repword(w, "{quantity}", quantity)
            repword(w, "{address}", address)
            repword(w, "{weight}", str(weight) + ".00")
            repword(w, "{price}", str(price) + ".00")
            repword(w, "{ld}", str(ld) + ".00")
            repword(w, "{total}", str(total) + ".00")

            repword(w, "{remarks}", remarks)

            w.Selection.MoveDown(Unit=7, Count=3)
            i += 1
            print i

    sheet_w = wb_w.get_sheet(0)  
    sheet_w.write(3, 8, ld_total)

    # w.Selection.InsertBefore(str(ld_total))

    # w.Visible = 1

    # doc.SaveAs(r"new.doc")

    # w.Documents.Close(0)
    # w.Quit()

    wb_w.save(filename)

    print ld_total

    print "ok"

    raw_input()

except Exception, e:
    traceback.print_exc()
    error = traceback.format_exc()
    open("error.txt", "w").write(error)