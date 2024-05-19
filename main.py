import xlwings as xw
import time

book = xw.Book("Book1.xlsx")

sheet1 = book.sheets[0]

v1 = float(sheet1.range("C7").value)
vlph = float(sheet1.range("D7").value)

condition = True
postions = []

while condition:
    time.sleep(5)
    v1 = round((v1 + (v1*0.02)), 2)
    if v1>vlph:
        sheet1.range("H3").value = sheet1.range("B7").value
        if not (sheet1.range("B18").value):
            p = {}
            p_name = sheet1.range("H3").value
            p_price = v1
            p["symbol"] = p_name
            p["price"] = p_price
            postions.append(p)
            # sheet1.range("B18").value = sheet1.range("H3").value
            sheet1.range("B18").value = p_name
            sheet1.range("C18").value = str(v1)

    try:
        sv = float(sheet1.range("C18").value)

        if (((v1-sv)*100)/sv) >= 50.00:
            sheet1.range("H11").value = postions[0]["symbol"] 
    except:
        pass

    print(v1)
    sheet1.range("E7").value = str(v1)

