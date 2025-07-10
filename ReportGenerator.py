import win32com.client as win32

def ReportGeneration():

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    SalesData = excel.Workbooks.Open(
        f"F:\\New folder\\Programs\\Excel goodies\\Python\\30 lines Code project\\Sales_Data.xlsx")
  
    SalesData.Activate()
    Countries = SalesData.Worksheets("Summary")
    Master = SalesData.Worksheets("Sales Master")
    SalesSourceRange = Master.Range("A1", Master.Range("A1").End(win32.constants.xlDown).End(win32.constants.xlToRight))
    CountriesList = Countries.Range("A2:A5")

    for country in CountriesList:
        print(country)
        SalesSourceRange.AutoFilter(4,country)
        SalesSourceRange.Copy()
        TemplateWorkbook = excel.Workbooks.Open(
        f"F:\\New folder\\Programs\\Excel goodies\\Python\\30 lines Code project\\Template.xlsx")
        TemplateWorkbook.Activate()
        TemData = TemplateWorkbook.Worksheets("Data")
        TemReport = TemplateWorkbook.Worksheets("Report")
        TemData.Activate()
        TemData.Range("A1").PasteSpecial()
        excel.CutCopyMode = False
        TemReport.Activate()
        TotalSales = TemReport.Range("B5").Value
        TotalSales = int(TotalSales)
        TotalAmountReceived = TemReport.Range("C5").Value
        TotalAmountReceived = int(TotalAmountReceived)
        TotalBalanceDue = TemReport.Range("D5").Value
        TotalBalanceDue = int(TotalBalanceDue)
        print(f"Total Sales is {TotalSales}")
        if TotalSales > 500000:
            print("Doing Great")
        elif TotalSales > 300000:
            print("Doing Average")
        else:
                print("Need Improvement")
        print(f"Total Amount Received is {TotalAmountReceived}")
        if TotalAmountReceived > 450000:
            print("Doing Great")
        elif TotalAmountReceived > 400000:
            print("Doing Average")
        else:
            print("Need Improvement")
        print(f"Total Balance Due is {TotalBalanceDue}")
        if TotalBalanceDue < 100000:
            print("Doing Good")
        else:
            print("Need Improvement")
        TemplateWorkbook.SaveAs(
            f"F:\\New folder\\Programs\\Excel goodies\\Python\\30 lines Code project\\" + country.Value + ".xlsx")
        TemplateWorkbook.Close()

    SalesData.Activate()
    SalesSourceRange.AutoFilter(4)
    SalesData.Save()
    SalesData.Close()
    excel.Quit()


ReportGeneration()