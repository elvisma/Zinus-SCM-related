Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open"C:\Users\Elvis Ma\Desktop\Weekly Work\Actual Sales\eBay_testing.xlsm"
objExcel.Run"ThisWorkbook.Auto_eBay_Actual"
objExcel.DisplayAlerts = False

objExcel.Activeworkbook.Save
objExcel.Activeworkbook.Close(0)
objExcel.Quit