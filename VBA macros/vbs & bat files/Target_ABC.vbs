Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open"C:\Users\Elvis Ma\Desktop\Weekly Work\Actual Sales\target_testing.xlsm"
objExcel.Run"ThisWorkbook.Auto_Target_ABC"
objExcel.DisplayAlerts = False

objExcel.Activeworkbook.Save
objExcel.Activeworkbook.Close(0)
objExcel.Quit