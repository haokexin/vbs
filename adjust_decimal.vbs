dim oExcel, oWb, oSheet
dim iRowCount, iLoop
dim fMyFloat, fMyDecimal

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

Set oWb = oExcel.Workbooks.Open("\\VBOXSVR\xp\xls\tmp.xls.xls")
Set oSheet = oWb.Sheets("Sheet1")

iRowCount = oSheet.usedRange.Rows.Count

for iLoop = 2 to iRowCount
	fMyFloat = oSheet.Cells(iLoop, 3)
	fMyDecimal = fMyFloat - Fix(fMyFloat)
	select case true
		case (fMyDecimal <= 0.2)
			fMyDecimal = 0
		case (fMyDecimal <= 0.5)
			fMyDecimal = 0.5
		case else
			fMyDecimal = 1
	end select
	fMyFloat = Fix(fMyFloat) + fMyDecimal
	oSheet.Cells(iLoop, 3) = fMyFloat
next

msgbox "OK, we are done"
	
oWb.save
oWb.Close
oExcel.Quit
