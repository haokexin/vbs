dim oExcel, oWb, oSheet
dim iRowCount, iLoop, iStrIndex, iStrLen
dim strTmp, strChar

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

Set oWb = oExcel.Workbooks.Open("\\VBOXSVR\xp\xls\a.xls")
Set oSheet = oWb.Sheets("SQL Results")

iRowCount = oSheet.usedRange.Rows.Count

for iLoop = 2 to iRowCount
	strTmp = oSheet.Cells(iLoop, 18)
	iStrLen = len(strTmp)
	for iStrIndex = 1 to iStrLen
		strChar = mid(strTmp, iStrIndex, 1)
		if IsNumeric(strChar) then
			oSheet.Cells(iLoop, 10) = strChar
			exit for
		end if
	next
next

msgbox "OK, we are done"
	
oWb.save
oWb.Close
oExcel.Quit

