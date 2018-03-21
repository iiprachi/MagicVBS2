On Error Resume Next

strExcelFile = "C:\\giteg\\magicvbs_2\\data.xlsx"
strXMLStubFile = "C:\\giteg\\magicvbs_2\\stub.xml"
strXMLOutputFile = "C:\\giteg\\magicvbs_2\\output1.xml"
headerSheet = 1
transactionSheet = 2
coverageSheet = 3
subjectSheet = 4

'EXCEL Stuff

Dim objExcel,ObjWorkbook,objsheet
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(strExcelFile)
objExcel.DisplayAlerts = 0 


'XML Stuff

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
xmlDoc.Load(strXMLStubFile)
Set objRoot = xmlDoc.documentElement

Function FillHeaderDetails()

	count = GetColumnCount(headerSheet)
	
	i = 1
	do while i < count
		columnName = GetColumnName(i,headerSheet)
		Set colNodes = xmlDoc.selectNodes("//"&columnName)
		For Each objNode in colNodes
			objNode.Text = GetData(columnName,headerSheet)
		Next
		i = i + 1
	loop

End function

Function FillTransactionDetails()

	Rows = GetRowCount(transactionSheet)
	ColumnCount = GetColumnCount(transactionSheet)
	row = 2
	do while row <= Rows 
	i = 1
	set transactionTag = xmlDoc.createElement("Transaction")
		do while i <= ColumnCount
			ColumnName = GetColumnName (i,transactionSheet)
			value = GetDataByRow(columnName,row,transactionSheet)
			set colTag = xmlDoc.createElement(columnName)
			colTag.Text = value
			transactionTag.appendChild(colTag)
			i = i+1
		loop
	Set colNodes = xmlDoc.selectNodes("//History")
	For Each objNode in colNodes
		objNode.appendChild(transactionTag)
	Next
	row = row + 1
	loop
end function



Function FillCoverageDetails()

	Rows = GetRowCount(coverageSheet)
	ColumnCount = GetColumnCount(coverageSheet)
	row = 2
	do while row <= Rows 
	i = 1
	set parentTag = xmlDoc.createElement("CoverageItem")
		do while i <= ColumnCount
			ColumnName = GetColumnName (i,coverageSheet)
			value = GetDataByRow(columnName,row,coverageSheet)
			set colTag = xmlDoc.createElement(columnName)
			colTag.Text = value
			parentTag.appendChild(colTag)
			i = i+1
		loop
	Set colNodes = xmlDoc.selectNodes("//PolicyCoverageInfo")
	For Each objNode in colNodes
		objNode.appendChild(parentTag)
	Next
	row = row + 1
	loop
end function



Function FillSubjectInfoDetails()

	Rows = GetRowCount(subjectSheet)
	ColumnCount = GetColumnCount(subjectSheet)
	row = 2
	do while row <= Rows 
	i = 1
	set parentTag = xmlDoc.createElement("SubjectInfo")
		do while i <= ColumnCount
			ColumnName = GetColumnName (i,subjectSheet)
			value = GetDataByRow(columnName,row,subjectSheet)
			set colTag = xmlDoc.createElement(columnName)
			colTag.Text = value
			parentTag.appendChild(colTag)
			i = i+1
		loop
	Set colNodes = xmlDoc.selectNodes("//SubjectDetail")
	For Each objNode in colNodes
		objNode.appendChild(parentTag)
	Next
	row = row + 1
	loop
end function



Function GetColumnName(index,sheet)
    Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheet)
	ColoumnCount = objSheet.usedRange.Columns.Count
    GetColumnName = objSheet.Cells(1,index)
End Function

Function GetColumnCount(sheet)
    Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheet)
	ColoumnCount = objSheet.usedRange.Columns.Count
    GetColumnCount = ColoumnCount
End Function

Function GetRowCount(sheet)
    Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheet)
	rowCount = objSheet.usedRange.Rows.Count
    GetRowCount = rowCount
End Function


Function GetData(tag,sheet)

    Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheet)
	RowsCount = objSheet.usedRange.Rows.Count
	ColoumnCount = objSheet.usedRange.Columns.Count
	i = 1
	found = false
	do while i <= ColoumnCount
	    if tag = objSheet.Cells(1,i) then
		    GetData = objSheet.Cells(2,i)
			found = true
			Exit Do 
		end if
		i = i + 1
	loop
	if found <> true then
	GetData = "not found"
	end if
end Function


Function GetDataByRow(tag,r,sheet)

    Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheet)
	RowsCount = objSheet.usedRange.Rows.Count
	ColoumnCount = objSheet.usedRange.Columns.Count
	i = 1
	found = false
	do while i <= ColoumnCount
	    if tag = objSheet.Cells(1,i) then
		    GetDataByRow = objSheet.Cells(r,i)
			found = true
			Exit Do 
		end if
		i = i + 1
	loop
	if found <> true then
	GetDataByRow = "not found"
	end if
end Function


FillHeaderDetails()
FillTransactionDetails()
FillCoverageDetails()
FillSubjectInfoDetails()


If Err.Number <> 0 then
MsgBox  Err.Number & " " & Err.Description
end if




objWorkbook.Close  
objExcel.Application.Quit
objExcel.Quit

msgBox("End")

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
xmlDoc.Save strXMLOutputFile  
