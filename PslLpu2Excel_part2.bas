'#Reference {00020813-0000-0000-C000-000000000046}#1.5#0#C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE#Microsoft Excel 11.0 Object Library

Option Explicit

Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If

Dim prjPath As String
prjPath = prj.Location & "\" & prj.Name & ".xlsx"
Dim newExcel As String
newExcel = prj.Location & "\" & prj.Name & "_ChangedStringsReport.xlsx"

Dim langNum As Integer
langNum = prj.Languages.Count

Dim xlsheet As Excel.Worksheet

	Dim src As PslSourceList
	Dim trn As PslTransList

	Dim xlapp As Excel.Application
	Set xlapp = CreateObject("Excel.Application")

	Dim xlwb As Excel.Workbook
	Set xlwb = xlapp.Workbooks.Open(prjPath)

Dim deleteRow As Boolean

Dim i,j,k As Integer
i = 2
For k = 1 To langNum
	Set xlsheet = xlwb.Worksheets(k)
	Do Until xlsheet.Cells(i,1).Value = ""

		Dim title,number,source,target As String
		title = xlsheet.Cells(i,1).Value
		number = xlsheet.Cells(i,2).Value
		source = xlsheet.Cells(i,4).Value
		target = xlsheet.Cells(i,5).Value

		For Each trn In prj.TransLists
		If trn.Language.LangCode = prj.Languages.Item(k).LangCode And trn.Title = title Then
			For j = 1 To trn.StringCount
				Dim tString As PslTransString
				Set tString = trn.String(j)
				If tString.Number = number And tString.SourceText = source And tString.Text <> target Then
					xlsheet.Cells(i,6) = tString.Text

				End If
			Next
		End If
		Next
		i = i + 1
	Loop

	i = 2
	Do Until xlsheet.Cells(i,1).Value = ""
		If xlsheet.Cells(i,6).Value = "" Then
			xlsheet.Cells(i,1).EntireRow.Delete
		End If
		i = i + 1
	Loop

	If xlsheet.Cells(xlsheet.UsedRange.Rows.Count,6).Value = "" Then
		xlsheet.Cells(xlsheet.UsedRange.Rows.Count,1).EntireRow.Delete
	End If

	With xlsheet.Cells
		.EntireColumn.AutoFit
		.EntireRow.AutoFit
	End With

Next

	xlwb.SaveAs(newExcel)
	xlapp.Quit

	Set xlsheet = Nothing
	Set xlwb = Nothing
	Set xlapp = Nothing

	'MsgBox("Done, the file has been saved in " & prjPath)

End Sub
