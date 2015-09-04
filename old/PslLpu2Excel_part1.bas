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

Dim langNum As Integer
langNum = prj.Languages.Count

Dim xlsheet As Excel.Worksheet

	Dim src As PslSourceList
	Dim trn As PslTransList

	Dim xlapp As Excel.Application
	Set xlapp = CreateObject("Excel.Application")

	Dim xlwb As Excel.Workbook
	Set xlwb = xlapp.Workbooks.Add
	Set xlsheet = xlwb.ActiveSheet

Dim i,j,k,h As Integer
i = 0
h = 2
For k = 1 To langNum
Dim sheetName As String
	If k <> 1 Then
		Set xlsheet = xlwb.Sheets.Add(After:=xlwb.Sheets(xlwb.Sheets.Count))
	End If
	xlsheet.Name = prj.Languages.Item(k).LangCode
	xlsheet.Cells(1,1) = "Title"
	xlsheet.Cells(1,2) = "Number"
	xlsheet.Cells(1,3) = "ID"
	xlsheet.Cells(1,4) = "Source"
	xlsheet.Cells(1,5) = "Translation"
	xlsheet.Cells(1,6) = "New Translation"
	xlsheet.Cells(1,7) = "Comment"

	For Each trn In prj.TransLists
		If trn.Language.LangCode = prj.Languages.Item(k).LangCode Then
			For j = 1 To trn.StringCount
				Dim tString As PslTransString
				Set tString = trn.String(j)
				If tString.State(pslStateTranslated) = True And tString.State(pslStateReview) = True And tString.State(pslStateLocked) = False And tString.State(pslStateReadOnly) = False Then
					xlsheet.Cells(h,1) = trn.Title
					xlsheet.Cells(h,2) = tString.Number
					xlsheet.Cells(h,3) = tString.ID
					xlsheet.Cells(h,4) = tString.SourceText
					xlsheet.Cells(h,5) = tString.Text
					xlsheet.Cells(h,7) = tString.Comment
					h = h + 1
				End If
		    Next j
		End If
	Next

	With xlsheet.Cells
		.EntireColumn.AutoFit
		.EntireRow.AutoFit
	End With

	h = 2
Next

	xlwb.SaveAs(prjPath)
	xlapp.Quit

	Set xlsheet = Nothing
	Set xlwb = Nothing
	Set xlapp = Nothing

	MsgBox("Done, the file has been saved in " & prjPath)

End Sub
