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

Dim ExcelPath As String
ExcelPath = prj.Location & "\Exception_List.xlsx"

Dim langNum As Integer
langNum = prj.Languages.Count

Dim xlsheet As Excel.Worksheet

	Dim src As PslSourceList
	Dim trn As PslTransList

	Dim xlapp As Excel.Application
	Set xlapp = CreateObject("Excel.Application")

	Dim xlwb As Excel.Workbook
	Set xlwb = xlapp.Workbooks.Open(ExcelPath)
	Set xlsheet = xlwb.Worksheets(1)

Dim ExceptionChar As String
Dim i,j,k As Integer
i = 1
Do Until xlsheet.Cells(i,1).Value = ""
	j = 1
	Do Until xlsheet.Cells(i,j).Value = ""
		ExceptionChar = ExceptionChar & xlsheet.Cells(i,j).Value
		j = j + 1
	Loop
	i = i + 1
Loop

	'xlwb.SaveAs(prjPath)
	xlapp.Quit

	Set xlsheet = Nothing
	Set xlwb = Nothing
	Set xlapp = Nothing


Const ForWriting = 2

Dim objFSO, objFile, objFile2

Dim logFile As String

logFile = prj.Location + "\" + prj.Name +"_InvalidChar.log"

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(logFile,True)

objFile.Close

Set objFile2 = objFSO.OpenTextFile(logFile,ForWriting)
Dim outprint As String
Dim encoding As Integer

For Each src In prj.SourceLists
	For j = 1 To src.StringCount
		Dim sString As PslSourceString
		Set sString = src.String(j)
		Dim sText As String
		sText = sString.Text
		For k = 1 To Len(sText)
			encoding = Asc(Mid(sText,k,1))
			If encoding > 127 And InStr(ExceptionChar,Mid(sText,k,1)) = 0 Then
				outprint = "Invalid character of encoding " & encoding & " in source string list: " src.Title & " of number " & sNumber & Chr(13) & Chr(10)
				objFile2.writeLine(outprint)
			End If
		Next
	Next j
Next

For Each trn In prj.TransLists
	For i = 1 To trn.StringCount
		Dim tStrng As PslTransString
		Set tString = trn.String(i)
		If (tString.State(pslStateReadOnly) = False And tString.State(pslStateLocked) = False ) Then

		End If
	Next
Next

	PSL.Output("Done")

End Sub
