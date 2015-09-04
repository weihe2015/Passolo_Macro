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

Dim xlsheet As Excel.Worksheet

	Dim src As PslSourceList
	Dim trn As PslTransList

	Dim xlapp As Excel.Application
	Set xlapp = CreateObject("Excel.Application")

	Dim xlwb As Excel.Workbook
	Set xlwb = xlapp.Workbooks.Add
	Set xlsheet = xlwb.ActiveSheet

Dim sheetName As String
sheetName = prj.TransLists(1).Title

If InStr(sheetName,"\") > 0 Then

	sheetName = Replace(sheetName,"\","_")

End If

If Len(sheetName) > 31 Then

	sheetName = Left(sheetName,31)

End If

	xlsheet.Name = sheetName
	xlsheet.Cells(1,1) = "Number"
	xlsheet.Cells(1,2) = "ID"
	xlsheet.Cells(1,3) = "State"
	xlsheet.Cells(1,4) = "English"

Dim i,j,k As Integer
i = 0
k = 1
Dim tranEng As Boolean
tranEng = True

Dim lang As String

	For Each trn In prj.TransLists

		i = i + 1

		If i > prj.Languages.Count Then

		i = 1

		tranEng = True

		With xlsheet.Cells
			.WrapText = True
			.EntireColumn.AutoFit
			.EntireRow.AutoFit
		End With

		sheetName = trn.Title

			If InStr(sheetName,"\") > 0 Then

				sheetName = Replace(sheetName,"\","_")

			End If

			If Len(sheetName) > 31 Then

				sheetName = Left(sheetName,31)

			End If

		Set xlsheet = xlwb.Sheets.Add(After:=xlwb.Sheets(xlwb.Sheets.Count))

		If check(xlwb,sheetName) = True Then

			sheetName = sheetName & k

			k = k + 1

		End If
		xlsheet.Name = sheetName

		xlsheet.Cells(1,1) = "Number"
		xlsheet.Cells(1,2) = "ID"
		xlsheet.Cells(1,3) = "State"
		xlsheet.Cells(1,4) = "English"

		End If

		lang = LangName(trn)

		xlsheet.Cells(1,4+i) = lang

		For j = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(j)

			xlsheet.Cells(j+1,1) = tString.Number
			xlsheet.Cells(j+1,2) = tString.ID
			xlsheet.Cells(j+1,3) = tString.State(1)

			If tranEng = True Then
				xlsheet.Cells(j+1,4) = tString.SourceText
			End If

			xlsheet.Cells(j+1,4+i) = tString.Text

			If StrComp(tString.SourceText,tString.Text) = 0 Then

				xlsheet.Cells(j+1,4+i).Interior.ColorIndex = 6

			End If
		Next j

   		tranEng = False

	Next trn

	With xlsheet.Cells
		.WrapText = True
		.EntireColumn.AutoFit
		.EntireRow.AutoFit
	End With

	xlwb.SaveAs(prjPath)
	xlapp.Quit

	Set xlsheet = Nothing
	Set xlwb = Nothing
	Set xlapp = Nothing


	'MsgBox("Done, the file has been saved in " & prjPath)

End Sub

Function check(xlwb As Excel.Workbook,currName As String) As Boolean

	check = False

	Dim ws As Excel.Worksheet

	For Each ws In xlwb.Sheets

		If ws.Name = currName Then

			check = True

			Exit Function

		End If

	Next ws

End Function

Function LangName(trn As PslTransList) As String

Select Case trn.Language.LangCode

Case "dan"
	LangName = "Danish"
Case "fin"
	LangName = "Finnish"
Case "sve"
	LangName = "Swedish"
Case "chs"
	LangName = "Chinese"
Case "eti"
	LangName = "Estonian"
Case "ita"
	LangName = "Italian"
Case "lth"
	LangName = "Lithuanian"
Case "lvi"
	LangName = "Latvian"
Case "nld"
	LangName = "Dutch"
Case "plk"
	LangName = "Polish"
Case "ptb"
	LangName = "Portuguese Portugal"
Case "ptg"
	LangName = "Portuguese Brazil"
Case "rom"
	LangName = "Romanian"
Case "ara"
	LangName = "Arabic"
Case "csy"
	LangName = "Czech"
Case "deu"
	LangName = "German"
Case "esp"
	LangName = "Spanish"
Case "fra"
	LangName = "French"
Case "heb"
	LangName = "Hebrew"
Case "rus"
	LangName = "Russian"
Case "jpn"
	LangName = "Japanese"
Case "kor"
	LangName = "Korean"
Case "tha"
	LangName = "Thai"
Case "trk"
	LangName = "Turkish"
Case "nor"
	LangName = "Norwegian"
Case "vit"
	LangName = "Vietnamese"
Case "ell"
	LangName = "Greek"
Case "cht"
	LangName = "Chinese Traditional"
End Select

End Function
