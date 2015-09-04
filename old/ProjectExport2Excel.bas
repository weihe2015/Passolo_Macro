'#Reference {00020813-0000-0000-C000-000000000046}#1.5#0#C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE#Microsoft Excel 11.0 Object Library
Dim xlsheet As Excel.Worksheet
Option Explicit

Sub Main
	Dim prj As PslProject
	Set prj = PSL.ActiveProject

	Dim src As PslSourceList
	Dim trn As PslTransList

	Dim xlapp As Excel.Application
	Set xlapp = CreateObject("Excel.Application")


	Dim xlwb As Excel.Workbook
	Set xlwb = xlapp.Workbooks.Add
	Set xlsheet = xlwb.ActiveSheet

	xlsheet.Cells(1,1) = "Number"
	xlsheet.Cells(1,2) = "ID"
	xlsheet.Cells(1,3) = "State"
	xlsheet.Cells(1,4) = "English"

	xlsheet.Name = "oauth.js"

	xlapp.Visible = True
	Set xlsheet = Nothing
	Set xlwb = Nothing
	Set xlapp = Nothing
End Sub
