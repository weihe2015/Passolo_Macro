'#Language "WWB-COM"

Option Explicit

Sub Main
Dim prj As PslProject
Set prj = PSL.ActiveProject
Dim objFSO, objFile, objFile2
Dim logFile As String
Const ForWriting = 2
Const ForReading = 1
logFile = prj.Location + "\" + prj.Name +".txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(logFile,True)
objFile.Close
Set objFile2 = objFSO.OpenTextFile(logFile,ForWriting)
objFile2.writeLine("hello")
objFile2.Close
End Sub
