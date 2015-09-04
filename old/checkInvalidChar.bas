Option Explicit

Sub Main

Dim trn As PslTransList
Dim prj As PslProject
Set prj = PSL.ActiveProject

If prj Is Nothing Then
	MsgBox("No active Passolo Project")
	Exit Sub
End If

Dim i As Long

Const ForWriting = 2

Dim objFSO, objFile, objFile2

Dim logFile As String

logFile = prj.Location + "\" + prj.Name +".log"

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(logFile,True)

objFile.Close

Set objFile2 = objFSO.OpenTextFile(logFile,ForWriting)

PSL.Output("Begin to scan all the string lists...")

Dim sourceText As String

Dim j As Integer

Dim encoding As Integer

Dim sNumber As Integer

Dim sID As String

Dim out As String

Dim test As String

For Each trn In prj.TransLists

	If StrComp(trn.Title,test) = 0 Then

	For i = 1 To trn.StringCount

		Dim tString As PslTransString

		Set tString = trn.String(i)

		If (tString.State(pslStateReadOnly) = False And tString.State(pslStateLocked) = False ) Then

		sourceText = tString.SourceText

		For j = 1 To Len(sourceText)

			encoding = Asc(Mid(sourceText,j,1))

			If(encoding = 63) And (Asc(Mid(sourceText,j+1,1)) = 63) Then

				sID = tString.ID

				sNumber = tString.Number

				out = "Invalid character in string list: " & trn.Title & " of number " & sNumber & Chr(13) & Chr(10)

				objFile2.writeLine(out)

				sNumber = 0

				sID = ""

				out = ""

				GoTo nextStringList

			End If

		Next j

	  End If

nextStringList:

	Next i

	End If

Next trn

objFile2.Close

PSL.Output ("--- Done ---")
MsgBox ("Successfully scanning all string lists!!!")

End Sub
