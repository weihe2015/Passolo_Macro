Option Explicit

Sub Main

Dim src As PslSourceList
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

Dim j,k As Integer

Dim encoding As Integer

Dim sNumber As Integer

Dim sID As String

Dim out As String

For Each src In prj.SourceLists

	For j = 1 To src.StringCount

	Dim sString As PslSourceString

	Set sString = src.String(j)

	Dim sText As String

	sText = sString.Text

		For k = 1 To Len(sText)

		encoding = Asc(Mid(sText,k,1))

		If encoding > 127 Or ((encoding = 63) And (Asc(Mid(sourceText,j+1,1)) = 63)) Then

			sID = sString.ID

			sNumber = sString.Number

			out = "Invalid character in source string list: " & src.LangID & " " & src.Title & " of number " & sNumber & Chr(13) & Chr(10)

			objFile2.writeLine(out)

			GoTo NextSrcString

		End If

		Next

NextSrcString:
	Next j

Next

For Each trn In prj.TransLists

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

				out = "Invalid character in string list: " & trn.Language.LangCode & " " & trn.Title & " of number " & sNumber & Chr(13) & Chr(10)

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

Next trn

objFile2.Close

PSL.Output ("--- Done ---")
'MsgBox ("Successfully scanning all string lists!!!")

End Sub
