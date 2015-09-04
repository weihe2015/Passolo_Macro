Option Explicit

Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If

Const ForWriting = 2

Dim logFile As String

logFile = prj.Location + "\" + prj.Name + ".log"

Dim objFSO, objFile, objFile2

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Dir(logFile) <> "" Then

	Set objFile2 = objFSO.OpenTextFile(logFile,ForWriting)

Else

	Set objFile = objFSO.CreateTextFile(logFile,True)
	objFile.Close
	Set objFile2 = objFSO.OpenTextFile(logFile,ForWriting)

End If

Dim prjName As String

prjName = prj.Name

If (InStr(prj.Name,"ECI_10") > 0) Then

	ECI_th(prj,objFile2)

ElseIf (InStr(prj.Name,"ECI") > 0) Then

	ECI(prj,objFile2)

ElseIf (InStr(prj.Name,"AAC") > 0) Then

	AAC(prj,objFile2)

ElseIf (InStr(prj.Name,"TOIN") > 0) Then

	TOIN(prj,objFile2)

ElseIf (InStr(prj.Name,"LION_10") > 0) Then

	LION_Self(prj,objFile2)

ElseIf (InStr(prj.Name,"LION") > 0) Then

	LION_main(prj,objFile2)

End If

objFile2.Close

End Sub

Function ECI_th(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

'Remove the language sets

'The pointer of each string list of each language
'It goes thought only one string list
i = 0

'The pointer of language
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If (StrComp(Lang,"tha") = 0 ) Then

	j = j + 1

	Else

	logoutput = "Removing language: " & Lang

	objFile2.writeLine(logoutput)

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True

j = 0

i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 1) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 1) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 1) Then

			delete = True

		End If

Next trn

End Function

Function ECI(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

'Remove the language sets

'The pointer of each string list of each language
'It goes thought only one string list
i = 0

'The pointer of language
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"chs") = 0) Or (StrComp(Lang,"vit") = 0) Or (StrComp(Lang, "cht") = 0) ) Then

	j = j + 1

	Else

	logoutput = "Removing language " & Lang

	objFile2.writeLine(logoutput)

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True
'The pointer of each string list of each language
j = 0

'The pointer of each string list
i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 3) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 3) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 3) Then

			delete = True

		End If

Next trn


End Function

Function AAC(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

'Remove the language sets

'The pointer of each string list of each language
'It goes thought only one string list
i = 0

'The pointer of language
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"sve") = 0) Or (StrComp(Lang,"fin") = 0) Or (StrComp(Lang, "dan") = 0) Or (StrComp(Lang,"nor") = 0) Or (StrComp(Lang,"plk") = 0) Or (StrComp(Lang,"nld") = 0)) Then

	j = j + 1

	Else

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True
'The pointer of each string list of each language
j = 0

'The pointer of each string list
i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 6) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 6) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 6) Then

			delete = True

		End If

Next trn

End Function

Function TOIN(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

i = 0
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"jpn") = 0) Or (StrComp(Lang,"kor") = 0)) Then

	j = j + 1

	Else

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True
'The pointer of each string list of each language
j = 0

'The pointer of each string list
i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 2) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 2) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 2) Then

			delete = True

		End If

Next trn

End Function


Function LION_Self(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

i = 0
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"eti") = 0) Or (StrComp(Lang,"lth") = 0) Or (StrComp(Lang, "lvi") = 0) Or (StrComp(Lang, "ptg") = 0) ) Then

	j = j + 1

	Else

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True
j = 0

i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 4) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 4) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 4) Then

			delete = True

		End If

Next trn

End Function


Function LION_main(prj,objFile2)

Dim Lang As String
Dim trn As PslTransList
Dim i,j,k As Integer
Dim LangNum As Integer
Dim logoutput As String
logoutput = "Begin to process file " & prj.Name
objFile2.writeLine(logoutput)

i = 0
j = 1

LangNum = prj.Languages.Count

For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"ita") = 0) Or (StrComp(Lang,"ptb") = 0) Or (StrComp(Lang, "rom") = 0) Or (StrComp(Lang,"ara") = 0) Or (StrComp(Lang,"csy") = 0) Or (StrComp(Lang, "deu") = 0) Or (StrComp(Lang,"fra") = 0) Or (StrComp(Lang,"heb") = 0) Or (StrComp(Lang, "rus") = 0) Or (StrComp(Lang,"trk") = 0) Or (StrComp(Lang,"ell") = 0) Or (StrComp(Lang, "esp") = 0)) Then

	j = j + 1

	Else

	prj.Languages.Remove(j)

	End If

Next trn

'Delete those source lists that are all translated and validated
Dim delete As Boolean
delete = True
j = 0

i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 12) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

		For k = 1 To trn.StringCount

			Dim tString As PslTransString

			Set tString = trn.String(k)

		    If StrComp(tString.SourceText,tString.Text) <> 0 And tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = True Then

				logoutput = "Find one source and translation are not identical in untranslated state. Number: " & tString.Number & "in Stringlist: " & trn.Title

				objFile2.writeLine(logoutput)

		    	tString.Text = tString.SourceText

				tString.TransList.Save

		    End If

		Next k

	End If

		If (delete = True) And (j = 12) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 12) Then

			delete = True

		End If

Next trn

End Function
