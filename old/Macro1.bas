'#Language "WWB-COM"

Option Explicit

Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If



Dim Lang As String
Dim trn As PslTransList
Dim i,j As Integer
Dim LangNum As Integer

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

	If (StrComp(Lang,"fra") = 0 Or StrComp(Lang,"deu") = 0) Then

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

	If(j > 2) Then

		j = 1

		i = i + 1

	End If

	If (StrComp(trn.Title,"JS\AddExternalData\common") = 0 Or StrComp(trn.Title,"JS\AddSharePointList") = 0 ) Then

		delete = False

	End If

		If (delete = True) And (j = 2) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 2) Then

			delete = True

		End If

Next trn

End Sub
