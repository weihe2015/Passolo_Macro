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
Dim i,j,k As Integer
Dim LangNum As Integer
Dim delete As Boolean

delete = True
i = 0
j = 1

LangNum = prj.Languages.Count


For Each trn In prj.TransLists

	i = i + 1

	If ( i > LangNum) Then

	Exit For

	End If

	Lang = trn.Language.LangCode

	If ((StrComp(Lang,"chs") = 0) Or (StrComp(Lang,"vit") = 0) Or (StrComp(Lang, "cht") = 0) ) Then

	Else

	prj.Languages.Remove(j)

	End If

Next trn


j = 0

i = 1

For Each trn In prj.TransLists

	j = j + 1

	If(j > 3) Then

		j = 1

		i = i + 1

	End If

	If (trn.TransRate <> 100) Then

		delete = False

	End If

		If (delete = True) And (j = 3) Then

			prj.SourceLists.Remove(i)

			i = i - 1

		ElseIf (j = 3) Then

			delete = True

		End If

Next trn


End Sub
