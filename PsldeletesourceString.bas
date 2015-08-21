Option Explicit

Sub Main
Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If

Dim delete As Boolean
delete = True

Dim i, j As Integer

j = 0

i = 1

Dim trn As PslTransList

For Each trn In prj.TransLists

	j = j + 1

	If (j > 3) Then
		j = 1
		i = i + 1
	End If

	If StrComp(trn.Title,"Analyst3DService") = 0 Or StrComp(trn.Title,"CoreService") = 0 Or StrComp(trn.Title,"LayoutService") = 0 Then
		delete = False
	End If

		If (delete = True) And j = 3 Then
			prj.SourceLists.Remove(i)
			i = i - 1
		ElseIf (j = 3) Then
			delete = True
		End If

Next



End Sub
