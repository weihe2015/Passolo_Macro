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

Dim k As Integer

For Each trn In prj.TransLists

	If StrComp(trn.Title,"ADDataSourcesRaster") = 0 Then

	For k = 1 To trn.StringCount

		Dim tString As PslTransString

		Set tString = trn.String(k)

		If tString.State(pslStateTranslated) = False And tString.State(pslStateReadOnly) = False Then


		End If

	Next k

	End If

Next trn


End Sub
