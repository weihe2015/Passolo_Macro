Option Explicit

Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If

Dim i As Integer

Dim trn As PslTransList

Dim js As Boolean

For Each trn In prj.TransLists

	If InStr(trn.Title,".js") > 0 Then

		js = True

	End If


	For i = 1 To trn.StringCount

		Dim tString As PslTransString

		Set tString = trn.String(i)


		If js = True Then

			If InStr(tString.Text,Chr$(39)) > 0  Then

				tString.Text = Replace(tString.Text,Chr$(39), Chr$(92) & Chr$(39))

						If InStr(tString.Text,Chr$(92) & Chr$(92)) > 0 Then

							tString.Text = Replace(tString.Text,Chr$(92) & Chr$(92),Chr$(92))

						End If

				tString.TransList.Save

         		tString.State(pslStateTranslated) = True

         		tString.State(pslStateReview) = True

			End If

			If InStr(tString.Text,Chr$(34)) > 0  Then

				tString.Text = Replace(tString.Text,Chr$(34), Chr$(92) & Chr$(34))

						If InStr(tString.Text,Chr$(92) & Chr$(92)) > 0 Then

							tString.Text = Replace(tString.Text,Chr$(92) & Chr$(92),Chr$(92))

						End If

				tString.TransList.Save

         		tString.State(pslStateTranslated) = True

         		tString.State(pslStateReview) = True

			End If


		End If

		If InStr(tString.Text,Chr$(92) & Chr$(92)) > 0 Then

			tString.Text = Replace(tString.Text,Chr$(92) & Chr$(92),Chr$(92))

			tString.TransList.Save

			tString.State(pslStateTranslated) = True

			tString.State(pslStateReview) = True

		End If

	Next i

	js = False

Next trn

End Sub
