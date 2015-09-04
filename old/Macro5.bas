Option Explicit

Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

If prj Is Nothing Then
	MsgBox("No active Passolo Project")
	Exit Sub
End If

Dim trn As PslTransList
Dim src As PslSourceList
Dim i,j,k As Integer
i = 0
Dim encoding As Integer


For Each trn In prj.TransLists
	If trn.Language.LangCode = "chs" Then
		For i = 1 To trn.StringCount
			Dim tString As PslTransString
			Set tString = trn.String(i)
			If tString.Number = 106 Then
			If (tString.State(pslStateReadOnly) = False And tString.State(pslStateLocked) = False ) Then
				Dim transText As String
				transText = tString.Text
				For j = 1 To Len(transText)
					encoding = AscW(Mid(transText,j,1))
				Next
			End If
			End If

		Next
	End If

Next

For Each src In prj.SourceLists

	For j = 1 To src.StringCount

	Dim sString As PslSourceString

	Set sString = src.String(j)

	If sString.Number = 17 Then

		Dim sText As String

		sText = sString.Text

		For k = 1 To Len(sText)

			encoding = Asc(Mid(sText,k,1))

			If encoding > 127 Then

			End If

		Next

	End If

	Next j

Next

End Sub
