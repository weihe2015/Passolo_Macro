Option Explicit

Sub Main

Dim s As String

s = "??[{0}]"

Dim i As Integer

Dim code As Integer

For i = 1 To Len(s)

	code = Asc(Mid(s,i,1))

	If (code = 63) Then

		PSL.Output("This " & code)

	End If

Next i

End Sub
