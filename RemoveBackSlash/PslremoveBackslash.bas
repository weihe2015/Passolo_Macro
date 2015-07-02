'Search and remove the extra backslash in javascript

Option Explicit
Sub Main

  Dim index As Integer

  Dim backSlash As String
  backSlash = Chr$(92)

  Dim EmptyString As String
  EmptyString = ""

  Dim trn As PslTransList
  Dim i As Long

  ' Get Passolo Project
  Dim prj As PslProject
  Set prj = PSL.ActiveProject

  ' Check whether we have open a project or not
  If prj Is Nothing Then
     MsgBox("No active Passolo project.")
     Exit Sub
  End If

  For Each trn In prj.TransLists

      For i = 1 To trn.StringCount

       Dim tString As PslTransString

       Set tString = trn.String(i)

        If InStr(tString.Text, Chr$(92)) > 0 Then

           'Remove the backslash in the translated string
            tString.Text = Replace(tString.Text, Chr$(92), EmptyString)

            tString.TransList.Save

            tString.State(pslStateTranslated) = True

        End If

       Next i

  Next trn

 MsgBox("Done!")

End Sub
