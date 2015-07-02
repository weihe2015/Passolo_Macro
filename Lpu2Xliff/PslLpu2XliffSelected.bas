'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\WINDOWS\system32\msxml6.dll#Microsoft XML, v6.0
''Export files to XLIFF document for processing in a TM system
' -------------------------------
' Version 1.0 Wei He August 13 2014.
' -------------------------------
' Macro executes on:
' Projects		Active/opened translation bundles.
' Only Export those selected source
' Source lists	N/A
' Target lists	ALL
'
' Output: It will export all the selected source in all status
' The export xliff files will be saved in the same folder of the lpu file
' ===============================


Option Explicit

Sub Main
    Const ForWriting = 2
	Dim XLIFFDoc As DOMDocument
	Dim XLIFFElem As IXMLDOMElement
	Dim newFileNode As IXMLDOMElement
	Dim newStringNode As IXMLDOMElement
	Dim newSourceString As IXMLDOMElement
	Dim newTargetString As IXMLDOMElement
	Dim newNoteString As IXMLDOMElement
	Dim rdr As New SAXXMLReader
	Dim wrt As New MXXMLWriter
	Dim XLIFFfile As Object
	Dim FileStringCount As Integer


Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If


Dim Lang As String
Dim trn As PslTransList
Dim i, j As Long
Dim selectedItem As Long
selectedItem = 0


Dim fso, fso1, fso2, objFso, objFile, MyFile, FileName, f, f1, f2
Dim Path As String
Dim Path1 As String

Set fso = CreateObject("Scripting.FileSystemObject")

    Path = prj.Location + "\" + prj.Name

    If (fso.FolderExists(Path)) = False Then

    Set f = fso.CreateFolder(Path)

    End If

Set objFso = CreateObject("Scripting.FileSystemObject")

Dim logFile As String

logFile = Path + "\log.txt"

Set objFile = objFso.CreateTextFile(logFile,True)

PSL.Output ("Preparing XLIFF data...")

For Each trn In prj.TransLists

    If trn.Selected = True Then

    selectedItem = selectedItem + 1

    If StrComp(Lang, trn.Language.LangCode) <> 0 Then

    Lang = LCase(PSL.GetLangCode(trn.Language.LangID,pslCodeTrados))

    Path = prj.Location + "\" + prj.Name + "\" + Lang

      Set fso1 = CreateObject("Scripting.FileSystemObject")

        If (fso1.FolderExists(Path)) = False Then

           Set f1 = fso1.CreateFolder(Path)

      End If

    End If

    Dim targetName As String

    targetName = trn.Title

    If InStr(targetName, Chr$(92)) <> 0 Then

     Set fso2 = CreateObject("Scripting.FileSystemObject")

     Dim Path2 As String

     Path2 = StrReverse( Mid (StrReverse(targetName), InStr(StrReverse(targetName),Chr$(92)) + 1))

         If(fso2.FolderExists(Path + "\" + Path2)) = False Then

         	Set f2 = fso2.CreateFolder(Path + "\" + Path2)

         	Path = Path + "\" + Path2

         targetName = Mid(targetName, InStr(targetName, Chr$(92)) + 1)

         'targetName = Replace(targetName, Chr$(46), "_")

         End If

    End If

    FileName = Path + "\" + targetName + ".xliff"

    Dim sourceLang As String

    sourceLang = LCase(PSL.GetLangCode(trn.SourceList.LangID,pslCodeTrados))

    ' Create the XLIFF object

    Set XLIFFDoc=CreateObject("Msxml2.DOMDocument.6.0")

    XLIFFDoc.async=False
    XLIFFDoc.preserveWhiteSpace=True

    ' Create the root node.

    Set XLIFFElem = XLIFFDoc.createNode(1,"xliff","")
    XLIFFElem.setAttribute("version","1.2")
    XLIFFDoc.documentElement = XLIFFElem

    Set newFileNode = XLIFFDoc.createNode(1,"file","")
    newFileNode.setAttribute("original",StrReverse(Left(StrReverse(trn.TargetFile),(InStr(StrReverse(trn.TargetFile), "\") - 1))))
	newFileNode.setAttribute("source-language",sourceLang)
	newFileNode.setAttribute("target-language",Lang)
	newFileNode.setAttribute("datatype","plaintext")

    XLIFFElem.appendChild(newFileNode)

    PSL.Output("Exporting " & trn.Title & " " & Lang)

    On Error GoTo nextStringList

	For i = 1 To trn.StringCount

      Dim tString As PslTransString

      Set tString = trn.String(i)

      Dim translatable As Boolean

      translatable = Not tString.State(pslStateReadOnly)

      If (tString.SourceText <> "") And translatable Then

      Set newStringNode = XLIFFDoc.createNode(1,"trans-unit","")
      newStringNode.setAttribute("Number",tString.Number)
	  newStringNode.setAttribute("id",tString.ID)
	  newFileNode.appendChild(newStringNode)

      Set newSourceString = XLIFFDoc.createNode(1,"source","")
      newSourceString.Text = tString.SourceText
      newStringNode.appendChild(newSourceString)

      Set newTargetString = XLIFFDoc.createNode(1,"target","")
      newTargetString.Text = tString.Text
      newStringNode.appendChild(newTargetString)

      Set newNoteString = XLIFFDoc.createNode(1,"note","")
      newNoteString.Text = tString.Comment
      newStringNode.appendChild(newNoteString)

	  On Error GoTo nextStringList

      End If

GoTo NextString

nextStringList:

	If Err.Description <> "" Then
		PSL.Output("Skipping " & tString.Number & " " & tString.ID & " " & Err.Description)
		objFile.Write "Skipping " & tString.Number & " " & tString.ID & " " & Err.Description
        Err.Clear
    End If

nextString:

	Next i

wrt.byteOrderMark = False
wrt.omitXMLDeclaration = False
wrt.indent = True
wrt.standalone=True
wrt.encoding= "UTF-8"

'Set the XML writer to the SAX content handler.
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
'Parse the DOMDocument object.
rdr.parse XLIFFDoc

'Open the file for writing

Set XLIFFfile = fso.OpenTextFile(FileName, ForWriting, True, True)
XLIFFfile.Write (wrt.Output)
Set wrt = Nothing
XLIFFfile.Close

   End If

Next trn

If selectedItem = 0 Then

	MsgBox("You have not selected any string list to export, please run the macro again.")

	Exit Sub

Else

	Done:
	objFile.Close
	PSL.Output ("--- Done ---")
	MsgBox ("Successfully exporting Xliff to the same folder which contains the lpu file!")

End If

End Sub


