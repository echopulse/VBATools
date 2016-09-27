Sub batchXMLImport()
	Application.DisplayAlerts = False
	Application.ScreenUpdating = False

	Dim strFile As String
	Dim folderPath As String
	Dim mainWb As Workbook
	Set mainWb = ThisWorkbook
	Dim tempSheet As Worksheet

	'Prompt Select Folder
	Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
	fDialog.AllowMultiSelect = False
	intChoice = fDialog.Show

	If intChoice <> 0 Then
		folderPath = fDialog.SelectedItems(1)
		strFile = Dir(folderPath & "\*.xml")
		Do While Len(strFile) > 0
			Set tempSheet = mainWb.Sheets.Add(After:=mainWb.Sheets(mainWb.Sheets.Count))
			Set tempWb = Workbooks.OpenXML(Filename:=folderPath & "\" & strFile, LoadOption:=xlXmlLoadImportToList)
			tempWb.Sheets(1).UsedRange.Copy mainWb.Sheets(tempSheet.Name).Range("A1")
			tempWb.Saved = True
			tempWb.Close savechanges:=False
			strFile = Dir
		Loop
	End If
	Application.DislpayAlerts = True
	Application.ScreenUpdating = True
End Sub

