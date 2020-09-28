Sub GetFileList()
'Mengimport nama file dari folder ke Exel'
	Dim xFSO As Object
	Dim xFolder As Object
	Dim xFile As Object
	Dim xFiDialog As FileDialog
	Dim xPath As String
	Dim i As Integer
	Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
	If xFiDialog.Show = -1 Then
		xPath = xFiDialog.SelectedItems(1)
	End If
	Set xFiDialog = Nothing
	If xPath = "" Then Exit Sub
	Set xFSO = CreateObject("Scripting.FileSystemObject")
	Set xFolder = xFSO.GetFolder(xpath)
	ActiveSheet.Cells(1, 1) = "File Name"
	i = 1
	For Each xFile In xFolder.Files
		i = i + 1
		ActiveSheet.Cells(i, 1) = Left(xFile.Name, InStrRev(xFile.Name, ".") -1)
	Next
End Sub