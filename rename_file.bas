Sub RenameFiles()
'Updateby20141124 modif by Dek Sudiana'
    Dim xDir As String
    Dim xFile As String
    Dim xRow As Long
    With Application.FileDialog(msoFileDialogFolderPicker)
	.AllowMultiSelect = False
If .Show = -1 Then
            xDir = .SelectedItems(1)
            xFile = Dir(xDir & Application.PathSeparator & "*")
            Do Until xFile = ""
                xRow = 0
                On Error Resume Next
                xRow = Application.Match(xFile, Range("A:A"), 0)
                If xRow > 0 Then
                    Name xDir & Application.PathSeparator & xFile As _
                    xDir & Application.PathSeparator & Cells(xRow, "B").Value
                End If
                xFile = Dir
            Loop
End If
End With
End Sub