Sub Test()
    Dim FolderPath As String
    FolderPath = "C:\Users\LesBlack\OneDrive - Marconi LLC\Documents\JLR Invoice Extraction\June 2024" ' Replace with your folder path
    
    Dim FileNames() As String
    FileNames = ListFilesInFolder(FolderPath)
    
    Dim i As Integer
    For i = LBound(FileNames) To UBound(FileNames)
        Debug.Print FileNames(i)
    Next i
End Sub

Function ListFilesInFolder(FolderPath As String) As Variant
    ' Create an instance of the FileSystemObject
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the folder object
    Dim objFolder As Object
    Set objFolder = objFSO.GetFolder(FolderPath)
    
    ' Loop through each file in the folder
    Dim objFile As Object
    Dim FileArray() As String
    Dim i As Integer
    i = 0
    For Each objFile In objFolder.Files
        ReDim Preserve FileArray(i)
        FileArray(i) = objFile.Name
        i = i + 1
    Next objFile
    
    ' Return the array of file names
    ListFilesInFolder = FileArray
End Function

