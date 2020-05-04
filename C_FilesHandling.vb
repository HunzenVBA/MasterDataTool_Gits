Option Explicit

Public Function dictFilesInWorkbookFolder(DataFolder As String) As Dictionary

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Dim dictReturn As Dictionary

Set dictReturn = New Dictionary
Set oFSO = CreateObject("Scripting.FileSystemObject")

DataFolder = ThisWorkbook.Path
Set oFolder = oFSO.GetFolder(DataFolder)

    For Each oFile In oFolder.Files
        dictReturn.Add oFile, i
    '    ThisWorkbook.Worksheets("Test").Cells(i + 1, 1) = oFile.Name
         i = i + 1
    Next oFile

Set dictFilesInWorkbookFolder = dictReturn

End Function

Public Function collFilesInWorkbookFolder(DataFolder As String) As Collection

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Dim dictReturn As Collection

Set dictReturn = New Collection
Set oFSO = CreateObject("Scripting.FileSystemObject")

DataFolder = ThisWorkbook.Path
Set oFolder = oFSO.GetFolder(DataFolder)

    For Each oFile In oFolder.Files
        dictReturn.Add oFile
    '    ThisWorkbook.Worksheets("Test").Cells(i + 1, 1) = oFile.Name
         i = i + 1
    Next oFile

Set collFilesInWorkbookFolder = dictReturn

End Function

Sub testDictOfFiles()
Dim vardict As Variant
Dim dict As Dictionary


Set dict = dictFilesInWorkbookFolder(strDataFolder)

For Each vardict In dict.Items()
    Debug.Print
Next vardict

For Each vardict In dict.Keys()
    Debug.Print TypeName(vardict)
    Debug.Print vardict.Name, vardict.Size
Next vardict

End Sub

Function collGetCollectionofFiles() As Collection
Set collGetCollectionofFiles = collFilesInWorkbookFolder(strDataFolder)
'
'For Each vardict In coll
'    Debug.Print TypeName(vardict)
'    Debug.Print vardict.Name, vardict.Size
'Next vardict

End Function

Sub tt()
Dim c As Collection

Set c = collGetCollectionofFiles

Debug.Print c.Count
End Sub
