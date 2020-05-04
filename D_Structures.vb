Option Explicit

'Take a Range and Put the values into an Array of type String
Function arrCreateArrayFromColumns(rColumns As Range)
Dim arrTemparray()
Dim counter As Long
For counter = 1 To rColumns.CountLarge
    arrTemparray = rColumns
Next counter
arrCreateArrayFromColumns = arrTemparray
End Function
Function rGetColumnsAsARange(ws As Worksheet) As Range 'From A1 to lastwritten row and column
Dim r As Range
ws.Activate
Set r = ws.Range(Cells(1, 1), Cells(fLastWrittenRow(ws, 1), fLastWrittenCol(ws, 1)))
Set rGetColumnsAsARange = r
End Function
Function fNameARange(r As Range, Name As String) 'what does range.name do? Which property does it change?
r.Name = Name
End Function
Function fWriteArrayToRange(arrRangeToWriteToCells As Variant, ws As Worksheet) 'From A1 write an Array to WS with the upper boundaries of an Array
'ws.Activate
With ws
ws.Range(Cells(1, 1), Cells(UBound(arrRangeToWriteToCells, 1), UBound(arrRangeToWriteToCells, 2))) = arrRangeToWriteToCells 'write Range purely based on ubounds of a 2D array
End With
End Function
Sub TestNameRange()
Dim r As Range
Set r = rGetColumnsAsARange(wsTest1)
fNameARange r, "TestName"
r.Name = "test2"
Debug.Print r.Name

End Sub
Sub GetColumns()
Dim arrColumns As Variant
arrColumns = rGetColumnsAsARange(wsTest1) 'get all written Range of a sheet and put into a 2D array
fWriteArrayToRange arrColumns, wsTest2 'Function to Write an 2D Array to a sheet, starting from A1 until the Ubounds of the array
End Sub

'======================================================================================================================='
'===== Test section ======
'======================================================================================================================='

Sub testcol()
Debug.Print fLastWrittenCol(shtTest, 1)
End Sub
Sub TestRangeToArray()

Dim rRangeOfColumns As Range
Dim arrRangeArray As Variant

Set rRangeOfColumns = shtTest.Range(Cells(1, 1), Cells(1, 4))
arrRangeArray = rRangeOfColumns
End Sub

Sub rangetests()
ThisWorkbook.Worksheets(4).Select
ThisWorkbook.Worksheets("Test").Select
ThisWorkbook.Worksheets("Test").Range("A5:D5").Select
shtTest.Range("G1").Select
End Sub

Sub testRefToCodename()
Dim ws As Worksheet

Set ws = wsTest2
Debug.Print ws.Name
Debug.Print ws.CodeName
Debug.Print TypeName(wsTest2.CodeName)
Debug.Assert ws.Range.Name
'With wsTest2.CodeName
'    ws.Range("A4").Select
'End With
End Sub

Sub testUsedRange()
wsTest1.UsedRange.Select
End Sub
