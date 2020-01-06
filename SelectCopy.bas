REM  *****  BASIC  *****

Option Explicit
' Action : Selects and copies data from rows A2 to last nonempty row and columns A to z

Sub SelectCopy()

' Define variables
Dim lastRow As Long ' Use As Long to avoid overflow (handles values up to 2,147,483,647)
Dim filename1 As String
Dim sheet1 As Worksheet

filename1 = ActiveSheet.Name
Set sheet1 = Sheets(filename1)

' Finds the last nonempty row in sheet1
lastRow = sheet1.Cells.Find(What:="*", After:== Range("A1"), _
SearchOrder:=xlByRows , SearchDirection:=xlPrevious).Row ' finds the last nonempty row
' You may also use lastRow = ActiveSheet . Range( "A" & Rows .Count) . End(xlUp).Row

' Select and Copy from rows A2 to lastRow and columns A to Z
Range("A2:Z" & lastRow).Select
Range("A2:Z" & lastRow).Copy
MsgBox "Copy and Selected " & lastRow - 1 & "row(s) from" & filename1
End Sub
' Source:
' https://trumpexcel.com/vba-ranges/
