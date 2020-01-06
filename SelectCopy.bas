REM  *****  BASIC  *****

Option Explicit
'Action : Resize pasted selection for highlighting. You must use SelectCopy before using this macro .
'A popup window shows asking the user to type the number of the week to populate the selection

Sub WeekNumber()

Dim MyNumber As Integer ' declare a variable called MyNumber of type Integer

' change selection size from 26th column (Z) to 1st column (A)
Selection.Resize(Selection.Rows.Count, Selection.Columns.Count - 25).Select

' display dialog box asking user for number
MyNumber = InputBox("Enter number for the week ","Week Number ","Type your number here")

' set selection equal to same defined number
Selection.Value = MyNumber

End Sub
' Sources:
' https://stackoverflow.com/questions/10692213/excel-vba-how-to-extend-a-range-given-a-current-selection
' https://www.wiseowl.co.uk/blog/s2458/inputbox.htm
' https://www.mrexcel.com/forum/excel-questions/73371-vba-fill-selected-cells-letter.html
