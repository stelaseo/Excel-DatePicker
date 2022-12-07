# Excel-DatePicker
Excel VBA Date Picker



## Known Issues:
1. If there is a shape with the same name of this module's, the macro functions may not work correctly or crash in the middle of process in the worst case


2. If you save the file while the date picker is visible and reopen the file, the date picker does not function correctly. To fix this:
   1. remove the date picker in the sheet (selecting a non-targeted cell will automatically delete the picker)
   2. save the file, and
   3. close the file and reopen it.



## Usage: (easy 3-step usage!)
1. Import the macro file


2. add the following constant and function with your target cell(s) on your target sheet(s)
```
Const DATEPICKER_TARGET = "A2,B5:B10" 'target cells

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If Not Application.Intersect(Range(DATEPICKER_TARGET), Range(target.address)) Is Nothing Then
        DatePicker.DPOpen target.Parent.Name, target.address, target.Left + target.Width, target.Top
    Else
        DatePicker.DPClose
    End If
End Sub
```


3. add the following function to ThisWorkbook to prevent one of the known issues
```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    DatePicker.DPClose
End Sub
```
