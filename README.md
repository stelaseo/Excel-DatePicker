# Excel-DatePicker
Excel VBA Date Picker for Worksheet

![[sample.png]]

'------------------------------------------------------------------------------  
' Author : Stela H. Seo (https://github.com/stelaseo/)                           
' Project: Excel VBA Date Picker                                                 
' Date   : 09 January 2025
' Version: 1.7  

' Version History                                                                
<b>' 1.7  Jan 09, 2025  alexofrhodes : added button to return to TODAY - added time selection in increments of 15,30,60 </b>  
' 1.6  Dec 16, 2022  add an optional parameter (close on selecting a day)        
' 1.5  Dec 07, 2022  upload to github                                            
' 1.4  Nov 07, 2022  update the license to BSD                                   
' 1.3  Oct 31, 2022  use the given sheet name instead of ActiveSheet             
' 1.2  Oct 31, 2022  fix the usage comment to handle merged target cells         
' 1.1  Oct 28, 2022  delete old date picker,                                     
'                    when opening new one in a different sheet.                  
' 1.0  Oct 27, 2022  initial version.                                            
'------------------------------------------------------------------------------  




## Usage: (easy 3-step usage!)
1. Import the macro file

2. For example, add the following constant and function with your target cell(s) on your target sheet(s)
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


## Known Issues:
1. If there is a shape with the same name of this module's, the macro functions may not work correctly or crash in the middle of process in the worst case
2. If you save the file while the date picker is visible and reopen the file, the date picker does not function correctly. To fix this:
   1. remove the date picker in the sheet (selecting a non-targeted cell will automatically delete the picker)
   2. save the file, and
   3. close the file and reopen it.

-alexofrhodes : i didn't test these