Attribute VB_Name = "DatePicker"
Option Explicit
'------------------------------------------------------------------------------'
' Author : Stela H. Seo (https://github.com/stelaseo/)                         '
' Project: Excel VBA Date Picker
' Website: https://github.com/stelaseo/Excel-DatePicker/tree/main
' Date   : 09 January 16 2025                                                 '
' Version: 1.7                                                                 '
' Version History                                                              '
' 1.7  Dec 09, 2024  alexofrhodes - added button for TODAY and time selection in increments of 15,30,60
' 1.6  Dec 16, 2022  add an optional parameter (close on selecting a day)      '
' 1.5  Dec 07, 2022  upload to github                                          '
' 1.4  Nov 07, 2022  update the license to BSD                                 '
' 1.3  Oct 31, 2022  use the given sheet name instead of ActiveSheet           '
' 1.2  Oct 31, 2022  fix the usage comment to handle merged target cells       '
' 1.1  Oct 28, 2022  delete old date picker,                                   '
'                    when opening new one in a different sheet.                '
' 1.0  Oct 27, 2022  initial version.                                          '
'------------------------------------------------------------------------------'
'BSD 2-Clause License
'
'Copyright (c) 2022, Stela H. Seo
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are met:
'
'1. Redistributions of source code must retain the above copyright notice, this
'   list of conditions and the following disclaimer.
'
'2. Redistributions in binary form must reproduce the above copyright notice,
'   this list of conditions and the following disclaimer in the documentation
'   and/or other materials provided with the distribution.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
'FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
'DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
'CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
'OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'------------------------------------------------------------------------------'
''use example on worksheet
'Const DATEPICKER_TARGET = "A2,B5:B10" 'target cells
'
'Private Sub Worksheet_SelectionChange(ByVal target As Range)
''    If Not Application.Intersect(Range(DATEPICKER_TARGET), target) Is Nothing Then
'        DPOpen target.Parent.Name, target.address, target.Left + target.Width, target.Top
'    Else
'        DPClose
'    End If
'End Sub
'------------------------------------------------------------------------------'
Private Const DP_NAME As String = "SSDP_OBJECT"
Private Const DP_BACKGROUND As Long = &HFFF7F7F7
Private Const DP_SELECTED As Long = &HFFDAEFE2
Private Const DP_DAY_PREFIX As String = "SSDP_Calendar_DAY_"
Private Const DP_DAY_COLUMNS As Integer = 7
Private Const DP_DAY_ROWS As Integer = 6
Private Const DP_DAY_WIDTH As Single = 28
Private Const DP_DAY_HEIGHT As Single = 24
Private Const DP_DAY_IN_COLOR As Long = &HFF000000
Private Const DP_DAY_OUT_COLOR As Long = &HFF7F7F7F
Private Const DP_WEEK_PREFIX As String = "SSDP_Calendar_WEEK_"
Private Const DP_WEEK_WIDTH As Single = DP_DAY_WIDTH
Private Const DP_WEEK_HEIGHT As Single = DP_DAY_HEIGHT
Private Const DP_MONTH_PREFIX As String = "SSDP_Calendar_MONTH_"
Private Const DP_MONTH_COLUMNS As Integer = 4
Private Const DP_MONTH_ROWS As Integer = 3
Private Const DP_MONTH_WIDTH As Single = DP_DAY_WIDTH * 7 / 4
Private Const DP_MONTH_HEIGHT As Single = DP_DAY_HEIGHT * 7 / 3
Private Const DP_TITLE_NAME As String = "SSDP_Calendar_TITLE"
Private Const DP_TITLE_WIDTH As Single = DP_DAY_WIDTH * 5
Private Const DP_TITLE_HEIGHT As Single = 32
Private Const DP_SEL_BTN_TODAY As String = "SSDP_Calendar_TODAY" '<-------------- alexofrhodes
Private Const DP_SEL_BTN_PREVIOUS As String = "SSDP_Calendar_PREV"
Private Const DP_SEL_BTN_NEXT As String = "SSDP_Calendar_NEXT"
Private Const DP_SEL_BTN_WIDTH As Single = DP_DAY_WIDTH
Private Const DP_SEL_BTN_HEIGHT As Single = DP_TITLE_HEIGHT
Private Const DP_PANEL_BORDER_COLOR As Long = &HFFF7EBDE
Private Const DP_PANEL_BORDER_THICKNESS As Single = 4
Private Const DP_PANEL_MARGIN As Single = 10
Private Const DP_PANEL_NAME As String = "SSDP_Calendar_Panel"
Private Const DP_PANEL_WIDTH As Single = DP_DAY_WIDTH * 7 + DP_PANEL_MARGIN * 2
Private Const DP_PANEL_HEIGHT As Single = DP_TITLE_HEIGHT + DP_DAY_HEIGHT * 7 + DP_PANEL_MARGIN * 2
Private dpCloseOnDaySelection As Boolean
Private dpSheetName As String
Private firstDayOfCalendar As Date
Private firstDayOfMonth As Date
Private lastDayOfCalendar As Date
Private lastDayOfMonth As Date
Private selectedDate As Date
Private selectMonth As Boolean
Private targetAddress As String
Private targetMonth As Integer
Private targetYear As Integer
'<-------------- alexofrhodes
Private Const DP_DEFAULT_TIMEFORMAT As String = "dd/mm/yyy hh:nn"
Private Const DP_TIME_INCREMENT As Integer = 60 ' 15-30-60 - Time increment in minutes, 15-30-60
Private Const DP_TIME_COLUMN_WIDTH As Single = DP_DAY_WIDTH * 1.5
Private Const DP_TIME_BUTTON_HEIGHT As Single = 20
Private Const DP_TIME_PREFIX As String = "SSDP_Time_"
Private Const DP_NO_TIME_BUTTON As String = "SSDP_No_Time"
'<-------------- alexofrhodes
Public Function DPClose()
    DPCloseNonActiveSheets dpSheetName
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" Then
        Dim dp As Variant
        Set dp = DPFind(Sheets(dpSheetName))
        If Not dp Is Nothing Then
            dp.Delete
        End If
    End If
End Function
Public Function DPOpen(ByVal sheetName As String, ByVal address As String, ByVal x As Single, ByVal y As Single, Optional ByVal closeOnDaySelection As Boolean = False)
    DPCloseNonActiveSheets sheetName
    x = WorksheetFunction.Max(0, x)
    y = WorksheetFunction.Max(0, y)
    dpCloseOnDaySelection = closeOnDaySelection
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" Then
        Dim dp As Variant
        Set dp = DPFind(Sheets(dpSheetName))
        If dp Is Nothing Then
            Set dp = DPCreate(x, y)
        Else
            dp.Left = x
            dp.Top = y
            dp.Visible = True
        End If
        selectedDate = -1
        selectMonth = False
        targetAddress = address
        If IsDate(Sheets(dpSheetName).Range(address).value) Then
            selectedDate = Sheets(dpSheetName).Range(address).value
            DPCalculateTarget selectedDate
        Else
            DPCalculateTarget Now()
        End If
        DPUpdate
    End If
End Function
Private Function DPCreate(ByVal initialX As Single, ByVal initialY As Single) As Variant
    ' Determine the total number of buttons based on increment
    '<--- alexofrhodes
    Dim totalButtons As Integer
    Select Case DP_TIME_INCREMENT
        Case 15
            totalButtons = 12 * 4 * 2 ' 12 hours, 4 slots/hour, 2 periods (AM/PM)
        Case 30
            totalButtons = 12 * 2 * 2 ' 12 hours, 2 slots/hour, 2 periods (AM/PM)
        Case 60
            totalButtons = 12 * 1 * 2 ' 12 hours, 1 slot/hour, 2 periods (AM/PM)
        Case Else
            MsgBox "Invalid increment! Please enter 15 or 30 or 60.", vbExclamation
            DPClose
    End Select
    ' Calculate total array size
    Dim staticComponents As Integer
    staticComponents = 1 + 1 + 3 + 2 + DP_DAY_COLUMNS + DP_MONTH_COLUMNS * DP_MONTH_ROWS + DP_DAY_COLUMNS * DP_DAY_ROWS + 1
    '<--- alexofrhodes


    ReDim items(0 To staticComponents + totalButtons)
    Dim currentHour As Integer
    Dim currentMinute As Integer
    Dim i As Integer
    Dim x As Single
    Dim y As Single
    Dim targetSheet As Worksheet
    Set targetSheet = Sheets(dpSheetName)
    x = initialX
    y = initialY
    i = 0
    items(i) = DP_PANEL_NAME
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_PANEL_WIDTH, DP_PANEL_HEIGHT)
        .Name = DP_PANEL_NAME
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = DP_PANEL_BORDER_COLOR
        .Line.Weight = DP_PANEL_BORDER_THICKNESS
    End With
    x = initialX + DP_PANEL_MARGIN
    y = initialY + DP_PANEL_MARGIN
    i = i + 1
    items(i) = DP_TITLE_NAME
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_TITLE_WIDTH, DP_TITLE_HEIGHT)
        .Name = DP_TITLE_NAME
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .TextFrame.Characters.Font.Color = &H0
        .TextFrame.Characters.Text = "MONTH YEAR"
        .TextFrame.MarginTop = 0
        .TextFrame.MarginBottom = 0
        .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignLeft
        .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
        .OnAction = "'DPClickNavigate 0'"
    End With

    '<--- alexofrhodes
    x = x + DP_TITLE_WIDTH
    i = i + 1
    items(i) = DP_SEL_BTN_TODAY
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x - 2 * DP_SEL_BTN_WIDTH, y, DP_SEL_BTN_WIDTH, DP_SEL_BTN_HEIGHT)
        .Name = DP_SEL_BTN_TODAY
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickToday'"
    End With
    i = i + 1
    items(i) = DP_SEL_BTN_TODAY + "_LABEL"
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartMerge, x - 2 * DP_SEL_BTN_WIDTH + DP_SEL_BTN_WIDTH / 2 - DP_SEL_BTN_WIDTH / 6, y + 8, DP_SEL_BTN_WIDTH / 1.5, DP_SEL_BTN_HEIGHT - 16)
        .Name = DP_SEL_BTN_TODAY + "_LABEL"
        .Fill.ForeColor.RGB = &HFF000000
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickToday'"
    End With
    '<--- alexofrhodes

    i = i + 1
    items(i) = DP_SEL_BTN_PREVIOUS
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_SEL_BTN_WIDTH, DP_SEL_BTN_HEIGHT)
        .Name = DP_SEL_BTN_PREVIOUS
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickNavigate -1'"
    End With
    i = i + 1
    items(i) = DP_SEL_BTN_PREVIOUS + "_LABEL"
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeChevron, x + DP_SEL_BTN_WIDTH / 2 - DP_SEL_BTN_WIDTH / 6, y + 8, DP_SEL_BTN_WIDTH / 3, DP_SEL_BTN_HEIGHT - 16)
        .Name = DP_SEL_BTN_PREVIOUS + "_LABEL"
        .Fill.ForeColor.RGB = &HFF000000
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickNavigate -1'"
        .Flip (MsoFlipCmd.msoFlipHorizontal)
    End With
    x = x + DP_SEL_BTN_WIDTH
    i = i + 1
    items(i) = DP_SEL_BTN_NEXT
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_SEL_BTN_WIDTH, DP_SEL_BTN_HEIGHT)
        .Name = DP_SEL_BTN_NEXT
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickNavigate 1'"
    End With
    i = i + 1
    items(i) = DP_SEL_BTN_NEXT + "_LABEL"
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeChevron, x + DP_SEL_BTN_WIDTH / 2 - DP_SEL_BTN_WIDTH / 6, y + 8, DP_SEL_BTN_WIDTH / 3, DP_SEL_BTN_HEIGHT - 16)
        .Name = DP_SEL_BTN_NEXT + "_LABEL"
        .Fill.ForeColor.RGB = &HFF000000
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = False
        .OnAction = "'DPClickNavigate 1'"
    End With
   Dim row As Integer
    Dim col As Integer
    Dim value As Integer
    x = initialX + DP_PANEL_MARGIN
    y = initialY + DP_PANEL_MARGIN + DP_TITLE_HEIGHT
    value = 1
    For row = 1 To DP_MONTH_ROWS
        For col = 1 To DP_MONTH_COLUMNS
            i = i + 1
            items(i) = DP_MONTH_PREFIX + CStr(row) + CStr(col)
            With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_MONTH_WIDTH, DP_MONTH_HEIGHT)
                .Name = DP_MONTH_PREFIX + CStr(row) + CStr(col)
                .Fill.ForeColor.RGB = DP_BACKGROUND
                .Line.ForeColor.RGB = &H0
                .Line.Weight = 0
                .Line.Visible = False
                .TextFrame.Characters.Font.Color = &H0
                .TextFrame.Characters.Text = MonthName(CStr(value), True)
                .TextFrame.MarginLeft = 0
                .TextFrame.MarginRight = 0
                .TextFrame.MarginTop = 0
                .TextFrame.MarginBottom = 0
                .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignCenter
                .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
                .OnAction = "'DPClickMonth " + CStr(value) + "'"
            End With
            value = value + 1
            x = x + DP_MONTH_WIDTH
        Next col
        x = initialX + DP_PANEL_MARGIN
        y = y + DP_MONTH_HEIGHT
    Next row
    x = initialX + DP_PANEL_MARGIN
    y = initialY + DP_PANEL_MARGIN + DP_TITLE_HEIGHT
    For col = 1 To DP_DAY_COLUMNS
        i = i + 1
        items(i) = DP_WEEK_PREFIX + CStr(col)
        With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_WEEK_WIDTH, DP_WEEK_HEIGHT)
            .Name = DP_WEEK_PREFIX + CStr(col)
            .Fill.ForeColor.RGB = DP_BACKGROUND
            .Line.ForeColor.RGB = &H0
            .Line.Weight = 0
            .Line.Visible = False
            .TextFrame.Characters.Font.Color = &H0
            .TextFrame.Characters.Text = WeekdayName(col, True)
            .TextFrame.MarginLeft = 0
            .TextFrame.MarginRight = 0
            .TextFrame.MarginTop = 0
            .TextFrame.MarginBottom = 0
            .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignCenter
            .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
            .OnAction = "'DPClickWeekDay " + CStr(col) + "'"
        End With
        x = x + DP_WEEK_WIDTH
    Next col
    x = initialX + DP_PANEL_MARGIN
    y = y + DP_WEEK_HEIGHT
    For row = 1 To DP_DAY_ROWS
        For col = 1 To DP_DAY_COLUMNS
            i = i + 1
            items(i) = DP_DAY_PREFIX + CStr(row) + CStr(col)
            With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_DAY_WIDTH, DP_DAY_HEIGHT)
                .Name = DP_DAY_PREFIX + CStr(row) + CStr(col)
                .Fill.ForeColor.RGB = DP_BACKGROUND
                .Line.ForeColor.RGB = &H0
                .Line.Weight = 0
                .Line.Visible = False
                .TextFrame.Characters.Font.Color = &H0
                .TextFrame.Characters.Text = CStr(row) + CStr(col)
                .TextFrame.MarginLeft = 0
                .TextFrame.MarginRight = 0
                .TextFrame.MarginTop = 0
                .TextFrame.MarginBottom = 0
                .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignCenter
                .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
                .OnAction = "'DPClickDay " + CStr(row) + ", " + CStr(col) + "'"
            End With
            x = x + DP_DAY_WIDTH
        Next col
        x = initialX + DP_PANEL_MARGIN
        y = y + DP_DAY_HEIGHT
    Next row

'<--- alexofrhodes
    ' Time selection setup
    Dim timeRow As Integer
    Dim time As Date
    Dim totalTimeSlots As Integer
    Dim incrementMinutes As Integer
    Dim timeColumns As Integer
    Dim rowsPerColumn As Integer
    Dim currentColumn As Integer
    incrementMinutes = DP_TIME_INCREMENT ' Time increment in minutes (15 or 30)
    totalTimeSlots = 24 * 60 / incrementMinutes ' Total slots for a day based on increment
    ' Determine the number of columns
    Select Case incrementMinutes
        Case 15
            timeColumns = 8
        Case 30
            timeColumns = 4
        Case Else
            timeColumns = 2 ' Default to 2 columns for other increments
    End Select
    rowsPerColumn = totalTimeSlots / timeColumns ' Rows per column
    x = initialX + DP_PANEL_WIDTH + DP_PANEL_MARGIN ' Start X position
    y = initialY + DP_PANEL_MARGIN ' Start Y position
    time = TimeSerial(0, 0, 0) ' Start at midnight
    For timeRow = 1 To totalTimeSlots
        i = i + 1
        items(i) = DP_TIME_PREFIX & CStr(timeRow)
        With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_TIME_COLUMN_WIDTH, DP_TIME_BUTTON_HEIGHT)
            .Name = DP_TIME_PREFIX & CStr(timeRow)
            .Fill.ForeColor.RGB = DP_BACKGROUND
            .Line.ForeColor.RGB = &H0
            .Line.Weight = 0
            .Line.Visible = True
            .TextFrame.Characters.Font.Color = &H0
            .TextFrame.Characters.Text = Format(time, "hh:nn")
            .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignCenter
            .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
            .OnAction = "'DPClickTime """ & Format(time, "hh:nn") & """'"
        End With
        ' Increment time and adjust position
        time = DateAdd("n", incrementMinutes, time)
        y = y + DP_TIME_BUTTON_HEIGHT ' Move to the next row
        ' Move to the next column if necessary
        If (timeRow Mod rowsPerColumn = 0) Then
            y = initialY + DP_PANEL_MARGIN ' Reset Y position
            x = x + DP_TIME_COLUMN_WIDTH + DP_PANEL_MARGIN ' Move to the next column
        End If
    Next timeRow
    ' Add "No Time" button
    x = initialX + DP_PANEL_WIDTH + DP_PANEL_MARGIN
    y = y + DP_PANEL_MARGIN
    i = i + 1
    items(i) = DP_NO_TIME_BUTTON
    With targetSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, DP_TIME_COLUMN_WIDTH * 2 + DP_PANEL_MARGIN, DP_TIME_BUTTON_HEIGHT)
        .Left = targetSheet.Shapes(DP_TIME_PREFIX & "1").Left
        .Top = targetSheet.Shapes(DP_TIME_PREFIX & "1").Top - targetSheet.Shapes(DP_TIME_PREFIX & "1").Height - DP_PANEL_MARGIN
        .Name = DP_NO_TIME_BUTTON
        .Fill.ForeColor.RGB = DP_BACKGROUND
        .Line.ForeColor.RGB = &H0
        .Line.Weight = 0
        .Line.Visible = True
        .TextFrame.Characters.Font.Color = &H0
        .TextFrame.Characters.Text = "No Time"
        .TextFrame.HorizontalAlignment = XlHAlign.xlHAlignCenter
        .TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter
        .OnAction = "'DPClickNoTime'"
    End With
'<--- alexofrhodes

    ' Group shapes and return
    Dim dp As Variant
    Set dp = targetSheet.Shapes.Range(items).Group
    dp.Name = DP_NAME
    Set DPCreate = dp
End Function

'<-------------- alexofrhodes
Public Function DPClickTime(ByVal selectedTime As String)
    Debug.Print "Selected Time: " & selectedTime
    With Sheets(dpSheetName).Range(targetAddress)
        .NumberFormat = "dd/mm/yyyy hh:mm"
        .value = DateSerial(Year(selectedDate), _
                            Month(selectedDate), _
                            Day(selectedDate)) _
                            + CDate(selectedTime)
    End With
    If dpCloseOnDaySelection Then
        DPClose
    Else
        DPUpdate
    End If
End Function
Public Function DPClickNoTime()
    Debug.Print "No Time Selected"
    With Sheets(dpSheetName).Range(targetAddress)
        .NumberFormat = "dd/mm/yyyy"
        .value = selectedDate
    End With
    If dpCloseOnDaySelection Then
        DPClose
    Else
        DPUpdate
    End If
End Function
Public Function DPClickToday()
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" And Not DPFind(Sheets(dpSheetName)) Is Nothing Then
        Debug.Print "Select Today"
        selectedDate = Date
        With Sheets(dpSheetName).Range(targetAddress)
            .NumberFormat = "dd/mm/yyyy"
            .value = selectedDate
        End With
        DPCalculateTarget selectedDate
        DPUpdate
    End If
End Function

Public Function DPClickDay(ByVal row As Integer, ByVal col As Integer)
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" And Not DPFind(Sheets(dpSheetName)) Is Nothing Then
        selectedDate = DateAdd("d", DP_DAY_COLUMNS * (row - 1) + (col - 1), firstDayOfCalendar)
        Debug.Print "Select " + CStr(selectedDate)
        With Sheets(dpSheetName).Range(targetAddress)
            .NumberFormat = "dd/mm/yyyy"
            .value = selectedDate
        End With
        If dpCloseOnDaySelection Then
            DPClose
        Else
            DPUpdate
        End If
    End If
End Function
'<-------------- alexofrhodes

Public Function DPClickMonth(ByVal value As Integer)
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" And Not DPFind(Sheets(dpSheetName)) Is Nothing Then
        Debug.Print "Select " + MonthName(value)
        selectMonth = False
        DPCalculateTarget DateSerial(targetYear, value, 1)
        DPUpdate
    End If
End Function
Public Function DPClickNavigate(ByVal value As Integer)
    If Not IsEmpty(dpSheetName) And dpSheetName <> "" And Not DPFind(Sheets(dpSheetName)) Is Nothing Then
        If value = 0 Then
            Debug.Print "Toggle between the month view and the day view"
            selectMonth = Not selectMonth
        Else
            If selectMonth Then
                Debug.Print "Navigate year " + CStr(value)
                DPCalculateTarget DateSerial(targetYear + value, targetMonth, 1)
            Else
                Debug.Print "Navigate month " + CStr(value)
                DPCalculateTarget DateAdd("m", value, DateSerial(targetYear, targetMonth, 1))
            End If
        End If
        DPUpdate
    End If
End Function
Public Function DPClickWeekDay(ByVal value As Integer)
    'Do Nothing
End Function
Private Function DPCalculateTarget(ByVal targetDate As Date)
    targetDate = WorksheetFunction.Max(DateSerial(1901, 1, 1), WorksheetFunction.Min(DateSerial(9998, 12, 31), targetDate))
    targetYear = Year(targetDate)
    targetMonth = Month(targetDate)
    firstDayOfMonth = DateSerial(targetYear, targetMonth, 1)
    lastDayOfMonth = DateAdd("d", -1, DateAdd("m", 1, firstDayOfMonth))
    firstDayOfCalendar = DateAdd("d", -(Weekday(firstDayOfMonth) - 1), firstDayOfMonth)
    lastDayOfCalendar = DateAdd("d", DP_DAY_COLUMNS * DP_DAY_ROWS - 1, firstDayOfCalendar)
End Function
Private Function DPCloseNonActiveSheets(ByVal sheetName As String)
    Dim dp As Variant
    If dpSheetName <> sheetName Then
        On Error Resume Next
        Set dp = DPFind(Sheets(dpSheetName))
        If Not dp Is Nothing Then
            dp.Delete
        End If
    End If
    dpSheetName = sheetName
End Function
Private Function DPFind(targetSheet As Worksheet) As Variant
    Dim oShape As Shape
    For Each oShape In targetSheet.Shapes
        If oShape.Name = DP_NAME Then
            Set DPFind = oShape
            Exit Function
        End If
    Next oShape
    Set DPFind = Nothing
End Function
Private Function DPUpdate()
    Dim row As Integer
    Dim col As Integer
    Dim panelHeight As Single
    panelHeight = DP_PANEL_HEIGHT
    Dim targetSheet As Worksheet
    Set targetSheet = Sheets(dpSheetName)
    For row = 1 To DP_MONTH_ROWS
        For col = 1 To DP_MONTH_COLUMNS
            targetSheet.Shapes(DP_MONTH_PREFIX + CStr(row) + CStr(col)).Visible = selectMonth
        Next col
    Next row
    For col = 1 To DP_DAY_COLUMNS
        targetSheet.Shapes(DP_WEEK_PREFIX + CStr(col)).Visible = Not selectMonth
    Next col
    If selectMonth Then
        targetSheet.Shapes(DP_TITLE_NAME).TextFrame.Characters.Text = CStr(targetYear)
        For row = 1 To DP_DAY_ROWS
            For col = 1 To DP_DAY_COLUMNS
                targetSheet.Shapes(DP_DAY_PREFIX + CStr(row) + CStr(col)).Visible = False
            Next col
        Next row
    Else
        Dim currentDay As Date
        currentDay = firstDayOfCalendar
        targetSheet.Shapes(DP_TITLE_NAME).TextFrame.Characters.Text = MonthName(targetMonth) + " " + CStr(targetYear)
        For row = 1 To DP_DAY_ROWS
            If currentDay <= lastDayOfMonth Then
                For col = 1 To DP_DAY_COLUMNS
                    With targetSheet.Shapes(DP_DAY_PREFIX + CStr(row) + CStr(col))
                        .TextFrame.Characters.Text = CStr(Day(currentDay))
                        .Visible = True
                        If currentDay = selectedDate Then
                            .Fill.ForeColor.RGB = DP_SELECTED
                        Else
                            .Fill.ForeColor.RGB = DP_BACKGROUND
                        End If
                        If firstDayOfMonth <= currentDay And currentDay <= lastDayOfMonth Then
                            .TextFrame.Characters.Font.Color = DP_DAY_IN_COLOR
                        Else
                            .TextFrame.Characters.Font.Color = DP_DAY_OUT_COLOR
                        End If
                    End With
                    currentDay = DateAdd("d", 1, currentDay)
                Next col
            Else
                For col = 1 To DP_DAY_COLUMNS
                    targetSheet.Shapes(DP_DAY_PREFIX + CStr(row) + CStr(col)).Visible = False
                    currentDay = DateAdd("d", 1, currentDay)
                Next col
                panelHeight = panelHeight - DP_DAY_HEIGHT
            End If
        Next row
    End If
    targetSheet.Shapes(DP_PANEL_NAME).Height = panelHeight
End Function





