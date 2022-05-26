VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "frmCalendar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'Variables

Private Const calendarConst As Integer = 2  '1 = SUA; 2 = EUR

Private currentDate As Date
Private selectedDate As Date

Private Const redColor As Long = &H8080FF
Private Const greenColor As Long = &H80FF80
Private Const grayColor As Long = &H80000004
Private Const whiteColor As Long = &HFFFFFF
Private Const purpleColor As Long = &H800080
Private Const blackColor As Long = &H80000012


Private Sub UserForm_Click()

End Sub



'----------------------------------------------------------------------------------------
'Commands


Private Sub UserForm_Initialize()

'Design in EUR/SUA style.
If calendarConst = 2 Then

    WeekDay_1.Caption = "M"
    WeekDay_2.Caption = "T"
    WeekDay_3.Caption = "W"
    WeekDay_4.Caption = "T"
    WeekDay_5.Caption = "F"
    WeekDay_6.Caption = "S"
    WeekDay_7.Caption = "S"
    
    day_11.ForeColor = blackColor
    day_21.ForeColor = blackColor
    day_31.ForeColor = blackColor
    day_41.ForeColor = blackColor
    day_51.ForeColor = blackColor
    day_61.ForeColor = blackColor

    day_16.ForeColor = purpleColor
    day_26.ForeColor = purpleColor
    day_36.ForeColor = purpleColor
    day_46.ForeColor = purpleColor
    day_56.ForeColor = purpleColor
    day_66.ForeColor = purpleColor

End If

'Display today.
currentDate = Date
Label_Today.Caption = currentDate

Dim activeValue As Variant
activeValue = ActiveCell.Value

'Display custom or current month.
If IsDate(activeValue) Then
    selectedDate = activeValue
    Label_Selection.Caption = selectedDate
    Label_Month.Caption = month(selectedDate)
    Label_Year.Caption = year(selectedDate)
    
    Call RecalculateCalendar
Else
    Call CommandButton_Today_Click
End If

End Sub


Private Sub CommandButton_Today_Click()

Label_Month.Caption = month(Date)
Label_Year.Caption = year(Date)

Call RecalculateCalendar

End Sub

Private Sub Validate_Click()

'Pass selected date and close form.
If selectedDate >= 1 Then
    ActiveCell.Value = selectedDate
End If

Unload Me

End Sub



'----------------------------------------------------------------
'Month


Private Sub Label_MonthMinus_Click()

Dim monthID As Integer
monthID = CInt(Label_Month.Caption) - 1
Call DisplayMonth(monthID)

Call RecalculateCalendar

End Sub

Private Sub Label_MonthPlus_Click()

Dim monthID As Integer
monthID = CInt(Label_Month.Caption) + 1
Call DisplayMonth(monthID)

Call RecalculateCalendar

End Sub

Sub DisplayMonth(monthID As Integer)

If monthID < 1 Then
    monthID = 12
ElseIf monthID > 12 Then
    monthID = 1
End If

Label_Month.Caption = monthID

End Sub



'----------------------------------------------------------------
'Year


Private Sub Label_YearMinus_Click()

Dim yearID As Integer
yearID = CInt(Label_Year) - 1
Label_Year.Caption = yearID

Call RecalculateCalendar

End Sub

Private Sub Label_YearPlus_Click()

Dim yearID As Integer
yearID = CInt(Label_Year) + 1
Label_Year.Caption = yearID

Call RecalculateCalendar

End Sub



'----------------------------------------------------------------
'Calendar


Sub ResetCalendar()
'Reset .Caption and .BackColor

For Each element In Me.Controls
    If Left(element.Name, 4) = "day_" Then
        element.Caption = ""
        element.BackColor = grayColor
    End If
Next

End Sub

Sub RecalculateCalendar()

Dim month As Integer
Dim year As Integer
Dim currentDay, firstDay, lastDay As Date
Dim days As Integer
Dim weeks As Integer
Dim row, column As Integer

Call ResetCalendar

'Get month and year.
month = CInt(Label_Month.Caption)
year = CInt(Label_Year.Caption)

'Calculate firstDay, lastDay and number of weeks.
firstDay = DateSerial(year, month, 1)
lastDay = Application.WorksheetFunction.EoMonth(DateSerial(year, month, 1), 0)
days = lastDay - firstDay + 1
weeks = Application.WorksheetFunction.WeekNum(lastDay, calendarConst) - Application.WorksheetFunction.WeekNum(firstDay) + 1

'Do adjustments for a better display.
row = 1
If weeks = 4 Then row = 2

'Update calendar.
For i = 1 To days
    currentDay = firstDay + i - 1
    column = Application.WorksheetFunction.Weekday(currentDay, calendarConst) Mod 8
    Me.Controls("day_" & row & column).Caption = i
    If currentDay = currentDate Then
        Me.Controls("day_" & row & column).BackColor = greenColor
    End If
    If currentDay = selectedDate Then
        Me.Controls("day_" & row & column).BackColor = redColor
    End If
    If column = 7 Then
        column = 1
        row = row + 1
    End If
Next i

Call SetActiveDays

End Sub

Sub SetActiveDays()
'Enable/disable buttons based on Caption.

For Each element In Me.Controls
    If Left(element.Name, 4) = "day_" Then
        If element.Caption = "" Then
            element.Enabled = False
            element.BackColor = whiteColor
        Else
            element.Enabled = True
        End If
    End If
Next

End Sub

Sub ClearSelection()
'Clear existing selection (red cell).

For Each element In Me.Controls
    If Left(element.Name, 4) = "day_" Then
        If element.BackColor = redColor Then
            element.BackColor = grayColor
            Exit Sub
        End If
    End If
Next

End Sub


Sub SelectDate()

Dim month As Integer
Dim year As Integer
Dim day As Integer

Call ClearSelection

'Get month and year.
month = CInt(Label_Month.Caption)
year = CInt(Label_Year.Caption)
day = CInt(Me.ActiveControl.Caption)
selectedDate = DateSerial(year, month, day)
Label_Selection.Caption = selectedDate

Call RecalculateCalendar

Me.ActiveControl.BackColor = redColor

End Sub



'----------------------------------------------------------------------------------------
'Butons events.


Private Sub day_11_Click()
    Call SelectDate
End Sub

Private Sub day_12_Click()
    Call SelectDate
End Sub

Private Sub day_13_Click()
    Call SelectDate
End Sub

Private Sub day_14_Click()
    Call SelectDate
End Sub

Private Sub day_15_Click()
    Call SelectDate
End Sub

Private Sub day_16_Click()
    Call SelectDate
End Sub

Private Sub day_17_Click()
    Call SelectDate
End Sub

Private Sub day_21_Click()
    Call SelectDate
End Sub

Private Sub day_22_Click()
    Call SelectDate
End Sub

Private Sub day_23_Click()
    Call SelectDate
End Sub

Private Sub day_24_Click()
    Call SelectDate
End Sub

Private Sub day_25_Click()
    Call SelectDate
End Sub

Private Sub day_26_Click()
    Call SelectDate
End Sub

Private Sub day_27_Click()
    Call SelectDate
End Sub

Private Sub day_31_Click()
    Call SelectDate
End Sub

Private Sub day_32_Click()
    Call SelectDate
End Sub

Private Sub day_33_Click()
    Call SelectDate
End Sub

Private Sub day_34_Click()
    Call SelectDate
End Sub

Private Sub day_35_Click()
    Call SelectDate
End Sub

Private Sub day_36_Click()
    Call SelectDate
End Sub

Private Sub day_37_Click()
    Call SelectDate
End Sub

Private Sub day_41_Click()
    Call SelectDate
End Sub

Private Sub day_42_Click()
    Call SelectDate
End Sub

Private Sub day_43_Click()
    Call SelectDate
End Sub

Private Sub day_44_Click()
    Call SelectDate
End Sub

Private Sub day_45_Click()
    Call SelectDate
End Sub

Private Sub day_46_Click()
    Call SelectDate
End Sub

Private Sub day_47_Click()
    Call SelectDate
End Sub

Private Sub day_51_Click()
    Call SelectDate
End Sub

Private Sub day_52_Click()
    Call SelectDate
End Sub

Private Sub day_53_Click()
    Call SelectDate
End Sub

Private Sub day_54_Click()
    Call SelectDate
End Sub

Private Sub day_55_Click()
    Call SelectDate
End Sub

Private Sub day_56_Click()
    Call SelectDate
End Sub

Private Sub day_57_Click()
    Call SelectDate
End Sub

Private Sub day_61_Click()
    Call SelectDate
End Sub

Private Sub day_62_Click()
    Call SelectDate
End Sub

Private Sub day_63_Click()
    Call SelectDate
End Sub

Private Sub day_64_Click()
    Call SelectDate
End Sub

Private Sub day_65_Click()
    Call SelectDate
End Sub

Private Sub day_66_Click()
    Call SelectDate
End Sub

Private Sub day_67_Click()
    Call SelectDate
End Sub




