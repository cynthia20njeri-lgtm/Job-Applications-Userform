
Option Explicit
Public selected_date As Date
Public date_picked As Boolean
Private Sub cmb_month_Change()
Call show_dates
End Sub

Private Sub cmb_year_Change()
Call show_dates
End Sub

Private Sub CommandButton1_Click()
pick_date 1
End Sub

Private Sub CommandButton2_Click()
pick_date 2
End Sub

Private Sub CommandButton3_Click()
pick_date 3
End Sub

Private Sub CommandButton4_Click()
pick_date 4
End Sub

Private Sub CommandButton5_Click()
pick_date 5
End Sub

Private Sub CommandButton6_Click()
pick_date 6
End Sub
Private Sub CommandButton7_Click()
pick_date 7
End Sub

Private Sub CommandButton8_Click()
pick_date 8
End Sub

Private Sub CommandButton9_Click()
pick_date 9
End Sub

Private Sub CommandButton10_Click()
pick_date 10
End Sub

Private Sub CommandButton11_Click()
pick_date 11
End Sub

Private Sub CommandButton12_Click()
pick_date 12
End Sub

Private Sub CommandButton13_Click()
pick_date 13
End Sub

Private Sub CommandButton14_Click()
pick_date 14
End Sub

Private Sub CommandButton15_Click()
pick_date 15
End Sub

Private Sub CommandButton16_Click()
pick_date 16
End Sub

Private Sub CommandButton17_Click()
pick_date 17
End Sub

Private Sub CommandButton18_Click()
pick_date 18
End Sub

Private Sub CommandButton19_Click()
pick_date 19
End Sub

Private Sub CommandButton20_Click()
pick_date 20
End Sub
Private Sub CommandButton21_Click()
pick_date 21
End Sub
Private Sub CommandButton22_Click()
pick_date 22
End Sub
Private Sub CommandButton23_Click()
pick_date 23
End Sub

Private Sub CommandButton24_Click()
pick_date 24
End Sub
Private Sub CommandButton25_Click()
pick_date 25
End Sub
Private Sub CommandButton26_Click()
pick_date 26
End Sub
Private Sub CommandButton27_Click()
pick_date 27
End Sub
Private Sub CommandButton28_Click()
pick_date 28
End Sub
Private Sub CommandButton29_Click()
pick_date 29
End Sub
Private Sub CommandButton30_Click()
pick_date 30
End Sub

Private Sub CommandButton31_Click()
pick_date 31
End Sub

Private Sub CommandButton32_Click()
pick_date 32
End Sub

Private Sub CommandButton33_Click()
pick_date 33
End Sub

Private Sub CommandButton34_Click()
pick_date 34
End Sub

Private Sub CommandButton35_Click()
pick_date 35
End Sub

Private Sub CommandButton36_Click()
pick_date 36
End Sub

Private Sub CommandButton37_Click()
pick_date 37
End Sub

Private Sub CommandButton38_Click()
pick_date 38
End Sub

Private Sub CommandButton39_Click()
pick_date 39
End Sub

Private Sub CommandButton40_Click()
pick_date 40
End Sub

Private Sub CommandButton41_Click()
pick_date 41
End Sub

Private Sub CommandButton42_Click()
pick_date 42
End Sub

' populating the month combo box

Private Sub UserForm_Initialize()

Dim i As Integer
For i = 1 To 12
Me.cmb_month.AddItem Format(DateSerial(2026, i, 1), "mmm")
Next i

Dim y As Integer
For y = 2020 To 2060
Me.cmb_year.AddItem y
Next y

Me.cmb_month.Value = Format(Date, "mmm")
Me.cmb_year.Value = Format(Date, "yyyy")
Me.lbl_todays_date.Caption = Format(Date, "yyyy-mmm-dd")

Call show_dates

End Sub

' m represents selected month
'y represents selected year
'fd represents first date of the month
'sd represents the starting day of the month
'dm represents the number of days in a month
'butn represents the command buttons that should display the dates

Private Sub show_dates()

Dim m As Integer
Dim y As Integer
Dim fd As Date
Dim sd As Integer
Dim dm As Integer
Dim i As Integer
Dim today As Date
Dim butn As MSForms.CommandButton

m = Me.cmb_month.ListIndex + 1
y = Me.cmb_year.ListIndex + 2020
fd = DateSerial(y, m, 1)
sd = Weekday(fd, vbSunday)
dm = Day(DateSerial(y, m + 1, 0))
today = Date

For i = 1 To 42
Set butn = Me.Controls("CommandButton" & i)
butn.Enabled = False
butn.Caption = ""

Next i


For i = 1 To dm
Set butn = Me.Controls("CommandButton" & (i + sd - 1))
butn.Caption = i
butn.Enabled = True

' highlighting todays date in a different color

If Day(today) = i And Year(today) = y And Month(today) = m Then
butn.ForeColor = vbRed
butn.BackColor = vbWhite
butn.BackStyle = fmBackStyleOpaque
Else
butn.BackStyle = fmBackStyleOpaque
butn.ForeColor = vbBlack
End If

Next i

End Sub
Private Sub pick_date(butnindex As Integer)
Dim d As Integer

d = Val(Me.Controls("CommandButton" & butnindex).Caption)
If d = 0 Then
  Exit Sub
End If

Me.selected_date = DateSerial(Me.cmb_year.Value, Me.cmb_month.ListIndex + 1, d)
Me.date_picked = True
Me.Hide

End Sub
