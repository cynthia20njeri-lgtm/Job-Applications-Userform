Option Explicit

'populating the application date text box

Private Sub cmd_application_date_Click()
Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
TxtApplicationDate.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If

End Sub

'populating the completed assesment text box

Private Sub cmd_completed_assesment_Click()

Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
txt_completed_assesment.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If


End Sub

'populating the offer letter text box

Private Sub cmd_offer_letter_Click()

Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
txt_offer_letter.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If

End Sub

'populating the online interview text box

Private Sub cmd_online_interview_Click()

Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
txt_online_interview.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If

End Sub

'populating the physical interview text box

Private Sub cmd_physical_interview_Click()

Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
txt_physical_interview.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If

End Sub

'populating the received assesment text box

Private Sub cmd_received_assesment_Click()

Calendar.date_picked = False
Calendar.Show

If Calendar.date_picked = True Then
txt_received_assesment.Value = Format(Calendar.selected_date, "dd-mmm-yyyy")
End If


End Sub

'coding the add button

Private Sub CmdAdd_Click()

'ensuring the job id is not empty

If Me.txt_id.Value = "" Then
      txt_id.SetFocus
      MsgBox "Please input the id", vbCritical
Exit Sub
End If

'ensuring the job id is not empty

If Me.TxtJobTitle.Value = "" Then
      TxtJobTitle.SetFocus
      MsgBox "Please input the job title", vbCritical
Exit Sub
End If

'ensuring the company name is not empty

If Me.TxtCompanyName.Value = "" Then
      TxtCompanyName.SetFocus
      MsgBox "Please input the company name", vbCritical
Exit Sub
End If

'ensuring the work mode is not empty

If Me.Cmb_workmode = "" Then
      Cmb_workmode.SetFocus
      MsgBox "Please input the work mode", vbCritical
Exit Sub
End If

' ensuring only numeric values are entered in the job id text box
If Not IsNumeric(txt_id) Then
MsgBox "The job id can only be numeric", vbCritical
Exit Sub
 txt_id.SetFocus
End If

End Sub

' coding the delete button

Private Sub CmdDelete_Click()
    Call UserForm_Initialize
End Sub

'Initializing the userform

Private Sub UserForm_Initialize()


'clearing text boxes

TxtCompanyName.Value = ""
TxtJobTitle.Value = ""
txt_id.Value = ""
txt_completed_assesment.Value = ""
txt_offer_letter.Value = ""
txt_online_interview.Value = ""
txt_physical_interview.Value = ""
txt_received_assesment.Value = ""
TxtApplicationDate.Value = ""


'clearing commandbuttons

CmdAdd.Caption = "Add"
CmdDelete.Caption = "Delete"
CmdUpdate.Caption = "Update"

'Setting the cursor to be at the serial number text box

txt_id.SetFocus

'populating the work mode combo box

With Me.Cmb_workmode
        .Clear
        .AddItem ""
        .AddItem "Onsite"
        .AddItem "Remote"
        .AddItem "Hybrid"
End With

End Sub
