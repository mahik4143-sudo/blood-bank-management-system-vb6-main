VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BLOOD BANK MANAGEMENT SYSTEM"
   ClientHeight    =   11490
   ClientLeft      =   165
   ClientTop       =   1110
   ClientWidth     =   12840
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuDonor 
      Caption         =   "DONOR MANAGEMENT"
      Begin VB.Menu MDONOR_REGISTRATION_DETAIL 
         Caption         =   "DONOR REGISTRATION"
      End
      Begin VB.Menu dfrmDonationVisit 
         Caption         =   "DONOR VISIT"
      End
      Begin VB.Menu MfrmConsent 
         Caption         =   "DONOR CONSENT"
      End
      Begin VB.Menu dfrmDonarCard 
         Caption         =   "DONOR CARD"
      End
   End
   Begin VB.Menu MPATIENTMANAGEMENT 
      Caption         =   "PATIENT MANAGEMENT"
      Begin VB.Menu PATNT 
         Caption         =   "PATIENT REGISTRATION"
      End
      Begin VB.Menu REQUEST 
         Caption         =   "PATIENT REQUEST"
      End
   End
   Begin VB.Menu MBLOODPROCESSING 
      Caption         =   "BLOOD PROCESSING"
      Begin VB.Menu SAMP 
         Caption         =   "SAMPLE COLLECTION"
      End
      Begin VB.Menu TEST 
         Caption         =   "LAB TESTING"
      End
      Begin VB.Menu dfrmComponentProcessing 
         Caption         =   "COMPONENT PROCESSING"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "INVENTORY AND ISSUE"
      Begin VB.Menu dfrmBloodstock 
         Caption         =   "BLOOD STOCK"
      End
      Begin VB.Menu Mfrmblood_issue 
         Caption         =   "BLOOD ISSUE"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "REPORT"
      Begin VB.Menu MBLOOD_RANGE 
         Caption         =   "REFERENCE BLOOD RANGE REPORT"
      End
      Begin VB.Menu MREGISTRATION 
         Caption         =   "REGISTRATION REPORT"
      End
      Begin VB.Menu MVisit 
         Caption         =   "DONATION VISIT REPORT"
      End
      Begin VB.Menu MCARD 
         Caption         =   "DONAR CARD REPORT"
      End
      Begin VB.Menu mPATIENTREPORT 
         Caption         =   "PATIENT REPORT"
      End
      Begin VB.Menu MREQUEST 
         Caption         =   "PATIENT REQUEST REPORT"
      End
      Begin VB.Menu MSAMPLE 
         Caption         =   "SAMPLE REPORT"
      End
      Begin VB.Menu MTESTING 
         Caption         =   "LAB TESTING REPORT"
      End
      Begin VB.Menu MPROCESS 
         Caption         =   "COMPONENT PROCESSING REPORT"
      End
      Begin VB.Menu MISSUE 
         Caption         =   "BLOOD ISSUE REPORT"
      End
      Begin VB.Menu MSTOCK 
         Caption         =   "BLOOD STOCK REPORT"
      End
      Begin VB.Menu mnuDashboard 
         Caption         =   "LIVE DASHBOARD"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "SYSTEM AND SETTINGS"
      Begin VB.Menu MfrmRefBloodRange 
         Caption         =   "REFERENCE BLOOD RANGE"
      End
      Begin VB.Menu mnuSysUser 
         Caption         =   "ADD NEW SYSTEM USER"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-(HYPHEN FOR SEPARATOR)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu mnuChatbot 
      Caption         =   "CHAT WITH REDBOT"
   End
   Begin VB.Menu CASCADE 
      Caption         =   "CASCADE"
   End
   Begin VB.Menu MLOGOUT 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CASCADE_Click()
MDIForm1.Arrange vbCascade
End Sub

Private Sub dfrmBloodstock_Click()
frmBloodstock.Show
End Sub

Private Sub dfrmComponentProcessing_Click()
frmComponentProcessing.Show
End Sub

Private Sub dfrmDonarCard_Click()
frmDonarCard.Show
End Sub

Private Sub dfrmDonationVisit_Click()
frmDonationVisit.Show
End Sub

Private Sub MBLOOD_RANGE_Click()
Ref_blood_Report.Show
End Sub

Private Sub MCARD_Click()
Donar_card_Report.Show
End Sub

Private Sub MDIForm_Load()
    ' Software ko puri screen par kholne ke liye
    Me.WindowState = vbMaximized
    ' Optional: MDI ka background color change karne ke liye
    Me.BackColor = vbApplicationWorkspace
End Sub

Private Sub MDonation_Visit_Report_Click()
Donation_Visit_Report.Show
End Sub

Private Sub MDONOR_REGISTRATION_DETAIL_Click()
DONOR_REGISTRATION_DETAIL.Show
End Sub

Private Sub MfrmDonationVisit_Click()
frmDonationVisit.Show
End Sub

Private Sub Mfrmblood_issue_Click()
frmblood_issue.Show
End Sub

Private Sub MfrmConsent_Click()
frmConsent.Show
End Sub

Private Sub MfrmRefBloodRange_Click()
frmRefBloodRange.Show
End Sub

Private Sub MISSUE_Click()
Blood_issue_report.Show
End Sub

Private Sub MLOGOUT_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub mnuChatbot_Click()
    frmChatbot.Show
End Sub

Private Sub mnuDashboard_Click()
frmDashboard.Show
End Sub

Private Sub mnuLogout_Click()
    If MsgBox("Are you sure you want to Log Out?", vbQuestion + vbYesNo, "Confirm Logout") = vbYes Then
        ' Sare child forms close karo
        Dim frm As Form
        For Each frm In Forms
            If frm.Name <> "frmLogin" And frm.Name <> Me.Name Then
                Unload frm
            End If
        Next frm
        
        ' MDI band karke wapas Login screen kholo
        frmLogin.Show
        Unload Me
    End If
End Sub

Private Sub mnuExit_Click()
    If MsgBox("Do you really want to Exit the Software?", vbQuestion + vbYesNo, "Exit System") = vbYes Then
        End ' Pura project band karega
    End If
End Sub

' =========================================================
' SAFETY LOGIC (Top Right Close Button)
' =========================================================
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("Do you want to exit the Blood Bank System?", vbQuestion + vbYesNo, "Exit") = vbNo Then
            Cancel = True
        End If
    End If
End Sub



Private Sub mnuSysUser_Click()
frmRegister.Show
End Sub

Private Sub mPATIENTREPORT_Click()
Patient1_Report.Show
End Sub

Private Sub MPROCESS_Click()
COMPONENT_PROCESSING_REPORT.Show
End Sub

Private Sub MREGISTRATION_Click()
RegistrationReport.Show
End Sub

Private Sub MREQUEST_Click()
patient_request_report.Show
End Sub

Private Sub MSAMPLE_Click()
Sample_Report_form.Show
End Sub

Private Sub MSTOCK_Click()
Blood_Stock_Report.Show
End Sub

Private Sub MTESTING_Click()
Lab_Testing_Report.Show
End Sub

Private Sub MVisit_Click()
Donation_Visit_Report.Show
End Sub

Private Sub PATNT_Click()
frmPatient.Show

End Sub

Private Sub REQUEST_Click()
frmpatient_request.Show
End Sub

Private Sub SAMP_Click()
Sample.Show

End Sub

Private Sub TEST_Click()
frmLabTesting.Show
End Sub
