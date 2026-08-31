VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register Form"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox showpass 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3360
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   6440
      Width           =   180
   End
   Begin VB.ComboBox cmb_quest 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5100
      Width           =   3855
   End
   Begin VB.TextBox nm 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Text            =   "NAME"
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox sec_ans 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   4
      Text            =   "Security Answer"
      Top             =   5880
      Width           =   3615
   End
   Begin VB.TextBox new_pass 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   3
      Text            =   "PASSWORD"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox new_usernm 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Text            =   "USERNAME"
      Top             =   2480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Image register 
      Height          =   855
      Left            =   960
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Security Question"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label login 
      BackStyle       =   0  'Transparent
      Caption         =   "Already a member? Back to Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   7560
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   9255
      Left            =   0
      Picture         =   "frmRegister.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' NEW REGISTER FORM - FINAL CONSOLIDATED CODE
' =========================================================

' 1. Form Load - Setfocus hatane ke liye focus kisi neutral object (label/image) par dein
Private Sub Form_Load()
Text1.TabIndex = 0
    On Error Resume Next
    Call CONN
    ' Connection open karna
    ' Default Placeholders
    'new_usernm.Text = "USERNAME"
    'new_pass.Text = "PASSWORD"
    'nm.Text = "NAME"
    'sec_ans.Text = "Security Answer"
    
    ' ComboBox Questions
    cmb_quest.Clear
    cmb_quest.AddItem "What is your pet name?"
    cmb_quest.AddItem "What is your birth city?"
    cmb_quest.AddItem "What is your favorite color?"
    
    ' Password hide by default
    new_pass.PasswordChar = "*"
    
    ' Auto-focus hatane ke liye focus kisi neutral button ya image par shift karein
    'register.SetFocus
    showpass.Value = 1
End Sub



' 2. Show/Hide Password Logic
Private Sub showpass_Click()
    ' Agar checkbox 'showpass' checked hai toh password dikhao
    If showpass.Value = 1 Then
        new_pass.PasswordChar = ""
    Else
        ' Agar unchecked hai toh '*' dikhao
        new_pass.PasswordChar = "*"
    End If
End Sub

' 3. Placeholder Clear Logic (GotFocus Events)
Private Sub new_usernm_GotFocus()
    If new_usernm.Text = "USERNAME" Then new_usernm.Text = ""
End Sub

Private Sub new_pass_GotFocus()
    If new_pass.Text = "PASSWORD" Then new_pass.Text = ""
End Sub

Private Sub nm_GotFocus()
    If nm.Text = "NAME" Then nm.Text = ""
End Sub

Private Sub sec_ans_GotFocus()
    If sec_ans.Text = "Security Answer" Then sec_ans.Text = ""
End Sub

' 4. Enter Key Navigation Logic
Private Sub new_usernm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    KeyAscii = 0
    new_pass.SetFocus
    showpass.Value = 0
    End If
End Sub

Private Sub new_pass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: nm.SetFocus
End Sub

Private Sub nm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmb_quest.SetFocus
End Sub

' 5. Register Button Logic (Database Save)
Private Sub Register_Click()
    On Error GoTo RegErr
    
    ' Basic checks
    If new_usernm.Text = "" Or new_usernm.Text = "USERNAME" Or nm.Text = "NAME" Then
        MsgBox "Please fill in all the required details!", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Call CONN
    
    ' Query
    sql = "INSERT INTO USER_DETAILS (usernm, pass, fullname, sec_ques, sec_ans) VALUES (" & _
          "'" & new_usernm.Text & "', " & _
          "'" & new_pass.Text & "', " & _
          "'" & nm.Text & "', " & _
          "'" & cmb_quest.Text & "', " & _
          "'" & sec_ans.Text & "')"

    C.Execute sql
    MsgBox "Account Created Successfully! ?", vbInformation
    
    Unload Me
    frmLogin.Show ' Wapas Login par jane ke liye
    Exit Sub

RegErr:
    ' Detail error msg taaki pata chale kaun sa object missing hai
    MsgBox "Registration Error: " & Err.Description & vbCrLf & _
           "Check if all textbox names (new_usernm, new_pass, nm, sec_ans) are correct!", vbCritical
End Sub

' 6. Back to Login Label
Private Sub login_Click()
MsgBox "Redirecting to Login Form....", vbInformation, "Please Wait"
    Unload Me
    frmLogin.Show
End Sub

