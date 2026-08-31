VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Form"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox showpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
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
      Height          =   315
      Left            =   720
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox pass 
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
      Left            =   1680
      TabIndex        =   1
      Text            =   "PASSWORD"
      Top             =   4800
      Width           =   4335
   End
   Begin VB.TextBox usernm 
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
      Left            =   1680
      TabIndex        =   0
      Text            =   "USERNAME"
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label change_pass 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password?"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   3980
      X2              =   5700
      Y1              =   7390
      Y2              =   7390
   End
   Begin VB.Label register 
      BackStyle       =   0  'Transparent
      Caption         =   "New to Blood Bank?   Register Now"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Label forget_pass 
      BackStyle       =   0  'Transparent
      Caption         =   "Forget Password?"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   960
      TabIndex        =   3
      Top             =   5530
      Width           =   1935
   End
   Begin VB.Image login 
      Height          =   855
      Left            =   720
      Top             =   6120
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   9075
      Left            =   0
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sql As String

'=========================================================
' FORM LOAD
'=========================================================
Private Sub Form_Load()
    On Error GoTo ErrHandler

    pass.PasswordChar = "*"
    showpass.Value = 0

    usernm.Text = ""
    pass.Text = ""

    usernm.SetFocus
    Exit Sub

ErrHandler:
'    MsgBox "Error while loading login form : " & Err.Description, vbCritical, "Login Error"
End Sub

'=========================================================
' COMMON FUNCTION
'=========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

'=========================================================
' SHOW / HIDE PASSWORD
'=========================================================
Private Sub showpass_Click()
    On Error GoTo ErrHandler

    If showpass.Value = 1 Then
        pass.PasswordChar = ""
    Else
        pass.PasswordChar = "*"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error while changing password visibility : " & Err.Description, vbCritical, "Display Error"
End Sub

'=========================================================
' LOGIN BUTTON
'=========================================================
Private Sub login_Click()
    On Error GoTo LoginErr

    Dim rsUser As ADODB.Recordset
    Dim rsPass As ADODB.Recordset
    Dim FullName As String

    If Trim(usernm.Text) = "" Then
        MsgBox "Please enter User ID.", vbExclamation, "Login Validation"
        usernm.SetFocus
        Exit Sub
    End If

    If Trim(pass.Text) = "" Then
        MsgBox "Please enter Password.", vbExclamation, "Login Validation"
        pass.SetFocus
        Exit Sub
    End If

    Call CONN

    ' Step 1: Check whether username exists
    Set rsUser = New ADODB.Recordset
    sql = "SELECT * FROM USER_DETAILS WHERE USERNM='" & Esc(usernm.Text) & "'"
    rsUser.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If rsUser.EOF Then
        rsUser.Close
        Set rsUser = Nothing
        MsgBox "Wrong User ID or Password.", vbCritical, "Login Failed"
        usernm.SetFocus
        Exit Sub
    End If

    rsUser.Close
    Set rsUser = Nothing

    ' Step 2: Username exists, now check password
    Set rsPass = New ADODB.Recordset
    sql = "SELECT FULLNAME FROM USER_DETAILS WHERE USERNM='" & Esc(usernm.Text) & "' AND PASS='" & Esc(pass.Text) & "'"
    rsPass.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If rsPass.EOF Then
        rsPass.Close
        Set rsPass = Nothing
        MsgBox "Wrong Password.", vbCritical, "Login Failed"
        pass.Text = ""
        pass.SetFocus
        Exit Sub
    End If

    FullName = ""
    If Not IsNull(rsPass!FullName) Then
        FullName = Trim(CStr(rsPass!FullName))
    End If

    rsPass.Close
    Set rsPass = Nothing

    If FullName <> "" Then
        MsgBox "Welcome, " & FullName & "! Login successful.", vbInformation, "Login Success"
    Else
        MsgBox "Login successful.", vbInformation, "Login Success"
    End If

    Unload Me
    MDIForm1.Show
    Exit Sub

LoginErr:
    MsgBox "Error while logging in : " & Err.Description, vbCritical, "Login Error"
End Sub

'=========================================================
' REGISTER NOW
'=========================================================
Private Sub Register_Click()
    On Error GoTo ErrHandler

    MsgBox "Opening registration form...", vbInformation, "Please Wait"
    frmRegister.Show
    Unload Me
    Exit Sub

ErrHandler:
    MsgBox "Error while opening registration form : " & Err.Description, vbCritical, "Open Error"
End Sub

'=========================================================
' FORGOT PASSWORD
'=========================================================
Private Sub forget_pass_Click()
    On Error GoTo ErrHandler

    Dim rsRec As ADODB.Recordset
    Dim UserAns As String

    If Trim(usernm.Text) = "" Then
        MsgBox "Please enter your User ID first.", vbExclamation, "Forgot Password"
        usernm.SetFocus
        Exit Sub
    End If

    Call CONN

    Set rsRec = New ADODB.Recordset
    sql = "SELECT SEC_QUES, SEC_ANS, PASS FROM USER_DETAILS WHERE USERNM='" & Esc(usernm.Text) & "'"
    rsRec.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If rsRec.EOF Then
        rsRec.Close
        Set rsRec = Nothing
        MsgBox "Wrong User ID or Password.", vbCritical, "User Not Found"
        usernm.SetFocus
        Exit Sub
    End If

    UserAns = InputBox("Security Question : " & rsRec!SEC_QUES, "Verify Identity")

    If Trim(UserAns) = "" Then
        MsgBox "Answer cannot be blank.", vbExclamation, "Verification Failed"
        rsRec.Close
        Set rsRec = Nothing
        Exit Sub
    End If

    If UCase(Trim(UserAns)) = UCase(Trim(CStr(rsRec!sec_ans))) Then
        MsgBox "Verification successful." & vbCrLf & "Your Password is : " & rsRec!pass, vbInformation, "Password Recovery"
    Else
        MsgBox "Incorrect answer. Verification failed.", vbCritical, "Access Denied"
    End If

    rsRec.Close
    Set rsRec = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error while recovering password : " & Err.Description, vbCritical, "Recovery Error"
End Sub

'=========================================================
' CHANGE PASSWORD
'=========================================================
Private Sub change_pass_Click()
    On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset
    Dim NewPass As String
    Dim ConfirmPass As String

    If Trim(usernm.Text) = "" Then
        MsgBox "Please enter User ID first.", vbExclamation, "Change Password"
        usernm.SetFocus
        Exit Sub
    End If

    If Trim(pass.Text) = "" Then
        MsgBox "Please enter Current Password first.", vbExclamation, "Change Password"
        pass.SetFocus
        Exit Sub
    End If

    Call CONN

    Set rs = New ADODB.Recordset
    sql = "SELECT * FROM USER_DETAILS WHERE USERNM='" & Esc(usernm.Text) & "' AND PASS='" & Esc(pass.Text) & "'"
    rs.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "Wrong User ID or Password.", vbCritical, "Access Denied"
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    NewPass = InputBox("Enter New Password :", "Change Password")
    If StrPtr(NewPass) = 0 Then Exit Sub
    If Trim(NewPass) = "" Then
        MsgBox "New Password cannot be blank.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ConfirmPass = InputBox("Re-enter New Password :", "Confirm Password")
    If StrPtr(ConfirmPass) = 0 Then Exit Sub
    If Trim(ConfirmPass) = "" Then
        MsgBox "Confirm Password cannot be blank.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If NewPass <> ConfirmPass Then
        MsgBox "Passwords do not match.", vbCritical, "Mismatch"
        Exit Sub
    End If

    If Trim(NewPass) = Trim(pass.Text) Then
        MsgBox "New Password cannot be the same as Current Password.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    sql = "UPDATE USER_DETAILS SET PASS='" & Esc(NewPass) & "' WHERE USERNM='" & Esc(usernm.Text) & "'"
    C.Execute sql

    MsgBox "Password changed successfully. Please login with your new password.", vbInformation, "Success"

    pass.Text = ""
    showpass.Value = 0
    pass.PasswordChar = "*"
    pass.SetFocus
    Exit Sub

ErrHandler:
    'MsgBox "Error while changing password : " & Err.Description, vbCritical, "Change Password Error"
End Sub

'=========================================================
' ENTER KEY NAVIGATION
'=========================================================
Private Sub usernm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        pass.SetFocus
    End If
End Sub

Private Sub pass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        login_Click
    End If
End Sub
