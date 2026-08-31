VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPatient 
   Caption         =   "PATIENT"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   17775
   Begin VB.Frame Frame4 
      Height          =   4335
      Left            =   10440
      TabIndex        =   13
      Top             =   1800
      Width           =   6735
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   14
         Text            =   "Combo3"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label sea_by 
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label sea_val 
         Caption         =   "Search Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Image CmdSearch 
         Height          =   735
         Left            =   2160
         Picture         =   "PATIENT.frx":0000
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1800
      TabIndex        =   11
      Top             =   6360
      Width           =   15375
      Begin VB.Image cmdreport 
         Height          =   735
         Left            =   13080
         Picture         =   "PATIENT.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   735
         Left            =   10440
         Picture         =   "PATIENT.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdClear 
         Height          =   735
         Left            =   7800
         Picture         =   "PATIENT.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdSave 
         Height          =   735
         Left            =   5160
         Picture         =   "PATIENT.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdUpdate 
         Height          =   735
         Left            =   2640
         Picture         =   "PATIENT.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdAddNew 
         Height          =   735
         Left            =   120
         Picture         =   "PATIENT.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   8055
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Phone No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Blood Group "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Patient ID "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PATIENT.frx":20B9B
      Height          =   2295
      Left            =   1800
      TabIndex        =   12
      Top             =   7680
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Patient"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "PA_ID"
         Caption         =   "PATIENT ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "P_NM"
         Caption         =   "NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "P_GRP"
         Caption         =   "BLOOD GROUP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ADDR"
         Caption         =   "ADDRESS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "P_PN"
         Caption         =   "PHONE NO."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2099.906
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "PATIENT.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmPatient
' FINAL CODE WITH BUTTON RESTRICTIONS + PROPER CASE
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

' =========================================================
' FORM LOAD
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' Blood Group List Setup
    Combo1.Clear
    Combo1.AddItem "A+"
    Combo1.AddItem "A-"
    Combo1.AddItem "B+"
    Combo1.AddItem "B-"
    Combo1.AddItem "AB+"
    Combo1.AddItem "AB-"
    Combo1.AddItem "O+"
    Combo1.AddItem "O-"

    ' Search Combos Setup
    Combo2.Clear
    Combo2.AddItem "Patient ID"
    Combo2.AddItem "Name"
    Combo2.AddItem "Phone No"
    Combo3.Enabled = False

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
End Sub

' =========================================================
' MODE / BUTTON / CONTROL RESTRICTIONS
' =========================================================
Private Sub SetStartupMode()
    mMode = "STARTUP"

    Call EnableEntryControls(False)
    Call EnableSearchControls(True)

    CmdAddNew.Enabled = True
    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = False
    cmdSearch.Enabled = False
    cmdreport.Enabled = True

    Combo2.SetFocus
End Sub

Private Sub SetAddMode()
    mMode = "ADD"

    Call EnableEntryControls(True)
    Call EnableSearchControls(False)

    CmdAddNew.Enabled = False
    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = True
    cmdSearch.Enabled = False
    cmdreport.Enabled = False

    Text1.SetFocus
End Sub

Private Sub SetEditMode()
    mMode = "EDIT"

    Call EnableEntryControls(True)
    Call EnableSearchControls(True)

    CmdAddNew.Enabled = True
    CmdSave.Enabled = False
    CmdUpdate.Enabled = True
    CmdDelete.Enabled = True
    CmdClear.Enabled = True
    cmdSearch.Enabled = True
    cmdreport.Enabled = True

    Text1.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    Text1.Enabled = bFlag   ' Name
    Combo1.Enabled = bFlag  ' Blood Group
    Text2.Enabled = bFlag   ' Address
    Text3.Enabled = bFlag   ' Phone
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo2.Enabled = bFlag
    Combo3.Enabled = bFlag
End Sub

' =========================================================
' HELPERS
' =========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function ProperCaseText(ByVal s As String) As String
    s = Trim(s)

    If s = "" Then
        ProperCaseText = ""
    Else
        ProperCaseText = UCase(Left$(s, 1)) & LCase(Mid$(s, 2))
    End If
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Integer
    s = Trim(s)

    If s = "" Then
        IsDigitsOnly = False
        Exit Function
    End If

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i

    IsDigitsOnly = True
End Function

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(Label6.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Action Required"
        Exit Function
    End If

    If Trim(Text1.Text) = "" Then
        MsgBox "Please enter Patient Name.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Trim(Combo1.Text) = "" Then
        MsgBox "Please select Blood Group.", vbExclamation, "Validation Error"
        Combo1.SetFocus
        Exit Function
    End If

    If Trim(Text3.Text) = "" Then
        MsgBox "Please enter Phone Number.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(Text3.Text) Or Len(Trim(Text3.Text)) <> 10 Then
        MsgBox "Please enter a valid 10-digit Phone Number.", vbExclamation, "Invalid Phone"
        Text3.SetFocus
        Exit Function
    End If

    If Len(Trim(Text1.Text)) > 50 Then
        MsgBox "Patient Name cannot exceed 50 characters.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Len(Trim(Text2.Text)) > 200 Then
        MsgBox "Address cannot exceed 200 characters.", vbExclamation, "Validation Error"
        Text2.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

' =========================================================
' AUTO ID
' =========================================================
Private Sub Auto_PatientID()
    Dim rsMax As New ADODB.Recordset
    Dim nextNum As Long

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(pa_id, 3))), 0) FROM PATIENT"
    Set rsMax = C.Execute(sql)

    nextNum = rsMax.Fields(0).Value + 1
    Label6.Caption = "PA" & Format(nextNum, "0000")

    rsMax.Close
End Sub

' =========================================================
' ADD NEW
' =========================================================
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_PatientID
    Call SetAddMode
    Text1.SetFocus
End Sub

' =========================================================
' SAVE
' =========================================================
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If C.State = 0 Then Call CONN

    sql = "INSERT INTO PATIENT (pa_id, p_nm, p_grp, addr, p_pn, status) " & _
          "VALUES (" & _
          "'" & Esc(Label6.Caption) & "', " & _
          "'" & Esc(Text1.Text) & "', " & _
          "'" & Esc(Combo1.Text) & "', " & _
          "'" & Esc(Text2.Text) & "', " & _
          "" & Val(Text3.Text) & ", " & _
          "'Active')"

    C.Execute sql
    MsgBox "Patient registered successfully.", vbInformation, "Success"

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' UPDATE
' =========================================================
Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(Label6.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If MsgBox("Are you sure you want to update patient details?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        Call CONN

        sql = "UPDATE PATIENT SET " & _
              "p_nm='" & Esc(Text1.Text) & "', " & _
              "p_grp='" & Esc(Combo1.Text) & "', " & _
              "addr='" & Esc(Text2.Text) & "', " & _
              "p_pn=" & Val(Text3.Text) & " " & _
              "WHERE pa_id='" & Esc(Label6.Caption) & "'"

        C.Execute sql
        MsgBox "Patient details updated successfully.", vbInformation, "Success"

        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' DELETE (SOFT DELETE)
' =========================================================
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(Label6.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If MsgBox("Do you want to deactivate this patient?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN

        sql = "UPDATE PATIENT SET status='Inactive' WHERE pa_id='" & Esc(Label6.Caption) & "'"
        C.Execute sql

        MsgBox "Patient deactivated successfully.", vbInformation, "Deleted"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' SEARCH LOAD
' =========================================================
Private Sub Combo2_Click()
    If Combo2.Text = "" Then Exit Sub

    Combo3.Clear
    Combo3.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo2.Text = "Patient ID" Then sql = "SELECT pa_id FROM PATIENT WHERE status='Active' ORDER BY pa_id ASC"
    If Combo2.Text = "Name" Then sql = "SELECT DISTINCT p_nm FROM PATIENT WHERE status='Active' ORDER BY p_nm ASC"
    If Combo2.Text = "Phone No" Then sql = "SELECT DISTINCT p_pn FROM PATIENT WHERE status='Active' ORDER BY p_pn ASC"

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo3.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo3_Click()
    If Trim(Combo3.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

' =========================================================
' SEARCH
' =========================================================
Private Sub cmdSearch_Click()
    If Combo2.Text = "" Or Combo3.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Error"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo2.Text = "Patient ID" Then sql = "SELECT * FROM PATIENT WHERE pa_id='" & Esc(Combo3.Text) & "' AND status='Active'"
    If Combo2.Text = "Name" Then sql = "SELECT * FROM PATIENT WHERE p_nm='" & Esc(Combo3.Text) & "' AND status='Active'"
    If Combo2.Text = "Phone No" Then sql = "SELECT * FROM PATIENT WHERE p_pn=" & Val(Combo3.Text) & " AND status='Active'"

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label6.Caption = rsSearch!pa_id
        Text1.Text = "" & rsSearch!p_nm
        Combo1.Text = "" & rsSearch!p_grp
        Text2.Text = "" & rsSearch!addr
        Text3.Text = "" & rsSearch!p_pn

        MsgBox "Record found successfully.", vbInformation, "Success"
        Call SetEditMode
    Else
        MsgBox "Record not found or inactive.", vbCritical, "Not Found"
    End If
    rsSearch.Close
End Sub

' =========================================================
' REFRESH GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsPatient.State = adStateOpen Then
        DataEnvironment1.rsPatient.Close
    End If

    DataEnvironment1.Commands("Patient").CommandText = _
        "SELECT * FROM PATIENT WHERE (status='Active' OR status IS NULL) " & _
        "ORDER BY TO_NUMBER(SUBSTR(pa_id,3)) ASC"

    DataEnvironment1.Patient
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Patient"
    DataGrid1.Refresh
End Sub

' =========================================================
' CLEAR
' =========================================================
Public Sub Clear_form()
    Label6.Caption = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo1.ListIndex = -1
    'Combo1.Text = ""
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    Combo3.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

' =========================================================
' LOST FOCUS PROPER CASE
' =========================================================
Private Sub Text1_LostFocus()
    Text1.Text = ProperCaseText(Text1.Text)
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = ProperCaseText(Text2.Text)
End Sub

' =========================================================
' KEY NAVIGATION / RESTRICTIONS
' =========================================================
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Combo1.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    If Len(Text3.Text) >= 10 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

' =========================================================
' REPORT
' =========================================================
Private Sub cmdreport_Click()
    Patient1_Report.Show
End Sub

