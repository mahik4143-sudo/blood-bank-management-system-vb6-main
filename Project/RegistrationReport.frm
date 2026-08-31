VERSION 5.00
Begin VB.Form RegistrationReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Report"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   5280
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   480
         Left            =   1920
         TabIndex        =   9
         Text            =   "Combo3"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Blood Group"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Patient_Report 
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label sea_by 
         Caption         =   "Search by:-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label sea_val 
         Caption         =   "Search Value:-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "RegistrationReport.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "RegistrationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Donor Registration Report (Filter Form)
' DESCRIPTION : Provides an administrative interface to generate DataReports
'               for ACTIVE registered donors only. (CASE SENSITIVITY FIXED)
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event
' -------------------------------------------------------------------------
Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Donor ID"
    Combo1.AddItem "Donor Name"
    Combo1.AddItem "Mobile No."
    
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    Combo3.Clear
    Combo3.AddItem "A+"
    Combo3.AddItem "A-"
    Combo3.AddItem "B+"
    Combo3.AddItem "B-"
    Combo3.AddItem "AB+"
    Combo3.AddItem "AB-"
    Combo3.AddItem "O+"
    Combo3.AddItem "O-"
End Sub

' =========================================================================
' LEFT SECTION: FILTER BY SPECIFIC VALUE
' =========================================================================

' -------------------------------------------------------------------------
' Primary Dropdown Selection (BUG FIXED: UPPERCASE STATUS CHECK)
' -------------------------------------------------------------------------
Private Sub Combo1_Click()
    Combo2.Clear
    Combo2.Enabled = True
    Dim rsS As New ADODB.Recordset
    Dim sql As String
    
    If C.State = 0 Then Call CONN
    
    ' ---> FIX: Use UPPER(status) to prevent Oracle case-sensitive mismatch <---
    If Combo1.Text = "Donor ID" Then
        sql = "SELECT d_id FROM DONAR_DETAIL WHERE UPPER(status)='ACTIVE' ORDER BY d_id ASC"
    ElseIf Combo1.Text = "Donor Name" Then
        sql = "SELECT DISTINCT nm FROM DONAR_DETAIL WHERE UPPER(status)='ACTIVE' ORDER BY nm ASC"
    ElseIf Combo1.Text = "Mobile No." Then
        sql = "SELECT DISTINCT phn FROM DONAR_DETAIL WHERE UPPER(status)='ACTIVE' AND phn IS NOT NULL ORDER BY phn ASC"
    End If
    
    rsS.CursorLocation = adUseClient
    rsS.Open sql, C, adOpenStatic, adLockReadOnly
    
    If rsS.RecordCount > 0 Then
        rsS.MoveFirst
        Do While Not rsS.EOF
            If Not IsNull(rsS.Fields(0).Value) Then
                Combo2.AddItem Trim(rsS.Fields(0).Value & "")
            End If
            rsS.MoveNext
        Loop
    Else
        MsgBox "No active records found for " & Combo1.Text, vbInformation, "Check Database"
    End If
    
    rsS.Close
    Set rsS = Nothing
End Sub

' -------------------------------------------------------------------------
' Command1_Click: Print Selective Report
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria and value!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' ---> FIX: UPPER(status) <---
    If Combo1.Text = "Donor ID" Then
        sql = "SELECT * FROM DONAR_DETAIL WHERE d_id='" & Trim(Combo2.Text) & "' AND UPPER(status)='ACTIVE'"
    ElseIf Combo1.Text = "Donor Name" Then
        sql = "SELECT * FROM DONAR_DETAIL WHERE nm='" & Trim(Combo2.Text) & "' AND UPPER(status)='ACTIVE'"
    ElseIf Combo1.Text = "Mobile No." Then
        sql = "SELECT * FROM DONAR_DETAIL WHERE phn=" & Trim(Combo2.Text) & " AND UPPER(status)='ACTIVE'"
    End If

    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected criteria.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    On Error Resume Next
    If DataEnvironment1.rsCommand1.State = 1 Then
        DataEnvironment1.rsCommand1.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Command1").CommandText = sql
    DataEnvironment1.Commands("Command1").Execute
    donar_registration_DataReport1.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print Comprehensive Report
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' ---> FIX: UPPER(status) <---
    sql = "SELECT * FROM DONAR_DETAIL WHERE UPPER(status)='ACTIVE' ORDER BY d_id ASC"
    
    On Error Resume Next
    If DataEnvironment1.rsCommand1.State = 1 Then
        DataEnvironment1.rsCommand1.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Command1").CommandText = sql
    DataEnvironment1.Commands("Command1").Execute
    Donar_Registration_Report1.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY BLOOD GROUP
' =========================================================================

' -------------------------------------------------------------------------
' Command3_Click: Print by Blood Group
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    If Combo3.Text = "" Then
        MsgBox "Please select a Blood Group to print!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' ---> FIX: UPPER(status) <---
    sql = "SELECT * FROM DONAR_DETAIL WHERE bl_grp='" & Trim(Combo3.Text) & "' AND UPPER(status)='ACTIVE'"

    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active donors found with the selected Blood Group.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    On Error Resume Next
    If DataEnvironment1.rsCommand1.State = 1 Then
        DataEnvironment1.rsCommand1.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Command1").CommandText = sql
    DataEnvironment1.Commands("Command1").Execute
    Donar_Registration_Report1.Show
End Sub
