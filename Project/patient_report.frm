VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form patient_request_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Request Report"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "By Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "PRINT"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   174325761
         CurrentDate     =   46130
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   174325761
         CurrentDate     =   46130
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Select by"
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
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Req Date"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1680
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
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2400
         TabIndex        =   1
         Top             =   1920
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
      Picture         =   "patient_report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "patient_request_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Patient Request Report (Filter & Date Range Form)
' DESCRIPTION : Generates DataReports for ACTIVE Patient Requests.
'               Uses a JOIN query with the PATIENT table to fetch Name and
'               Blood Group. Features specific Search and Date Range filtering.
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes search criteria and sets default UI state.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    ' 1. Populate primary search criteria dropdown (Left Section)
    Combo1.Clear
    Combo1.AddItem "Request ID"
    Combo1.AddItem "Patient ID"
    
    ' Reset and disable the secondary search value dropdown by default
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' 2. Populate Date Filter dropdown (Right Section)
    Combo3.Clear
    Combo3.AddItem "Single Date"
    Combo3.AddItem "Date Range"
    
    ' Set default dates to current system date
    DTPicker1.Value = Date ' From Date
    DTPicker2.Value = Date ' To Date
    
    ' Default UI State: Hide the 'To Date' picker until 'Date Range' is selected
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
End Sub

' =========================================================================
' LEFT SECTION: FILTER BY SPECIFIC VALUE
' =========================================================================

' -------------------------------------------------------------------------
' Primary Dropdown Selection: Fetches ONLY ACTIVE data into Combo2
' -------------------------------------------------------------------------
Private Sub Combo1_Click()
    Combo2.Clear
    Combo2.Enabled = True
    Dim rsS As New ADODB.Recordset
    Dim sql As String
    
    ' Ensure database connection is active
    If C.State = 0 Then Call CONN
    
    ' --> RULE: Fetch ONLY Active records based on selection <--
    If Combo1.Text = "Request ID" Then
        sql = "SELECT req_id FROM PATIENT_REQUEST WHERE status='Active' ORDER BY req_id ASC"
    ElseIf Combo1.Text = "Patient ID" Then
        sql = "SELECT DISTINCT pa_id FROM PATIENT_REQUEST WHERE status='Active' ORDER BY pa_id ASC"
    End If
    
    ' Execute the query and populate Combo2
    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

' -------------------------------------------------------------------------
' Command1_Click: Print specific record (STRICT ACTIVE CHECK + JOIN)
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria and value!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Formulate the filtered SQL query with JOIN (PR = Patient Request, P = Patient)
    ' --> RULE: Both Request AND Patient must be Active <--
    If Combo1.Text = "Request ID" Then
        sql = "SELECT PR.*, P.p_nm, P.p_grp FROM PATIENT_REQUEST PR, PATIENT P " & _
              "WHERE PR.pa_id = P.pa_id AND PR.status='Active' AND P.status='Active' AND PR.req_id='" & Trim(Combo2.Text) & "'"
    ElseIf Combo1.Text = "Patient ID" Then
        sql = "SELECT PR.*, P.p_nm, P.p_grp FROM PATIENT_REQUEST PR, PATIENT P " & _
              "WHERE PR.pa_id = P.pa_id AND PR.status='Active' AND P.status='Active' AND PR.pa_id='" & Trim(Combo2.Text) & "'"
    End If

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active requests found for the selected criteria.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub ' Stop execution if no data exists
    End If
    rsCheck.Close

    ' State Management: Safely close previous Recordset
    On Error Resume Next
    If DataEnvironment1.rsPatient_request.State = 1 Then
        DataEnvironment1.rsPatient_request.Close
    End If
    On Error GoTo 0

    ' Execute query and display report
    DataEnvironment1.Commands("Patient_request").CommandText = sql
    DataEnvironment1.Commands("Patient_request").Execute
    Patient_Request_Report1.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print All (STRICT ACTIVE CHECK + JOIN)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' --> RULE: Only fetch Active records <--
    sql = "SELECT PR.*, P.p_nm, P.p_grp FROM PATIENT_REQUEST PR, PATIENT P " & _
          "WHERE PR.pa_id = P.pa_id AND PR.status='Active' AND P.status='Active' ORDER BY PR.req_id ASC"
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsPatient_request.State = 1 Then
        DataEnvironment1.rsPatient_request.Close
    End If
    On Error GoTo 0

    ' Execute query and display complete report
    DataEnvironment1.Commands("Patient_request").CommandText = sql
    DataEnvironment1.Commands("Patient_request").Execute
    Patient_Request_Report2.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY DATE RANGE (USING 'DT' COLUMN)
' =========================================================================

' -------------------------------------------------------------------------
' Dropdown Event (Combo3): Toggles visibility of the 'To Date' picker.
' -------------------------------------------------------------------------
Private Sub Combo3_Click()
    If Combo3.Text = "Single Date" Then
        DTPicker2.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        
    ElseIf Combo3.Text = "Date Range" Then
        DTPicker2.Visible = True
        Label1.Visible = True
        Label2.Visible = True
    End If
End Sub

' -------------------------------------------------------------------------
' Command3_Click: Print by Date Range (STRICT ACTIVE CHECK + JOIN)
' Note: Assign this code to the PRINT button on the right side.
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo3.Text = "" Then
        MsgBox "Please select a date filter type (Single Date / Date Range).", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' Base Query with JOIN
    Dim baseSQL As String
    baseSQL = "SELECT PR.*, P.p_nm, P.p_grp FROM PATIENT_REQUEST PR, PATIENT P " & _
              "WHERE PR.pa_id = P.pa_id AND PR.status='Active' AND P.status='Active'"

    ' Formulate the SQL query based on Date logic (Schema column is 'DT')
    If Combo3.Text = "Single Date" Then
        sql = baseSQL & " AND PR.dt = TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
        
    ElseIf Combo3.Text = "Date Range" Then
        ' Validate that From Date is not greater than To Date
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "The 'From' date cannot be later than the 'To' date.", vbCritical, "Invalid Date Range"
            Exit Sub
        End If
        
        sql = baseSQL & " AND PR.dt >= TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') " & _
                        " AND PR.dt <= TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
    End If

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active requests found for the selected date(s).", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsPatient_request.State = 1 Then
        DataEnvironment1.rsPatient_request.Close
    End If
    On Error GoTo 0

    ' Execute query and display report
    DataEnvironment1.Commands("Patient_request").CommandText = sql
    DataEnvironment1.Commands("Patient_request").Execute
    Patient_Request_Report2.Show
End Sub

