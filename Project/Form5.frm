VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Blood_issue_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blood Issue Report"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10395
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
      Left            =   5640
      TabIndex        =   7
      Top             =   2040
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
         Format          =   174456833
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
         Format          =   174456833
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
         Top             =   1150
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
         Caption         =   "Date"
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
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   2040
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
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2520
         TabIndex        =   1
         Top             =   2040
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
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Blood_issue_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Blood Issue Report (Filter & Date Range Form)
' DESCRIPTION : Generates DataReports for dispatched blood units. Implements
'               complex 3-table SQL JOINs to fetch Patient details alongside
'               issuance records. Features dynamic search and Date Range filtering.
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes dropdowns and UI default states.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    ' 1. Populate primary search criteria (Left Section)
    Combo1.Clear
    Combo1.AddItem "Issue ID"
    Combo1.AddItem "Request ID"
    Combo1.AddItem "Stock ID"
    Combo1.AddItem "Patient Name"
    'Combo1.AddItem "Blood Group"
    
    ' Default selections
    Combo1.ListIndex = 0
    Combo2.Text = ""
    Combo2.Enabled = False
    
    ' 2. Populate Date Filter dropdown (Right Section)
    Combo3.Clear
    Combo3.AddItem "Single Date"
    Combo3.AddItem "Date Range"
    
    ' Set default dates to current system date
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    
    ' Hide 'To Date' picker by default
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
End Sub

' =========================================================================
' LEFT SECTION: FILTER BY SPECIFIC VALUE
' =========================================================================

' -------------------------------------------------------------------------
' Primary Dropdown Selection: Fetches dynamic data into Combo2 via Connection.
' -------------------------------------------------------------------------
Private Sub Combo1_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    Combo2.Clear
    Combo2.Text = ""
    Combo2.Enabled = True
    
    ' Construct SQL query based on selected criteria
    Select Case Combo1.Text
        Case "Issue ID"
            sql = "SELECT DISTINCT issue_id FROM BLOOD_ISSUE WHERE status='Active'"
        Case "Request ID"
            sql = "SELECT DISTINCT req_id FROM BLOOD_ISSUE WHERE status='Active'"
        Case "Stock ID"
            sql = "SELECT DISTINCT st_id FROM BLOOD_ISSUE WHERE status='Active'"
        Case "Patient Name"
            sql = "SELECT DISTINCT P.p_nm FROM PATIENT P, PATIENT_REQUEST PR, BLOOD_ISSUE BI WHERE P.pa_id = PR.pa_id AND PR.req_id = BI.req_id AND BI.status='Active'"
        Case "Blood Group"
            ' Static addition for Blood Groups
            Combo2.AddItem "A+"
            Combo2.AddItem "B+"
            Combo2.AddItem "AB+"
            Combo2.AddItem "O+"
            Combo2.AddItem "A-"
            Combo2.AddItem "B-"
            Combo2.AddItem "AB-"
            Combo2.AddItem "O-"
            If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
            Exit Sub
    End Select
    
    On Error GoTo ErrorHandler
    
    ' Ensure DataEnvironment connection is open
    If DataEnvironment1.Connection1.State = 0 Then
        DataEnvironment1.Connection1.Open
    End If
    
    Set rs = DataEnvironment1.Connection1.Execute(sql)
    
    ' Populate Combo2
    Do While Not rs.EOF
        If Not IsNull(rs(0)) Then
            Combo2.AddItem rs(0)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading values: " & Err.Description, vbCritical, "Database Error"
End Sub

' -------------------------------------------------------------------------
' Button Click Event (Command1): Executes and displays a Selective DataReport.
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim searchField As String
    Dim searchValue As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria and enter a search value.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    searchValue = Combo2.Text
    
    ' Map dropdown text to actual Database Columns
    Select Case Combo1.Text
        Case "Issue ID"
            searchField = "BI.issue_id"
        Case "Request ID"
            searchField = "BI.req_id"
        Case "Stock ID"
            searchField = "BI.st_id"
        Case "Patient Name"
            searchField = "P.p_nm"
        Case "Blood Group"
            searchField = "P.p_grp"
    End Select
    
    ' Master Query: 3-Table JOIN to retrieve associated Patient Data
    sql = "SELECT BI.*, P.p_nm, P.p_grp FROM BLOOD_ISSUE BI, PATIENT_REQUEST PR, PATIENT P " & _
          "WHERE BI.req_id = PR.req_id AND PR.pa_id = P.pa_id AND BI.status='Active'"
    
    ' Append exact match filter
    sql = sql & " AND " & searchField & " = '" & searchValue & "'"
    
    ' ---> FOOLPROOF LOGIC: Prevent empty reports & Error 3704 <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No records found for the selected criteria.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsBlood_Issue.State = adStateOpen Then
        DataEnvironment1.rsBlood_Issue.Close
    End If
    On Error GoTo 0
    
    ' Execute Query and Show Report
    DataEnvironment1.Commands("Blood_Issue").CommandText = sql
    DataEnvironment1.Commands("Blood_Issue").Execute
    Blood_Issue_Report1.Show
End Sub

' -------------------------------------------------------------------------
' Button Click Event (Command2): Executes and displays Comprehensive Report.
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' Master Query without specific filters
    sql = "SELECT BI.*, P.p_nm, P.p_grp FROM BLOOD_ISSUE BI, PATIENT_REQUEST PR, PATIENT P " & _
          "WHERE BI.req_id = PR.req_id AND PR.pa_id = P.pa_id AND BI.status='Active' ORDER BY BI.issue_id ASC"
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsBlood_Issue.State = adStateOpen Then
        DataEnvironment1.rsBlood_Issue.Close
    End If
    On Error GoTo 0
    
    DataEnvironment1.Commands("Blood_Issue").CommandText = sql
    DataEnvironment1.Commands("Blood_Issue").Execute
    Blood_issue_Report2.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY DATE RANGE (USING ISSUE_DT)
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
' Button Click Event (Command3): Executes DataReport based on Date filtering.
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo3.Text = "" Then
        MsgBox "Please select a date filter type (Single Date / Date Range).", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' Base Query with 3-Table JOIN (Required for Report Fields)
    Dim baseSQL As String
    baseSQL = "SELECT BI.*, P.p_nm, P.p_grp FROM BLOOD_ISSUE BI, PATIENT_REQUEST PR, PATIENT P " & _
              "WHERE BI.req_id = PR.req_id AND PR.pa_id = P.pa_id AND BI.status='Active'"

    ' Formulate SQL Query based on Date logic matching 'issue_dt'
    If Combo3.Text = "Single Date" Then
        sql = baseSQL & " AND BI.issue_dt = TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
        
    ElseIf Combo3.Text = "Date Range" Then
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "The 'From' date cannot be later than the 'To' date.", vbCritical, "Invalid Date Range"
            Exit Sub
        End If
        
        sql = baseSQL & " AND BI.issue_dt >= TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') " & _
                        " AND BI.issue_dt <= TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
    End If

    ' ---> FOOLPROOF LOGIC: Prevent empty reports & Error 3704 <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No records found for the selected date(s).", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsBlood_Issue.State = adStateOpen Then
        DataEnvironment1.rsBlood_Issue.Close
    End If
    On Error GoTo 0

    ' Execute Query and Show Report
    DataEnvironment1.Commands("Blood_Issue").CommandText = sql
    DataEnvironment1.Commands("Blood_Issue").Execute
    Blood_issue_Report2.Show
End Sub

