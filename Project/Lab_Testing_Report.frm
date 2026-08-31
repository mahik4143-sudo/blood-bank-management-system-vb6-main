VERSION 5.00
Begin VB.Form Lab_Testing_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lab Testing Report"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
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
         TabIndex        =   9
         Text            =   "Combo3"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PRINT"
         Height          =   735
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Final Status"
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
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
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
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1080
         Width           =   2055
      End
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
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   480
         Width           =   2055
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
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
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Lab_Testing_Report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Lab_Testing_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Lab Testing Report (Filter Form)
' DESCRIPTION : Generates DataReports for ACTIVE Lab Testing records.
'               Features specific Search (Test ID, Sample ID) on the left,
'               and a dedicated Final Status (fin_res) filter on the right.
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes dropdowns and sets default UI state.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    On Error Resume Next
    Dim rsF As New ADODB.Recordset
    
    ' Ensure database connection is active
    If C.State = 0 Then Call CONN

    ' 1. Populate primary search criteria dropdown (Left Section)
    Combo1.Clear
    Combo1.AddItem "Test ID"
    Combo1.AddItem "Sample ID"
    
    ' Reset and disable the secondary search value dropdown by default
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' 2. Dynamically Populate Final Status dropdown (Right Section)
    Combo3.Clear
    ' Fetch unique active final results (e.g., PASS, FAIL) from database
    Dim sqlFin As String
    sqlFin = "SELECT DISTINCT fin_res FROM LAB_TESTING WHERE status='Active' AND fin_res IS NOT NULL ORDER BY fin_res ASC"
    
    rsF.Open sqlFin, C
    Do While Not rsF.EOF
        Combo3.AddItem rsF.Fields(0).Value
        rsF.MoveNext
    Loop
    rsF.Close
End Sub

' =========================================================================
' LEFT SECTION: FILTER BY SPECIFIC ID
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
    If Combo1.Text = "Test ID" Then
        sql = "SELECT t_id FROM LAB_TESTING WHERE status='Active' ORDER BY t_id ASC"
    ElseIf Combo1.Text = "Sample ID" Then
        sql = "SELECT DISTINCT s_id FROM LAB_TESTING WHERE status='Active' ORDER BY s_id ASC"
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
' Command1_Click: Print specific record (STRICT ACTIVE CHECK)
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria and value!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Formulate the filtered SQL query
    If Combo1.Text = "Test ID" Then
        sql = "SELECT * FROM LAB_TESTING WHERE t_id='" & Trim(Combo2.Text) & "' AND status='Active'"
    ElseIf Combo1.Text = "Sample ID" Then
        sql = "SELECT * FROM LAB_TESTING WHERE s_id='" & Trim(Combo2.Text) & "' AND status='Active'"
    End If

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected criteria.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub ' Stop execution if no data exists
    End If
    rsCheck.Close

    ' State Management: Safely close previous Recordset
    On Error Resume Next
    If DataEnvironment1.rsLab_Testing.State = 1 Then
        DataEnvironment1.rsLab_Testing.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Lab_Testing").CommandText = sql
    DataEnvironment1.Commands("Lab_Testing").Execute
    Lab_Testing_Report2.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print All (STRICT ACTIVE CHECK)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' --> RULE: Only fetch Active records <--
    sql = "SELECT * FROM LAB_TESTING WHERE status='Active' ORDER BY t_id ASC"
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsLab_Testing.State = 1 Then
        DataEnvironment1.rsLab_Testing.Close
    End If
    On Error GoTo 0

    ' Execute query and display complete report
    DataEnvironment1.Commands("Lab_Testing").CommandText = sql
    DataEnvironment1.Commands("Lab_Testing").Execute
    Lab_Testing_Report1.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY FINAL STATUS (FIN_RES)
' =========================================================================

' -------------------------------------------------------------------------
' Command3_Click: Print by Final Status (STRICT ACTIVE CHECK)
' Note: Assign this code to the PRINT button under Final Status.
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo3.Text = "" Then
        MsgBox "Please select a Final Status to print!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Formulate the SQL query based on Final Status (fin_res)
    sql = "SELECT * FROM LAB_TESTING WHERE fin_res='" & Trim(Combo3.Text) & "' AND status='Active'"

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected Final Status.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsLab_Testing.State = 1 Then
        DataEnvironment1.rsLab_Testing.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Lab_Testing").CommandText = sql
    DataEnvironment1.Commands("Lab_Testing").Execute
    Lab_Testing_Report1.Show
End Sub
