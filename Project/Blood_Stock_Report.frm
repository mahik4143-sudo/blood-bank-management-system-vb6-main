VERSION 5.00
Begin VB.Form Blood_Stock_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blood Stock Report"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   5040
      TabIndex        =   7
      Top             =   1920
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
         Left            =   840
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Availablity"
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
      Height          =   2775
      Left            =   120
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
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
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
      Picture         =   "Blood_Stock_Report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Blood_Stock_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Blood Stock Report (Filter Form)
' DESCRIPTION : Generates DataReports for ACTIVE Blood Inventory.
'               Features specific Search (Stock ID, Process ID) on the left,
'               and a dedicated Availability filter on the right.
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes dropdowns and sets default UI state.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    On Error Resume Next
    Dim rsA As New ADODB.Recordset
    
    ' Ensure database connection is active
    If C.State = 0 Then Call CONN

    ' 1. Populate primary search criteria dropdown (Left Section)
    Combo1.Clear
    Combo1.AddItem "Stock ID"
    Combo1.AddItem "Process ID"
    
    ' Reset and disable the secondary search value dropdown by default
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' 2. Dynamically Populate Availability dropdown (Right Section)
    Combo3.Clear
    ' Fetch unique availability statuses (AVAILABLE, ISSUED, EXPIRED) from database
    Dim sqlAvail As String
    sqlAvail = "SELECT DISTINCT avail FROM BLOOD_STOCK WHERE status='Active' AND avail IS NOT NULL ORDER BY avail ASC"
    
    rsA.Open sqlAvail, C
    Do While Not rsA.EOF
        Combo3.AddItem rsA.Fields(0).Value
        rsA.MoveNext
    Loop
    rsA.Close
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
    If Combo1.Text = "Stock ID" Then
        sql = "SELECT st_id FROM BLOOD_STOCK WHERE status='Active' ORDER BY st_id ASC"
    ElseIf Combo1.Text = "Process ID" Then
        sql = "SELECT DISTINCT p_id FROM BLOOD_STOCK WHERE status='Active' ORDER BY p_id ASC"
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
    If Combo1.Text = "Stock ID" Then
        sql = "SELECT * FROM BLOOD_STOCK WHERE st_id='" & Trim(Combo2.Text) & "' AND status='Active'"
    ElseIf Combo1.Text = "Process ID" Then
        sql = "SELECT * FROM BLOOD_STOCK WHERE p_id='" & Trim(Combo2.Text) & "' AND status='Active'"
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
    If DataEnvironment1.rsBlood_Stock.State = 1 Then
        DataEnvironment1.rsBlood_Stock.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Blood_Stock").CommandText = sql
    DataEnvironment1.Commands("Blood_Stock").Execute
    STOCK_REPOTY.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print All (STRICT ACTIVE CHECK)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' --> RULE: Only fetch Active records <--
    sql = "SELECT * FROM BLOOD_STOCK WHERE status='Active' ORDER BY st_id ASC"
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsBlood_Stock.State = 1 Then
        DataEnvironment1.rsBlood_Stock.Close
    End If
    On Error GoTo 0

    ' Execute query and display complete report
    DataEnvironment1.Commands("Blood_Stock").CommandText = sql
    DataEnvironment1.Commands("Blood_Stock").Execute
    Blood_Stock_Report1.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY AVAILABILITY (AVAIL)
' =========================================================================

' -------------------------------------------------------------------------
' Command3_Click: Print by Availability (STRICT ACTIVE CHECK)
' Note: Assign this code to the PRINT button under Availability.
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo3.Text = "" Then
        MsgBox "Please select Availability status to print!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Formulate the SQL query based on Availability
    sql = "SELECT * FROM BLOOD_STOCK WHERE avail='" & Trim(Combo3.Text) & "' AND status='Active'"

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected Availability status.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsBlood_Stock.State = 1 Then
        DataEnvironment1.rsBlood_Stock.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Blood_Stock").CommandText = sql
    DataEnvironment1.Commands("Blood_Stock").Execute
    Blood_Stock_Report1.Show
End Sub
