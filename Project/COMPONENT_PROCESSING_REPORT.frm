VERSION 5.00
Begin VB.Form COMPONENT_PROCESSING_REPORT 
   Caption         =   "COMPONENT PROCESSING REPORT"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   5160
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
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
         Left            =   1920
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
         Alignment       =   2  'Center
         Caption         =   "Component  Name"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   1800
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
      Picture         =   "COMPONENT_PROCESSING_REPORT.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "COMPONENT_PROCESSING_REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Component Processing Report (Filter Form)
' DESCRIPTION : Generates DataReports for ACTIVE Processed Blood Components.
'               Features specific Search (Processing ID, Test ID) on the left,
'               and a dedicated Component Name filter on the right.
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes dropdowns and sets default UI state.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    On Error Resume Next
    Dim rsC As New ADODB.Recordset
    
    ' Ensure database connection is active for dynamic loading
    If C.State = 0 Then Call CONN

    ' 1. Populate primary search criteria dropdown (Left Section)
    Combo1.Clear
    Combo1.AddItem "Processing ID"
    Combo1.AddItem "Test ID"
    
    ' Reset and disable the secondary search value dropdown by default
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' 2. Dynamically Populate Component Name dropdown (Right Section)
    Combo3.Clear
    ' Fetch only unique, active component names directly from the database
    Dim sqlComp As String
    sqlComp = "SELECT DISTINCT com_nm FROM COMPONENT_PROCESSING WHERE status='Active' AND com_nm IS NOT NULL ORDER BY com_nm ASC"
    
    rsC.Open sqlComp, C
    Do While Not rsC.EOF
        Combo3.AddItem rsC.Fields(0).Value
        rsC.MoveNext
    Loop
    rsC.Close
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
    If Combo1.Text = "Processing ID" Then
        sql = "SELECT p_id FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY p_id ASC"
    ElseIf Combo1.Text = "Test ID" Then
        sql = "SELECT DISTINCT t_id FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY t_id ASC"
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
    If Combo1.Text = "Processing ID" Then
        sql = "SELECT * FROM COMPONENT_PROCESSING WHERE p_id='" & Trim(Combo2.Text) & "' AND status='Active'"
    ElseIf Combo1.Text = "Test ID" Then
        sql = "SELECT * FROM COMPONENT_PROCESSING WHERE t_id='" & Trim(Combo2.Text) & "' AND status='Active'"
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
    If DataEnvironment1.rsComponent_Processing.State = 1 Then
        DataEnvironment1.rsComponent_Processing.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Component_Processing").CommandText = sql
    DataEnvironment1.Commands("Component_Processing").Execute
    Component_Processing_Report2.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print All (STRICT ACTIVE CHECK)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' --> RULE: Only fetch Active records <--
    sql = "SELECT * FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY p_id ASC"
    
    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsComponent_Processing.State = 1 Then
        DataEnvironment1.rsComponent_Processing.Close
    End If
    On Error GoTo 0

    ' Execute query and display complete report
    DataEnvironment1.Commands("Component_Processing").CommandText = sql
    DataEnvironment1.Commands("Component_Processing").Execute
    Component_Processing_Report1.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY COMPONENT NAME
' =========================================================================

' -------------------------------------------------------------------------
' Command3_Click: Print by Component Name (STRICT ACTIVE CHECK)
' Note: Assign this code to the PRINT button under Component Name.
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    ' Input Validation
    If Combo3.Text = "" Then
        MsgBox "Please select a Component Name to print!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Formulate the SQL query based on Component Name
    sql = "SELECT * FROM COMPONENT_PROCESSING WHERE com_nm='" & Trim(Combo3.Text) & "' AND status='Active'"

    ' ---> FOOLPROOF LOGIC: Empty Report & Error Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected Component Name.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    ' State Management
    On Error Resume Next
    If DataEnvironment1.rsComponent_Processing.State = 1 Then
        DataEnvironment1.rsComponent_Processing.Close
    End If
    On Error GoTo 0

    ' Execute query and display selective report
    DataEnvironment1.Commands("Component_Processing").CommandText = sql
    DataEnvironment1.Commands("Component_Processing").Execute
    Component_Processing_Report1.Show
End Sub
