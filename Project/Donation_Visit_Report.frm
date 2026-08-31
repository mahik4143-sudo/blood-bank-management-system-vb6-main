VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Donation_Visit_Report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Donation Visit Report"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
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
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PRINT"
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   2520
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   1680
         TabIndex        =   9
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
         Format          =   174784513
         CurrentDate     =   46130
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1680
         TabIndex        =   10
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
         Format          =   174784513
         CurrentDate     =   46130
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   480
         Width           =   1095
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
         TabIndex        =   12
         Top             =   1150
         Width           =   1095
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
         TabIndex        =   11
         Top             =   1750
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2400
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   2040
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
         Top             =   1200
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
         TabIndex        =   4
         Top             =   1320
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
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Donation_Visit_Report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Donation_Visit_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Donation Visit Report (Filter & Date Range Form)
' DESCRIPTION : Provides an administrative interface to generate DataReports
'               for ACTIVE donor visits only. (CASE SENSITIVITY FIXED)
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event
' -------------------------------------------------------------------------
Private Sub Form_Load()
    ' 1. Populate primary search criteria
    Combo1.Clear
    Combo1.AddItem "Visit ID"
    Combo1.AddItem "Donar ID"
    Combo1.AddItem "Patient Name"
    
    ' Reset and disable secondary dropdown
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' 2. Populate Date Filter dropdown
    Combo3.Clear
    Combo3.AddItem "Single Date"
    Combo3.AddItem "Date Range"
    
    ' Set default dates
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    
    ' Default UI State
    DTPicker2.Visible = False
    On Error Resume Next
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = True
    On Error GoTo 0
End Sub

' =========================================================================
' LEFT SECTION: FILTER BY SPECIFIC VALUE
' =========================================================================

' -------------------------------------------------------------------------
' Primary Dropdown Selection (BUG FIXED: UPPERCASE STATUS & CLIENT CURSOR)
' -------------------------------------------------------------------------
Private Sub Combo1_Click()
    Combo2.Clear
    Combo2.Enabled = True
    Dim rsS As New ADODB.Recordset
    Dim sql As String
    
    If C.State = 0 Then Call CONN
    
    ' ---> FIX: Use UPPER(status) <---
    If Combo1.Text = "Visit ID" Then
        sql = "SELECT v_id FROM DONATION_VISIT WHERE UPPER(status)='ACTIVE' ORDER BY v_id ASC"
    ElseIf Combo1.Text = "Donar ID" Then
        sql = "SELECT DISTINCT d_id FROM DONATION_VISIT WHERE UPPER(status)='ACTIVE' ORDER BY d_id ASC"
    ElseIf Combo1.Text = "Patient Name" Then
        sql = "SELECT DISTINCT ref_nm FROM DONATION_VISIT WHERE UPPER(status)='ACTIVE' AND ref_nm IS NOT NULL ORDER BY ref_nm ASC"
    End If
    
    ' ---> FIX: Client Cursor for Oracle <---
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
' Command1_Click: Print Selective Report (BUG FIXED: ADDED rsCheck & UPPER)
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset

    If Combo2.Text = "" Then
        MsgBox "Please select search criteria and value!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' ---> FIX: UPPER(status) <---
    If Combo1.Text = "Visit ID" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE v_id='" & Trim(Combo2.Text) & "' AND UPPER(status)='ACTIVE'"
    ElseIf Combo1.Text = "Donar ID" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE d_id='" & Trim(Combo2.Text) & "' AND UPPER(status)='ACTIVE'"
    ElseIf Combo1.Text = "Patient Name" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE ref_nm='" & Trim(Combo2.Text) & "' AND UPPER(status)='ACTIVE'"
    End If

    ' ---> FIX: Added Empty Report Prevention <---
    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected criteria.", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    On Error Resume Next
    If DataEnvironment1.rsDonation_Visit.State = 1 Then
        DataEnvironment1.rsDonation_Visit.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Donation_Visit").CommandText = sql
    DataEnvironment1.Commands("Donation_Visit").Execute
    Donation_Visit_Report1.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print Comprehensive Report (BUG FIXED: ADDED WHERE CLAUSE)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' ---> FIX: Added WHERE UPPER(status)='ACTIVE' <---
    sql = "SELECT * FROM DONATION_VISIT WHERE UPPER(status)='ACTIVE' ORDER BY v_id ASC"
    
    On Error Resume Next
    If DataEnvironment1.rsDonation_Visit.State = 1 Then
        DataEnvironment1.rsDonation_Visit.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Donation_Visit").CommandText = sql
    DataEnvironment1.Commands("Donation_Visit").Execute
    Donation_Visit_Report2.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY DATE RANGE (USING DON_DT)
' =========================================================================

' -------------------------------------------------------------------------
' Dropdown Event (Combo3): Toggles visibility
' -------------------------------------------------------------------------
Private Sub Combo3_Click()
    On Error Resume Next
    If Combo3.Text = "Single Date" Then
        DTPicker2.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label4.Visible = True
    ElseIf Combo3.Text = "Date Range" Then
        DTPicker2.Visible = True
        Label2.Visible = True
        Label1.Visible = True
        Label4.Visible = False
    End If
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' Command3_Click: Print by Date Range (BUG FIXED: UPPERCASE STATUS)
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    If Combo3.Text = "" Then
        MsgBox "Please select a date filter type (Single Date / Date Range).", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' ---> FIX: UPPER(status) <---
    If Combo3.Text = "Single Date" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE don_dt = TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') AND UPPER(status)='ACTIVE'"
        
    ElseIf Combo3.Text = "Date Range" Then
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "The 'From' date cannot be later than the 'To' date.", vbCritical, "Invalid Date Range"
            Exit Sub
        End If
        
        sql = "SELECT * FROM DONATION_VISIT WHERE don_dt >= TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') " & _
              "AND don_dt <= TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') AND UPPER(status)='ACTIVE'"
    End If

    If C.State = 0 Then Call CONN
    rsCheck.Open sql, C
    If rsCheck.EOF Then
        MsgBox "No active records found for the selected date(s).", vbInformation, "Red Cross Blood Bank"
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    On Error Resume Next
    If DataEnvironment1.rsDonation_Visit.State = 1 Then
        DataEnvironment1.rsDonation_Visit.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("Donation_Visit").CommandText = sql
    DataEnvironment1.Commands("Donation_Visit").Execute
    
    ' Form usually uses Report1 format for selective printing
    Donation_Visit_Report1.Show
End Sub

