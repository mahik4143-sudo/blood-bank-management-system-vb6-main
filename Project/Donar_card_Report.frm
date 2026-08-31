VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Donar_card_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donar Card Report"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9915
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
      Left            =   5400
      TabIndex        =   7
      Top             =   1440
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
         Format          =   174784513
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
         Format          =   174784513
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
         Width           =   1215
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
         Caption         =   "Issue Date"
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
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
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
         Caption         =   "Search by"
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
         Caption         =   "Search Value"
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
      Picture         =   "Donar_card_Report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Donar_card_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' MODULE      : Donor Card Report (Filter & Date Range Form)
' DESCRIPTION : Generates DataReports for ACTIVE Donor Identity Cards only.
'               (CASE SENSITIVITY FIXED FOR JOIN QUERIES)
' =========================================================================

Option Explicit

' -------------------------------------------------------------------------
' Form Load Event: Initializes search criteria and sets default UI state.
' -------------------------------------------------------------------------
Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Card Number"
    Combo1.AddItem "Donor ID"
    
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    Combo3.Clear
    Combo3.AddItem "Single Date"
    Combo3.AddItem "Date Range"
    
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker2.Visible = False
    Label1.Visible = False
    Label2.Visible = False
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
    If Combo1.Text = "Card Number" Then
        sql = "SELECT c_no FROM DONAR_CARD WHERE UPPER(status)='ACTIVE' ORDER BY c_no ASC"
    ElseIf Combo1.Text = "Donor ID" Then
        sql = "SELECT DISTINCT d_id FROM DONAR_CARD WHERE UPPER(status)='ACTIVE' ORDER BY d_id ASC"
    End If
    
    ' ---> FIX: Client Cursor added <---
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
    End If
    
    rsS.Close
    Set rsS = Nothing
End Sub

' -------------------------------------------------------------------------
' Command1_Click: Print specific record (STRICT ACTIVE CHECK WITH UPPER)
' -------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria and value!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' ---> FIX: Apply UPPER() on BOTH tables C and D <---
    If Combo1.Text = "Card Number" Then
        sql = "SELECT C.*, D.nm, D.bl_grp FROM DONAR_CARD C, DONAR_DETAIL D " & _
              "WHERE C.d_id = D.d_id AND UPPER(C.status)='ACTIVE' AND UPPER(D.status)='ACTIVE' AND C.c_no=" & Trim(Combo2.Text)
    ElseIf Combo1.Text = "Donor ID" Then
        sql = "SELECT C.*, D.nm, D.bl_grp FROM DONAR_CARD C, DONAR_DETAIL D " & _
              "WHERE C.d_id = D.d_id AND UPPER(C.status)='ACTIVE' AND UPPER(D.status)='ACTIVE' AND C.d_id='" & Trim(Combo2.Text) & "'"
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
    If DataEnvironment1.rsDonarCard_Report.State = 1 Then
        DataEnvironment1.rsDonarCard_Report.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("DonarCard_Report").CommandText = sql
    DataEnvironment1.Commands("DonarCard_Report").Execute
    Donar_Card_Report2.Show
End Sub

' -------------------------------------------------------------------------
' Command2_Click: Print All (STRICT ACTIVE CHECK WITH UPPER)
' -------------------------------------------------------------------------
Private Sub Command2_Click()
    Dim sql As String
    
    ' ---> FIX: Apply UPPER() on BOTH tables C and D <---
    sql = "SELECT C.*, D.nm, D.bl_grp FROM DONAR_CARD C, DONAR_DETAIL D " & _
          "WHERE C.d_id = D.d_id AND UPPER(C.status)='ACTIVE' AND UPPER(D.status)='ACTIVE' ORDER BY C.c_no ASC"
    
    On Error Resume Next
    If DataEnvironment1.rsDonarCard_Report.State = 1 Then
        DataEnvironment1.rsDonarCard_Report.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("DonarCard_Report").CommandText = sql
    DataEnvironment1.Commands("DonarCard_Report").Execute
    Donar_Card_Report3.Show
End Sub

' =========================================================================
' RIGHT SECTION: FILTER BY DATE RANGE (USING ISS_DT)
' =========================================================================

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
' Command3_Click: Print by Date (STRICT ACTIVE CHECK WITH UPPER)
' -------------------------------------------------------------------------
Private Sub Command3_Click()
    Dim sql As String
    Dim rsCheck As New ADODB.Recordset
    
    If Combo3.Text = "" Then
        MsgBox "Please select a date filter type (Single Date / Date Range).", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' ---> FIX: Apply UPPER() on BOTH tables C and D in Base Query <---
    Dim baseSQL As String
    baseSQL = "SELECT C.*, D.nm, D.bl_grp FROM DONAR_CARD C, DONAR_DETAIL D " & _
              "WHERE C.d_id = D.d_id AND UPPER(C.status)='ACTIVE' AND UPPER(D.status)='ACTIVE'"

    If Combo3.Text = "Single Date" Then
        sql = baseSQL & " AND C.iss_dt = TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
        
    ElseIf Combo3.Text = "Date Range" Then
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "The 'From' date cannot be later than the 'To' date.", vbCritical, "Invalid Date Range"
            Exit Sub
        End If
        sql = baseSQL & " AND C.iss_dt >= TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') " & _
                        " AND C.iss_dt <= TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY')"
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
    If DataEnvironment1.rsDonarCard_Report.State = 1 Then
        DataEnvironment1.rsDonarCard_Report.Close
    End If
    On Error GoTo 0

    DataEnvironment1.Commands("DonarCard_Report").CommandText = sql
    DataEnvironment1.Commands("DonarCard_Report").Execute
    
    ' Showing selective report format
    Donar_Card_Report3.Show
End Sub

