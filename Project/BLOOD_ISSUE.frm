VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmblood_issue 
   Caption         =   "BLOOD ISSUE"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   17475
   Begin VB.Frame Frame4 
      Caption         =   "Search value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   12480
      TabIndex        =   26
      Top             =   1560
      Width           =   5415
      Begin VB.ComboBox Combo4 
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
         Left            =   2280
         TabIndex        =   28
         Text            =   "Combo4"
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   2280
         TabIndex        =   27
         Text            =   "Combo5"
         Top             =   2040
         Width           =   2535
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
         Left            =   360
         TabIndex        =   30
         Top             =   960
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
         Left            =   360
         TabIndex        =   29
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1440
         Picture         =   "BLOOD_ISSUE.frx":0000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2160
      TabIndex        =   18
      Top             =   6600
      Width           =   13455
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   11400
         Picture         =   "BLOOD_ISSUE.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   9240
         Picture         =   "BLOOD_ISSUE.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   6840
         Picture         =   "BLOOD_ISSUE.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   4560
         Picture         =   "BLOOD_ISSUE.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   2280
         Picture         =   "BLOOD_ISSUE.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   0
         Picture         =   "BLOOD_ISSUE.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   11535
      Begin VB.ComboBox Combo6 
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
         Left            =   8400
         TabIndex        =   24
         Text            =   "Combo6"
         Top             =   3720
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
         Left            =   2280
         TabIndex        =   21
         Text            =   "Combo3"
         Top             =   2760
         Width           =   2535
      End
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
         Left            =   2280
         TabIndex        =   20
         Text            =   "Combo2"
         Top             =   1920
         Width           =   2535
      End
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
         Left            =   8400
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   2040
         Width           =   2895
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
         Left            =   8400
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1200
         Width           =   2895
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
         Left            =   8400
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   480
         Width           =   2895
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
         Left            =   2280
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   3600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   173998081
         CurrentDate     =   46061
      End
      Begin VB.Label Label15 
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
         Left            =   8400
         TabIndex        =   23
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label14 
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
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label13 
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label12 
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label11 
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Mode"
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
         Left            =   6000
         TabIndex        =   10
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Bill No"
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
         Left            =   6000
         TabIndex        =   9
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Fee "
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
         Left            =   6000
         TabIndex        =   8
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Received By "
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
         Left            =   5880
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Issued By"
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
         Left            =   5880
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Issue Date "
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
         Left            =   240
         TabIndex        =   5
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Cross Match "
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
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Stock ID"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Request ID "
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Issue ID"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   600
      TabIndex        =   25
      Top             =   8040
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "BLOOD_ISSUE.frx":20B9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmblood_issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Blood_issue_report.Show
End Sub

' =========================================================
' FORM LOAD & INITIALIZATION
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' Dropdowns Setup
    Call Fill_RequestIDs
    Call Fill_StockIDs

    ' Cross Match Dropdown
    Combo3.Clear
    Combo3.AddItem "MATCH"
    Combo3.AddItem "NOT MATCH"

    ' Mode Dropdown
    Combo6.Clear
    Combo6.AddItem "Cash"
    Combo6.AddItem "UPI"
    Combo6.AddItem "Card"
    Combo6.AddItem "Free / Camp Issue"

    ' Search Setup
    Combo4.Clear
    Combo4.AddItem "Issue ID"
    Combo4.AddItem "Request ID"
    Combo4.AddItem "Bill No"
    Combo5.Enabled = False

    DTPicker1.Value = Date

    Call ClearForm
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

    Combo4.SetFocus
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

    Combo1.SetFocus
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

    Combo1.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag    ' Request ID
    Combo2.Enabled = bFlag    ' Stock ID
    Combo3.Enabled = bFlag    ' Cross Match
    DTPicker1.Enabled = bFlag
    Text1.Enabled = bFlag     ' Issued By
    Text2.Enabled = bFlag     ' Received By
    Text3.Enabled = bFlag     ' Fee
    Combo6.Enabled = bFlag    ' Mode
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo4.Enabled = bFlag
    Combo5.Enabled = bFlag
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

Private Function IsDecimalValue(ByVal s As String) As Boolean
    Dim i As Integer
    Dim dotCount As Integer
    Dim ch As String

    s = Trim(s)
    If s = "" Then
        IsDecimalValue = False
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dotCount = dotCount + 1
            If dotCount > 1 Then
                IsDecimalValue = False
                Exit Function
            End If
        ElseIf ch < "0" Or ch > "9" Then
            IsDecimalValue = False
            Exit Function
        End If
    Next i

    IsDecimalValue = True
End Function

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(Label14.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Action Required"
        Exit Function
    End If

    If Trim(Combo1.Text) = "" Then
        MsgBox "Please select Request ID.", vbExclamation, "Validation Error"
        Combo1.SetFocus
        Exit Function
    End If

    If Trim(Combo2.Text) = "" Then
        MsgBox "Please select Stock ID.", vbExclamation, "Validation Error"
        Combo2.SetFocus
        Exit Function
    End If

    If Trim(Combo3.Text) = "" Then
        MsgBox "Please select Cross Match result.", vbExclamation, "Validation Error"
        Combo3.SetFocus
        Exit Function
    End If

    If UCase(Trim(Combo3.Text)) = "NOT MATCH" Then
        MsgBox "Cannot issue blood because cross match result is NOT MATCH.", vbCritical, "Safety Protocol"
        Combo3.SetFocus
        Exit Function
    End If

    If DTPicker1.Value > Date Then
        MsgBox "Issue Date cannot be greater than the current date.", vbExclamation, "Date Validation"
        DTPicker1.SetFocus
        Exit Function
    End If

    If Trim(Text3.Text) = "" Then
        MsgBox "Please enter Fee.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Not IsDecimalValue(Text3.Text) Then
        MsgBox "Fee must be numeric.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Val(Text3.Text) < 0 Then
        MsgBox "Fee cannot be negative.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Trim(Combo6.Text) = "" Then
        MsgBox "Please select Mode.", vbExclamation, "Validation Error"
        Combo6.SetFocus
        Exit Function
    End If

    If Trim(Text1.Text) <> "" And Len(Trim(Text1.Text)) > 25 Then
        MsgBox "Issued By cannot exceed 25 characters.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Trim(Text2.Text) <> "" And Len(Trim(Text2.Text)) > 25 Then
        MsgBox "Received By cannot exceed 25 characters.", vbExclamation, "Validation Error"
        Text2.SetFocus
        Exit Function
    End If

    If Len(Trim(Text3.Text)) > 8 Then
        ' Table NUMBER(8,2) - simple practical check
    End If

    ValidateForm = True
End Function

' =========================================================
' SMART LOADERS
' =========================================================
Private Sub Fill_RequestIDs()
    Dim rsReq As New ADODB.Recordset

    If C.State = 0 Then Call CONN
    Combo1.Clear

    sql = "SELECT req_id FROM PATIENT_REQUEST WHERE status='Active' AND " & _
          "req_id NOT IN (SELECT req_id FROM BLOOD_ISSUE WHERE status='Active') " & _
          "ORDER BY req_id ASC"

    rsReq.Open sql, C
    Do While Not rsReq.EOF
        Combo1.AddItem rsReq!req_id
        rsReq.MoveNext
    Loop
    rsReq.Close
End Sub

Private Sub Fill_StockIDs()
    Dim rsStk As New ADODB.Recordset

    If C.State = 0 Then Call CONN
    Combo2.Clear

    sql = "SELECT st_id FROM BLOOD_STOCK WHERE status='Active' AND avail='AVAILABLE' ORDER BY st_id ASC"
    rsStk.Open sql, C
    Do While Not rsStk.EOF
        Combo2.AddItem rsStk!st_id
        rsStk.MoveNext
    Loop
    rsStk.Close
End Sub

' =========================================================
' REFRESH GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    Dim rsGrid As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    rsGrid.CursorLocation = adUseClient
    sql = "SELECT * FROM BLOOD_ISSUE WHERE (status='Active' OR status IS NULL) " & _
          "ORDER BY TO_NUMBER(SUBSTR(issue_id,4)) ASC"
    rsGrid.Open sql, C, adOpenStatic, adLockReadOnly

    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = rsGrid
    DataGrid1.Refresh
End Sub

' =========================================================
' AUTO ISSUE ID & BILL NO
' =========================================================
Private Sub Auto_GenerateIDs()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(issue_id, 4))), 0) FROM BLOOD_ISSUE"
    Set rsMax = C.Execute(sql)
    Label14.Caption = "ISS" & Format(rsMax.Fields(0).Value + 1, "0000")
    rsMax.Close

    sql = "SELECT NVL(MAX(bill), 1000) FROM BLOOD_ISSUE"
    Set rsMax = C.Execute(sql)
    Label15.Caption = rsMax.Fields(0).Value + 1
    rsMax.Close
End Sub

' =========================================================
' NAVIGATION & RESTRICTIONS
' =========================================================
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo3.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPicker1.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Text1.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo6.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii = 46 Then
        If InStr(Text3.Text, ".") > 0 Then KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

' =========================================================
' ADD / SAVE / UPDATE / DELETE
' =========================================================
Private Sub cmdAddNew_Click()
    Call ClearForm
    Call Auto_GenerateIDs
    Call Fill_RequestIDs
    Call Fill_StockIDs
    Call SetAddMode
    Combo1.SetFocus
End Sub

' =========================================================
' SAVE BUTTON (With Auto Stock Update Logic)
' =========================================================
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    ' 1. Validations
    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    ' Database Connection check
    If C.State = 0 Then Call CONN

    ' 2. BLOOD_ISSUE table mein record INSERT karna
    sql = "INSERT INTO BLOOD_ISSUE (issue_id, req_id, st_id, x_match, issue_dt, issue_by, recv_by, fee, bill, mod, pay, status) " & _
          "VALUES (" & _
          "'" & Esc(Label14.Caption) & "', " & _
          "'" & Esc(Combo1.Text) & "', " & _
          "'" & Esc(Combo2.Text) & "', " & _
          "'" & Esc(Combo3.Text) & "', " & _
          "TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "'" & Esc(Text1.Text) & "', " & _
          "'" & Esc(Text2.Text) & "', " & _
          Val(Text3.Text) & ", " & _
          Val(Label15.Caption) & ", " & _
          "'" & Esc(Combo6.Text) & "', " & _
          "'PAID', " & _
          "'Active')"
          
    C.Execute sql

    ' -------------------------------------------------------------------------
    ' 3. STOCK UPDATE LOGIC (Jo aapne specifically manga tha)
    ' Jo Stock ID Combo2 mein select hui hai, uska status 'ISSUED' kar do
    ' -------------------------------------------------------------------------
    sql = "UPDATE BLOOD_STOCK SET avail='ISSUED' WHERE st_id='" & Esc(Combo2.Text) & "'"
    C.Execute sql
    ' -------------------------------------------------------------------------

    MsgBox "Blood issued successfully." & vbCrLf & "Bill No: " & Label15.Caption, vbInformation, "Success"

    ' Form reset karna
    Call ClearForm
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(Label14.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If MsgBox("Do you want to update issue details?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        Call CONN

        sql = "UPDATE BLOOD_ISSUE SET " & _
              "x_match='" & Esc(Combo3.Text) & "', " & _
              "issue_dt=TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "issue_by='" & Esc(Text1.Text) & "', " & _
              "recv_by='" & Esc(Text2.Text) & "', " & _
              "fee=" & Val(Text3.Text) & ", " & _
              "mod='" & Esc(Combo6.Text) & "' " & _
              "WHERE issue_id='" & Esc(Label14.Caption) & "'"

        C.Execute sql

        MsgBox "Issue record updated successfully.", vbInformation, "Updated"
        Call RefreshGrid
        Call ClearForm
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(Label14.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If MsgBox("Do you want to cancel this blood issue? The stock will return to AVAILABLE.", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN

        sql = "UPDATE BLOOD_ISSUE SET status='Inactive' WHERE issue_id='" & Esc(Label14.Caption) & "'"
        C.Execute sql

        sql = "UPDATE BLOOD_STOCK SET avail='AVAILABLE' WHERE st_id='" & Esc(Combo2.Text) & "'"
        C.Execute sql

        MsgBox "Blood issue cancelled successfully. Stock returned to inventory.", vbInformation, "Deleted"
        Call RefreshGrid
        Call ClearForm
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' SEARCH
' =========================================================
Private Sub Combo4_Click()
    If Combo4.Text = "" Then Exit Sub

    Combo5.Clear
    Combo5.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo4.Text = "Issue ID" Then sql = "SELECT issue_id FROM BLOOD_ISSUE WHERE status='Active' ORDER BY issue_id ASC"
    If Combo4.Text = "Request ID" Then sql = "SELECT DISTINCT req_id FROM BLOOD_ISSUE WHERE status='Active' ORDER BY req_id ASC"
    If Combo4.Text = "Bill No" Then sql = "SELECT DISTINCT bill FROM BLOOD_ISSUE WHERE status='Active' ORDER BY bill ASC"

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo5.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo5_Click()
    If Trim(Combo5.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo4.Text = "" Or Combo5.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Error"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo4.Text = "Issue ID" Then sql = "SELECT * FROM BLOOD_ISSUE WHERE issue_id='" & Esc(Combo5.Text) & "' AND status='Active'"
    If Combo4.Text = "Request ID" Then sql = "SELECT * FROM BLOOD_ISSUE WHERE req_id='" & Esc(Combo5.Text) & "' AND status='Active'"
    If Combo4.Text = "Bill No" Then sql = "SELECT * FROM BLOOD_ISSUE WHERE bill=" & Val(Combo5.Text) & " AND status='Active'"

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label14.Caption = rsSearch!issue_id

        Combo1.Clear
        Combo1.AddItem rsSearch!req_id
        Combo1.Text = rsSearch!req_id

        Combo2.Clear
        Combo2.AddItem rsSearch!st_id
        Combo2.Text = rsSearch!st_id

        Combo3.Text = rsSearch!x_match
        If Not IsNull(rsSearch!issue_dt) Then DTPicker1.Value = rsSearch!issue_dt
        Text1.Text = "" & rsSearch!issue_by
        Text2.Text = "" & rsSearch!recv_by
        Text3.Text = "" & rsSearch!fee
        Label15.Caption = "" & rsSearch!bill
        Combo6.Text = "" & rsSearch!mod

        MsgBox "Record found successfully.", vbInformation, "Success"
        Call SetEditMode
    Else
        MsgBox "Record not found or inactive.", vbCritical, "Not Found"
    End If
    rsSearch.Close
End Sub

' =========================================================
' CLEAR
' =========================================================
Private Sub ClearForm()
    Label14.Caption = ""
    Label15.Caption = ""
    Combo1.ListIndex = -1
    Combo1.Text = ""
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.ListIndex = -1
    Combo3.Text = ""
    Combo6.ListIndex = -1
    Combo6.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    DTPicker1.Value = Date
    Combo4.ListIndex = -1
    Combo4.Text = ""
    Combo5.Clear
    Combo5.Text = ""
    Combo5.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call ClearForm
    Call SetStartupMode
End Sub

' =========================================================
' LOST FOCUS FORMATTING
' =========================================================
Private Sub Text1_LostFocus()
    Text1.Text = ProperCaseText(Text1.Text)
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = ProperCaseText(Text2.Text)
End Sub

