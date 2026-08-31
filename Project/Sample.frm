VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Sample 
   Caption         =   "Sample"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   17805
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   11760
      TabIndex        =   15
      Top             =   1920
      Width           =   5055
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
         TabIndex        =   17
         Text            =   "Combo2"
         Top             =   840
         Width           =   2535
      End
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
         TabIndex        =   16
         Text            =   "Combo4"
         Top             =   1680
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
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
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1440
         Picture         =   "Sample.frx":0000
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2400
      TabIndex        =   9
      Top             =   6720
      Width           =   14415
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   12000
         Picture         =   "Sample.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   9600
         Picture         =   "Sample.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   7200
         Picture         =   "Sample.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   4800
         Picture         =   "Sample.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   2400
         Picture         =   "Sample.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   120
         Picture         =   "Sample.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   8415
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
         Left            =   3840
         TabIndex        =   10
         Text            =   "Combo3"
         Top             =   3840
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpCollection 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   3120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
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
         Format          =   139132929
         CurrentDate     =   46053
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
         Left            =   3840
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
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
         Left            =   3840
         TabIndex        =   13
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label Label7 
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
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Status "
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
         Left            =   720
         TabIndex        =   6
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Collection Date"
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
         TabIndex        =   5
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Collector ID"
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
         TabIndex        =   4
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Component Name "
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
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Visit ID "
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
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sample ID "
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
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Sample.frx":20B9B
      Height          =   2655
      Left            =   2280
      TabIndex        =   14
      Top             =   7920
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   26
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
      DataMember      =   "Sample"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "S_ID"
         Caption         =   "SAMPLE ID"
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
         DataField       =   "V_ID"
         Caption         =   "VISIT ID"
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
         DataField       =   "C_NM"
         Caption         =   "COMPONENT NAME"
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
         DataField       =   "U_ID"
         Caption         =   "COLLECTOR ID"
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
         DataField       =   "C_DATE"
         Caption         =   "DATE"
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
      BeginProperty Column05 
         DataField       =   "STATUS"
         Caption         =   "STATUS"
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
      BeginProperty Column06 
         DataField       =   "STATUS2"
         Caption         =   "STATUS2"
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
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Sample.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmSample (FINAL CODE WITH BUTTON/CONTROL RESTRICTIONS)
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Sample_Report_form.Show
End Sub



' =========================================================
' 1. FORM LOAD & INITIALIZATION
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' 1. Combos Setup
    Combo3.Clear
    Combo3.AddItem "Collected"
    Combo3.AddItem "Damaged"
    Combo3.AddItem "Pending"

    ' Search Combos
    Combo2.Clear
    Combo2.AddItem "Sample ID"
    Combo2.AddItem "Visit ID"
    Combo4.Enabled = False

    ' Fixed Labels
    Label9.Caption = "Whole Blood"

    ' 2. Initial Reset
    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
End Sub

' =========================================================
' 2. MODE / BUTTON / CONTROL RESTRICTIONS
' =========================================================
Private Sub SetStartupMode()
    mMode = "STARTUP"

    Call EnableEntryControls(False)
    Call EnableSearchControls(True)

    CmdAddNew.Enabled = True
    cmdSearch.Enabled = False
    cmdreport.Enabled = True

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = False

    Combo2.SetFocus
End Sub

Private Sub SetAddMode()
    mMode = "ADD"

    Call EnableEntryControls(True)
    Call EnableSearchControls(False)

    CmdAddNew.Enabled = False
    cmdSearch.Enabled = False
    cmdreport.Enabled = False

    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = True

    Combo1.SetFocus
End Sub

Private Sub SetEditMode()
    mMode = "EDIT"

    Call EnableEntryControls(True)
    Call EnableSearchControls(True)

    CmdAddNew.Enabled = True
    cmdSearch.Enabled = True
    cmdreport.Enabled = True

    CmdSave.Enabled = False
    CmdUpdate.Enabled = True
    CmdDelete.Enabled = True
    CmdClear.Enabled = True

    Combo1.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag
    dtpCollection.Enabled = bFlag
    Combo3.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo2.Enabled = bFlag
    Combo4.Enabled = bFlag
End Sub

' --- 1. AUTO ID LOGIC (MAX Formula) ---
Private Sub Auto_SampleID()
    Dim rsMax As New ADODB.Recordset
    Dim nextNum As Long

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(s_id, 4))), 0) FROM SAMPLE"
    Set rsMax = C.Execute(sql)

    nextNum = rsMax.Fields(0).Value + 1
    Label7.Caption = "SAM" & Format(nextNum, "0000")
    Label8.Caption = "WB-" & Label7.Caption

    rsMax.Close
End Sub

' --- 2. SMART VISIT ID LOADER (Duplicate Rokne Wala Logic) ---
Private Sub Fill_VisitIDs()
    Dim rsVis As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    Combo1.Clear

    sql = "SELECT v_id FROM DONATION_VISIT WHERE status='Active' AND " & _
          "v_id NOT IN (SELECT v_id FROM SAMPLE WHERE status2='Active') " & _
          "ORDER BY v_id ASC"

    rsVis.Open sql, C
    Do While Not rsVis.EOF
        Combo1.AddItem rsVis!v_id
        rsVis.MoveNext
    Loop
    rsVis.Close
End Sub

' --- 3. ADD NEW BUTTON ---
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_SampleID
    Call Fill_VisitIDs
    Call SetAddMode
    Combo1.SetFocus
End Sub

' --- 4. SAVE BUTTON ---
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first!", vbExclamation
        Exit Sub
    End If

    If Label7.Caption = "" Then
        MsgBox "Please click 'Add New' first!", vbExclamation
        Exit Sub
    End If

    If Combo1.Text = "" Or Combo3.Text = "" Then
        MsgBox "Visit ID aur Status select karna zaruri hai!", vbExclamation
        Exit Sub
    End If

    If C.State = 0 Then Call CONN

    sql = "INSERT INTO SAMPLE (s_id, v_id, c_nm, u_id, c_date, status, status2) " & _
          "VALUES (" & _
          "'" & Label7.Caption & "', " & _
          "'" & Combo1.Text & "', " & _
          "'" & Label9.Caption & "', " & _
          "'" & Label8.Caption & "', " & _
          "TO_DATE('" & Format(dtpCollection.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "'" & Combo3.Text & "', " & _
          "'Active')"

    C.Execute sql
    MsgBox "Sample Collected Successfully!", vbInformation

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical
End Sub

' --- 5. UPDATE BUTTON ---
Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please Search a record first!", vbExclamation
        Exit Sub
    End If

    If Label7.Caption = "" Then
        MsgBox "Please Search a record first!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Update this Sample record?", vbQuestion + vbYesNo) = vbYes Then
        Call CONN

        sql = "UPDATE SAMPLE SET " & _
              "c_date=TO_DATE('" & Format(dtpCollection.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "status='" & Combo3.Text & "' " & _
              "WHERE s_id='" & Label7.Caption & "'"

        C.Execute sql
        MsgBox "Record Updated!", vbInformation
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical
End Sub

' --- 6. SOFT DELETE (Using status2) ---
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please Search a record first!", vbExclamation
        Exit Sub
    End If

    If Label7.Caption = "" Then
        MsgBox "Please Search a record first!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete (Deactivate) this Sample?", vbQuestion + vbYesNo) = vbYes Then
        Call CONN

        sql = "UPDATE SAMPLE SET status2 = 'Inactive' WHERE s_id = '" & Label7.Caption & "'"
        C.Execute sql

        MsgBox "Sample Record Deleted (Inactive)!", vbInformation
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical
End Sub

' --- 7. SEARCH LOGIC (Filter status2) ---
Private Sub Combo2_Click()
    If Combo2.Text = "" Then Exit Sub

    Combo4.Clear
    Combo4.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    sql = ""
    If Combo2.Text = "Sample ID" Then
        sql = "SELECT s_id FROM SAMPLE WHERE status2='Active' ORDER BY s_id ASC"
    ElseIf Combo2.Text = "Visit ID" Then
        sql = "SELECT DISTINCT v_id FROM SAMPLE WHERE status2='Active' ORDER BY v_id ASC"
    End If

    If sql = "" Then Exit Sub

    rsS.Open sql, C

    Do While Not rsS.EOF
        Combo4.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
    Set rsS = Nothing
End Sub

Private Sub Combo4_Click()
    If Trim(Combo4.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo2.Text = "" Or Combo4.Text = "" Then
        MsgBox "Search Criteria Select Karein!", vbExclamation
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo2.Text = "Sample ID" Then
        sql = "SELECT * FROM SAMPLE WHERE s_id='" & Combo4.Text & "' AND status2='Active'"
    ElseIf Combo2.Text = "Visit ID" Then
        sql = "SELECT * FROM SAMPLE WHERE v_id='" & Combo4.Text & "' AND status2='Active'"
    End If

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label7.Caption = rsSearch!s_id

        Combo1.Clear
        Combo1.AddItem rsSearch!v_id
        Combo1.Text = rsSearch!v_id

        Label9.Caption = rsSearch!c_nm
        Label8.Caption = rsSearch!u_id

        If Not IsNull(rsSearch!c_date) Then dtpCollection.Value = rsSearch!c_date
        Combo3.Text = rsSearch!Status

        MsgBox "Record Found!", vbInformation
        Call SetEditMode
    Else
        MsgBox "Record not found (or Inactive).", vbCritical
    End If

    rsSearch.Close
End Sub

' --- 8. REFRESH GRID (DataEnvironment Logic) ---
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsSample.State = adStateOpen Then
        DataEnvironment1.rsSample.Close
    End If

    DataEnvironment1.Commands("Sample").CommandText = _
        "SELECT * FROM SAMPLE WHERE (status2 = 'Active' OR status2 IS NULL) ORDER BY s_id ASC"

    DataEnvironment1.Sample
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Sample"
    DataGrid1.Refresh
End Sub

' --- 9. CLEAR FORM ---
Public Sub Clear_form()
    Label7.Caption = ""
    Label8.Caption = ""
    Combo1.Clear
    Combo1.Text = ""
    Combo3.ListIndex = -1
    Combo3.Text = ""
    dtpCollection.Value = Date
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo4.Clear
    Combo4.Text = ""
    Combo4.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

' --- 10. NAVIGATION (ENTER KEY) ---
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo3.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub dtpCollection_Change()
    If dtpCollection.Value > Date Then
        MsgBox "Future Date Allowed Nahi Hai!", vbExclamation
        dtpCollection.Value = Date
    End If
End Sub

