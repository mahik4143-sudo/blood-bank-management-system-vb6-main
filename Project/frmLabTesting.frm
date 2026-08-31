VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLabTesting 
   Caption         =   "Lab Testing- Blood Screening"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   17895
   Begin VB.Frame fraFinalResult 
      Caption         =   "Final Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   1080
      TabIndex        =   22
      Top             =   6120
      Width           =   11895
      Begin VB.TextBox txtFinalResult 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "PASS /  FAIL"
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4335
      Left            =   13320
      TabIndex        =   17
      Top             =   1560
      Width           =   4575
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
         Left            =   2040
         TabIndex        =   19
         Text            =   "Combo3"
         Top             =   1920
         Width           =   2295
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
         Left            =   2040
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Image CmdSearch 
         Height          =   735
         Left            =   1080
         Picture         =   "frmLabTesting.frx":0000
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2415
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
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
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
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLabTesting.frx":3CDF
      Height          =   1575
      Left            =   1080
      TabIndex        =   16
      Top             =   8760
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   2778
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
      DataMember      =   "Lab_Testing"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "T_ID"
         Caption         =   "TESTING ID"
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
      BeginProperty Column02 
         DataField       =   "HIV_S"
         Caption         =   "HIV "
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
         DataField       =   "HCV_S"
         Caption         =   "HCV"
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
         DataField       =   "MALA"
         Caption         =   "MALARIA"
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
         DataField       =   "FIN_RES"
         Caption         =   "FINAL RESULT"
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
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2550.047
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1080
      TabIndex        =   14
      Top             =   7680
      Width           =   16695
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   14280
         Picture         =   "frmLabTesting.frx":3CFE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   11640
         Picture         =   "frmLabTesting.frx":F2AB
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   8760
         Picture         =   "frmLabTesting.frx":12E54
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   5880
         Picture         =   "frmLabTesting.frx":1662B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   3000
         Picture         =   "frmLabTesting.frx":19CA2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   120
         Picture         =   "frmLabTesting.frx":1D6F2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame fraSampleResult 
      Caption         =   "Sample && Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   11895
      Begin VB.Frame fraMalaria 
         Caption         =   "Malaria Test"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton optMal_Neg 
            Caption         =   "Negative"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   13
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton optMal_Pos 
            Caption         =   "Positive"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame fraHCVTest 
         Caption         =   "HCV Test"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   4440
         TabIndex        =   6
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton optHCV_Neg 
            Caption         =   "Negative"
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
            TabIndex        =   11
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optHCV_Pos 
            Caption         =   "Positive"
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
            TabIndex        =   10
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame fraHIVTest 
         Caption         =   "HIV Test"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton optHIV_Neg 
            Caption         =   "Negative"
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
            Left            =   480
            TabIndex        =   9
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optHIV_Pos 
            Caption         =   "Positive"
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
            Left            =   480
            TabIndex        =   8
            Top             =   480
            Width           =   1935
         End
      End
   End
   Begin VB.Frame fraTestDetails 
      Caption         =   "Test Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   11895
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
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblTestID 
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
         TabIndex        =   15
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblSampleID 
         Caption         =   "Sample ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label TestID 
         Caption         =   "Test ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmLabTesting.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmLabTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmLabTesting (COMPLETE FINAL CODE WITH RESTRICTIONS)
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Lab_Testing_Report.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' 1. Form Setup
    Call Fill_SampleIDs

    ' 2. Search Combos
    Combo2.Clear
    Combo2.AddItem "Test ID"
    Combo2.AddItem "Sample ID"
    Combo3.Enabled = False

    ' 3. Result Box Locked
    txtFinalResult.Locked = True

    ' 4. Initial Reset
    Call Clear_form
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

    optHIV_Pos.Enabled = bFlag
    optHIV_Neg.Enabled = bFlag

    optHCV_Pos.Enabled = bFlag
    optHCV_Neg.Enabled = bFlag

    optMal_Pos.Enabled = bFlag
    optMal_Neg.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo2.Enabled = bFlag
    Combo3.Enabled = bFlag
End Sub

' --- 1. SMART SAMPLE ID LOADER (Duplicate Rokne Wala) ---
Private Sub Fill_SampleIDs()
    Dim rsSam As New ADODB.Recordset
    Combo1.Clear

    If C.State = 0 Then Call CONN

    sql = "SELECT s_id FROM SAMPLE WHERE status2='Active' AND " & _
          "s_id NOT IN (SELECT s_id FROM LAB_TESTING WHERE status='Active') " & _
          "ORDER BY s_id ASC"

    rsSam.Open sql, C
    Do While Not rsSam.EOF
        Combo1.AddItem rsSam!s_id
        rsSam.MoveNext
    Loop
    rsSam.Close
End Sub

' --- 2. AUTO TEST ID (MAX Formula) ---
Private Sub Auto_TestID()
    Dim rsMax As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(t_id, 4))), 0) FROM LAB_TESTING"
    Set rsMax = C.Execute(sql)

    lblTestID.Caption = "TST" & Format(rsMax.Fields(0).Value + 1, "0000")

    rsMax.Close
End Sub

' --- 3. RESULT CALCULATION LOGIC (Auto-Color) ---
Private Sub CheckFinalResult()
    Dim isFail As Boolean
    isFail = False

    If optHIV_Pos.Value = True Then isFail = True
    If optHCV_Pos.Value = True Then isFail = True
    If optMal_Pos.Value = True Then isFail = True

    If isFail = True Then
        txtFinalResult.Text = "FAIL"
        txtFinalResult.BackColor = vbRed
        txtFinalResult.ForeColor = vbWhite
    Else
        txtFinalResult.Text = "PASS"
        txtFinalResult.BackColor = RGB(0, 153, 76)
        txtFinalResult.ForeColor = vbWhite
    End If
End Sub

' Radio Buttons Click Events
Private Sub optHIV_Pos_Click()
    Call CheckFinalResult
End Sub

Private Sub optHIV_Neg_Click()
    Call CheckFinalResult
End Sub

Private Sub optHCV_Pos_Click()
    Call CheckFinalResult
End Sub

Private Sub optHCV_Neg_Click()
    Call CheckFinalResult
End Sub

Private Sub optMal_Pos_Click()
    Call CheckFinalResult
End Sub

Private Sub optMal_Neg_Click()
    Call CheckFinalResult
End Sub

' --- 4. ADD NEW BUTTON ---
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_TestID
    Call Fill_SampleIDs
    Call SetAddMode
    Combo1.SetFocus
End Sub

' --- 5. SAVE BUTTON ---
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first!", vbExclamation
        Exit Sub
    End If

    If lblTestID.Caption = "" Then
        MsgBox "Click Add New first!", vbExclamation
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MsgBox "Select Sample ID!", vbExclamation
        Exit Sub
    End If

    If (optHIV_Pos.Value = False And optHIV_Neg.Value = False) Or _
       (optHCV_Pos.Value = False And optHCV_Neg.Value = False) Or _
       (optMal_Pos.Value = False And optMal_Neg.Value = False) Then
        MsgBox "Please complete all tests (HIV, HCV, Malaria)!", vbExclamation
        Exit Sub
    End If

    If C.State = 0 Then Call CONN

    Dim vHIV As String, vHCV As String, vMal As String
    If optHIV_Pos.Value = True Then vHIV = "POS" Else vHIV = "NEG"
    If optHCV_Pos.Value = True Then vHCV = "POS" Else vHCV = "NEG"
    If optMal_Pos.Value = True Then vMal = "POS" Else vMal = "NEG"

    sql = "INSERT INTO LAB_TESTING (t_id, s_id, hiv_s, hcv_s, mala, fin_res, status) " & _
          "VALUES (" & _
          "'" & lblTestID.Caption & "', " & _
          "'" & Combo1.Text & "', " & _
          "'" & vHIV & "', " & _
          "'" & vHCV & "', " & _
          "'" & vMal & "', " & _
          "'" & txtFinalResult.Text & "', " & _
          "'Active')"

    C.Execute sql
    MsgBox "Test results saved successfully!", vbInformation

    If txtFinalResult.Text = "FAIL" Then
        MsgBox "Warning: This blood unit is unsafe. Discard immediately.", vbCritical
    End If

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical
End Sub

' --- 6. UPDATE BUTTON ---
Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first!", vbExclamation
        Exit Sub
    End If

    If lblTestID.Caption = "" Or lblTestID.Caption = "TST0000" Then
        MsgBox "Please search for a record first!", vbExclamation
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MsgBox "Please select a Sample ID!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Do you want to update this test record?", vbQuestion + vbYesNo) = vbYes Then

        If C.State = 0 Then Call CONN

        Dim vHIV As String, vHCV As String, vMal As String

        If optHIV_Pos.Value = True Then vHIV = "POS" Else vHIV = "NEG"
        If optHCV_Pos.Value = True Then vHCV = "POS" Else vHCV = "NEG"
        If optMal_Pos.Value = True Then vMal = "POS" Else vMal = "NEG"

        Dim vFinal As String
        If vHIV = "POS" Or vHCV = "POS" Or vMal = "POS" Then
            vFinal = "FAIL"
        Else
            vFinal = "PASS"
        End If

        sql = "UPDATE LAB_TESTING SET " & _
              "s_id='" & Combo1.Text & "', " & _
              "hiv_s='" & vHIV & "', " & _
              "hcv_s='" & vHCV & "', " & _
              "mala='" & vMal & "', " & _
              "fin_res='" & vFinal & "' " & _
              "WHERE t_id='" & lblTestID.Caption & "'"

        C.Execute sql

        MsgBox "Record updated successfully!", vbInformation

        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical
End Sub

' --- 7. SOFT DELETE (Inactive Logic) ---
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Search a record first!", vbExclamation
        Exit Sub
    End If

    If lblTestID.Caption = "" Then
        MsgBox "Search a record first!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Do you want to deactivate this test record?", vbQuestion + vbYesNo) = vbYes Then
        Call CONN

        sql = "UPDATE LAB_TESTING SET status = 'Inactive' WHERE t_id = '" & lblTestID.Caption & "'"
        C.Execute sql

        MsgBox "Record deactivated successfully!", vbInformation
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical
End Sub

' --- 8. SEARCH LOGIC ---
Private Sub Combo2_Click()
    If Combo2.Text = "" Then Exit Sub

    Combo3.Clear
    Combo3.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo2.Text = "Test ID" Then
        sql = "SELECT t_id FROM LAB_TESTING WHERE status='Active' ORDER BY t_id ASC"
    ElseIf Combo2.Text = "Sample ID" Then
        sql = "SELECT s_id FROM LAB_TESTING WHERE status='Active' ORDER BY s_id ASC"
    End If

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo3.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo3_Click()
    If Trim(Combo3.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo2.Text = "" Or Combo3.Text = "" Then
        MsgBox "Search criteria missing!", vbExclamation
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo2.Text = "Test ID" Then
        sql = "SELECT * FROM LAB_TESTING WHERE t_id='" & Combo3.Text & "' AND status='Active'"
    ElseIf Combo2.Text = "Sample ID" Then
        sql = "SELECT * FROM LAB_TESTING WHERE s_id='" & Combo3.Text & "' AND status='Active'"
    End If

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        lblTestID.Caption = rsSearch!t_id

        Combo1.Clear
        Combo1.AddItem rsSearch!s_id
        Combo1.Text = rsSearch!s_id

        If rsSearch!hiv_s = "POS" Then
            optHIV_Pos.Value = True
        Else
            optHIV_Neg.Value = True
        End If

        If rsSearch!hcv_s = "POS" Then
            optHCV_Pos.Value = True
        Else
            optHCV_Neg.Value = True
        End If

        If rsSearch!mala = "POS" Then
            optMal_Pos.Value = True
        Else
            optMal_Neg.Value = True
        End If

        Call CheckFinalResult

        MsgBox "Record found successfully!", vbInformation
        Call SetEditMode
    Else
        MsgBox "Record not found or inactive.", vbCritical
    End If
    rsSearch.Close
End Sub

' --- 9. REFRESH GRID ---
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsLab_Testing.State = adStateOpen Then
        DataEnvironment1.rsLab_Testing.Close
    End If

    DataEnvironment1.Commands("Lab_Testing").CommandText = _
        "SELECT * FROM LAB_TESTING " & _
        "WHERE (status = 'Active' OR status IS NULL) " & _
        "ORDER BY TO_NUMBER(SUBSTR(t_id,4)) ASC"

    DataEnvironment1.Lab_Testing
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Lab_Testing"
    DataGrid1.Refresh
End Sub

' --- 10. CLEAR FORM ---
Public Sub Clear_form()
    lblTestID.Caption = ""
    Combo1.Clear
    Combo1.Text = ""

    optHIV_Pos.Value = False
    optHIV_Neg.Value = False
    optHCV_Pos.Value = False
    optHCV_Neg.Value = False
    optMal_Pos.Value = False
    optMal_Neg.Value = False

    txtFinalResult.Text = ""
    txtFinalResult.BackColor = vbWhite
    txtFinalResult.ForeColor = vbBlack

    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    Combo3.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

