VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComponentProcessing 
   Caption         =   "Component Processing"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmComponentProcessing.frx":0000
   ScaleHeight     =   9975
   ScaleWidth      =   18135
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
      ForeColor       =   &H00800000&
      Height          =   4335
      Left            =   13080
      TabIndex        =   16
      Top             =   2040
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
         Left            =   2160
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   1920
         Width           =   2415
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
         Left            =   2160
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image cmdSearch 
         Height          =   855
         Left            =   1680
         Picture         =   "frmComponentProcessing.frx":0342
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2175
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
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label sea_by 
         Caption         =   "Search by"
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
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2400
      TabIndex        =   13
      Top             =   6600
      Width           =   14175
      Begin VB.Image CmdAddNew 
         Height          =   735
         Left            =   0
         Picture         =   "frmComponentProcessing.frx":4021
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image CmdUpdate 
         Height          =   735
         Left            =   2400
         Picture         =   "frmComponentProcessing.frx":74E9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   735
         Left            =   4800
         Picture         =   "frmComponentProcessing.frx":AF39
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   735
         Left            =   7200
         Picture         =   "frmComponentProcessing.frx":E5B0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   735
         Left            =   9600
         Picture         =   "frmComponentProcessing.frx":11D87
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
      Begin VB.Image cmdreport 
         Height          =   735
         Left            =   12000
         Picture         =   "frmComponentProcessing.frx":15930
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmComponentProcessing.frx":20EDD
      Height          =   1815
      Left            =   840
      TabIndex        =   11
      Top             =   7920
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Component_Processing"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "P_ID"
         Caption         =   "PROCESS ID"
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
      BeginProperty Column02 
         DataField       =   "COM_NM"
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
         DataField       =   "UN_ID"
         Caption         =   "UNIT ID"
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
         DataField       =   "EXP"
         Caption         =   "EXPIRY DATE"
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
         DataField       =   "VOL"
         Caption         =   "VOLUME"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9000
      Top             =   8880
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=PRJ2531G/PRJ2531G;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=PRJ2531G/PRJ2531G;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from COMPONENT_PROCESSING"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Component Details"
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
      Height          =   2895
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Width           =   11895
      Begin MSComCtl2.DTPicker dtpExpiry 
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   174784515
         CurrentDate     =   46050
      End
      Begin VB.TextBox txtVolume 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   9
         Top             =   1200
         Width           =   8655
      End
      Begin VB.ComboBox cboComponent 
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
         Left            =   3120
         TabIndex        =   6
         Text            =   "cboComponent"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblUnitID 
         Caption         =   "lblUnitID"
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
         Left            =   8880
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Expiry Date"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Volume(ml):"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Unit ID"
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
         Left            =   6960
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Component Name"
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
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Processing Details"
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
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   11895
      Begin VB.ComboBox cboTestingID 
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
         Left            =   8760
         TabIndex        =   4
         Text            =   "cboTestingID"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblProcessID 
         Caption         =   "lblProcessID"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Lab Testing"
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
         Left            =   6840
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Process ID"
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
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   0
      Picture         =   "frmComponentProcessing.frx":20EFC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19695
   End
End
Attribute VB_Name = "frmComponentProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmComponentProcessing
' FINAL CODE WITH STRICT COMPONENT RULES
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    COMPONENT_PROCESSING_REPORT.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' Component list
    cboComponent.Clear
    cboComponent.AddItem "Whole Blood"
    cboComponent.AddItem "Packed Red Blood Cells (RBC)"
    cboComponent.AddItem "Fresh Frozen Plasma (FFP)"
    cboComponent.AddItem "Platelet Concentrate"

    ' Search combos
    Combo1.Clear
    Combo1.AddItem "Process ID"
    Combo1.AddItem "Test ID"
    Combo1.AddItem "Unit ID"
    Combo2.Enabled = False

    ' Expiry auto only
    dtpExpiry.Enabled = False

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
    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = False
    cmdSearch.Enabled = False
    cmdreport.Enabled = True

    Combo1.SetFocus
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

    cboTestingID.SetFocus
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

    cboTestingID.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    cboTestingID.Enabled = bFlag
    cboComponent.Enabled = bFlag
    txtVolume.Enabled = bFlag
    dtpExpiry.Enabled = False
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag
    Combo2.Enabled = bFlag
End Sub

' =========================================================
' HELPERS
' =========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function IsWholeBlood(ByVal sComp As String) As Boolean
    IsWholeBlood = (UCase(Trim(sComp)) = UCase("Whole Blood"))
End Function

Private Function IsSplitComponent(ByVal sComp As String) As Boolean
    Dim s As String
    s = UCase(Trim(sComp))

    If s = UCase("Packed Red Blood Cells (RBC)") Then IsSplitComponent = True: Exit Function
    If s = UCase("Fresh Frozen Plasma (FFP)") Then IsSplitComponent = True: Exit Function
    If s = UCase("Platelet Concentrate") Then IsSplitComponent = True: Exit Function

    IsSplitComponent = False
End Function

Private Function CountActiveComponents(ByVal sTestID As String, Optional ByVal excludePID As String = "") As Integer
    Dim rsCnt As New ADODB.Recordset

    sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(sTestID) & "' AND status='Active'"
    If Trim(excludePID) <> "" Then
        sql = sql & " AND p_id<>'" & Esc(excludePID) & "'"
    End If

    rsCnt.Open sql, C
    CountActiveComponents = rsCnt!CNT
    rsCnt.Close
End Function

Private Function ComponentExists(ByVal sTestID As String, ByVal sComponent As String, Optional ByVal excludePID As String = "") As Boolean
    Dim rsCnt As New ADODB.Recordset

    sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(sTestID) & "' " & _
          "AND com_nm='" & Esc(sComponent) & "' AND status='Active'"
    If Trim(excludePID) <> "" Then
        sql = sql & " AND p_id<>'" & Esc(excludePID) & "'"
    End If

    rsCnt.Open sql, C
    ComponentExists = (rsCnt!CNT > 0)
    rsCnt.Close
End Function

Private Function WholeBloodExists(ByVal sTestID As String, Optional ByVal excludePID As String = "") As Boolean
    Dim rsCnt As New ADODB.Recordset

    sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(sTestID) & "' " & _
          "AND com_nm='Whole Blood' AND status='Active'"
    If Trim(excludePID) <> "" Then
        sql = sql & " AND p_id<>'" & Esc(excludePID) & "'"
    End If

    rsCnt.Open sql, C
    WholeBloodExists = (rsCnt!CNT > 0)
    rsCnt.Close
End Function

Private Function AnySplitExists(ByVal sTestID As String, Optional ByVal excludePID As String = "") As Boolean
    Dim rsCnt As New ADODB.Recordset

    sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(sTestID) & "' " & _
          "AND com_nm<>'Whole Blood' AND status='Active'"
    If Trim(excludePID) <> "" Then
        sql = sql & " AND p_id<>'" & Esc(excludePID) & "'"
    End If

    rsCnt.Open sql, C
    AnySplitExists = (rsCnt!CNT > 0)
    rsCnt.Close
End Function

Private Function CanUseComponent(ByVal sTestID As String, ByVal sComponent As String, Optional ByVal excludePID As String = "") As Boolean
    Dim totalCnt As Integer

    CanUseComponent = False

    If Trim(sTestID) = "" Or Trim(sComponent) = "" Then Exit Function

    ' Rule 1: Same component duplicate not allowed
    If ComponentExists(sTestID, sComponent, excludePID) Then
        MsgBox "This component already exists for the selected Test ID.", vbExclamation, "Duplicate Component"
        Exit Function
    End If

    totalCnt = CountActiveComponents(sTestID, excludePID)

    ' Rule 2: Whole Blood only once and only if no split component exists
    If IsWholeBlood(sComponent) Then
        If AnySplitExists(sTestID, excludePID) Then
            MsgBox "Whole Blood cannot be processed because other components already exist for this Test ID.", vbCritical, "Process Rule Violation"
            Exit Function
        End If

        If WholeBloodExists(sTestID, excludePID) Then
            MsgBox "Whole Blood already exists for this Test ID.", vbExclamation, "Duplicate Component"
            Exit Function
        End If

        If totalCnt > 0 Then
            MsgBox "Only one Whole Blood record is allowed for a Test ID.", vbCritical, "Process Rule Violation"
            Exit Function
        End If

        CanUseComponent = True
        Exit Function
    End If

    ' Rule 3: Split components only if Whole Blood does not exist
    If IsSplitComponent(sComponent) Then
        If WholeBloodExists(sTestID, excludePID) Then
            MsgBox "This Test ID already has Whole Blood. Split components cannot be added now.", vbCritical, "Process Rule Violation"
            Exit Function
        End If

        If totalCnt >= 3 Then
            MsgBox "Only three unique components are allowed for one Test ID: RBC, FFP, and Platelet.", vbCritical, "Limit Reached"
            Exit Function
        End If

        CanUseComponent = True
        Exit Function
    End If
End Function

Private Function ValidateForm(Optional ByVal excludePID As String = "") As Boolean
    ValidateForm = False

    If Trim(lblProcessID.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Action Required"
        Exit Function
    End If

    If Trim(cboTestingID.Text) = "" Then
        MsgBox "Please select Lab Testing ID.", vbExclamation, "Validation Error"
        cboTestingID.SetFocus
        Exit Function
    End If

    If Trim(cboComponent.Text) = "" Then
        MsgBox "Please select Component Name.", vbExclamation, "Validation Error"
        cboComponent.SetFocus
        Exit Function
    End If

    If Trim(lblUnitID.Caption) = "" Then
        MsgBox "Unit ID was not generated. Please select a valid component.", vbExclamation, "Validation Error"
        cboComponent.SetFocus
        Exit Function
    End If

    If Trim(txtVolume.Text) = "" Then
        MsgBox "Please enter Volume.", vbExclamation, "Validation Error"
        txtVolume.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtVolume.Text) Then
        MsgBox "Volume must be numeric.", vbExclamation, "Validation Error"
        txtVolume.SetFocus
        Exit Function
    End If

    If Val(txtVolume.Text) <= 0 Then
        MsgBox "Volume must be greater than zero.", vbExclamation, "Validation Error"
        txtVolume.SetFocus
        Exit Function
    End If

    If dtpExpiry.Value <= Date Then
        MsgBox "Expiry Date must be greater than the current date.", vbExclamation, "Date Validation"
        dtpExpiry.SetFocus
        Exit Function
    End If

    If CanUseComponent(cboTestingID.Text, cboComponent.Text, excludePID) = False Then Exit Function

    ValidateForm = True
End Function

' =========================================================
' SMART TEST ID LOADER
' =========================================================
Private Sub Fill_TestingIDs()
    Dim rsTest As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim tid As String
    Dim wbCount As Integer
    Dim compCount As Integer

    If C.State = 0 Then Call CONN

    cboTestingID.Clear

    sql = "SELECT t_id FROM LAB_TESTING WHERE fin_res='PASS' AND status='Active' ORDER BY t_id ASC"
    rsTest.Open sql, C

    Do While Not rsTest.EOF
        tid = rsTest!t_id

        sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(tid) & "' AND com_nm='Whole Blood' AND status='Active'"
        Set rsCheck = C.Execute(sql)
        wbCount = rsCheck!CNT

        sql = "SELECT COUNT(*) CNT FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(tid) & "' AND status='Active'"
        Set rsCheck = C.Execute(sql)
        compCount = rsCheck!CNT

        ' Allow Test ID only if:
        ' 1. no Whole Blood exists
        ' 2. total active components < 3
        If wbCount = 0 And compCount < 3 Then
            cboTestingID.AddItem tid
        End If

        rsTest.MoveNext
    Loop

    rsTest.Close
End Sub

' =========================================================
' COMPONENT CLICK
' =========================================================
Private Sub cboComponent_Click()
    If cboComponent.Text = "" Or cboTestingID.Text = "" Then Exit Sub

    If CanUseComponent(cboTestingID.Text, cboComponent.Text) = False Then
        cboComponent.ListIndex = -1
        cboComponent.Text = ""
        lblUnitID.Caption = ""
        txtVolume.Text = ""
        dtpExpiry.Value = Date
        Exit Sub
    End If

    Call Generate_UnitID_Dates
End Sub

' =========================================================
' UNIT ID + EXPIRY + DEFAULT VOLUME
' =========================================================
Private Sub Generate_UnitID_Dates()
    Dim Prefix As String
    Dim Days As Integer
    Dim NumPart As String

    If InStr(cboComponent.Text, "Whole Blood") > 0 Then
        Prefix = "WB"
        Days = 35
        txtVolume.Text = "350"
    ElseIf InStr(cboComponent.Text, "RBC") > 0 Then
        Prefix = "RBC"
        Days = 42
        txtVolume.Text = "250"
    ElseIf InStr(cboComponent.Text, "Plasma") > 0 Then
        Prefix = "FFP"
        Days = 365
        txtVolume.Text = "200"
    ElseIf InStr(cboComponent.Text, "Platelet") > 0 Then
        Prefix = "PLT"
        Days = 5
        txtVolume.Text = "50"
    End If

    dtpExpiry.Value = Date + Days

    If lblProcessID.Caption <> "" Then
        NumPart = Right(lblProcessID.Caption, 4)
        lblUnitID.Caption = Prefix & "-" & NumPart
    End If
End Sub

' =========================================================
' AUTO PROCESS ID
' =========================================================
Private Sub Auto_ProcessID()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(p_id, 4))), 0) FROM COMPONENT_PROCESSING"
    Set rsMax = C.Execute(sql)

    lblProcessID.Caption = "PRC" & Format(rsMax.Fields(0).Value + 1, "0000")

    rsMax.Close
End Sub

' =========================================================
' ADD NEW
' =========================================================
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_ProcessID
    Call Fill_TestingIDs
    Call SetAddMode
    cboTestingID.SetFocus
End Sub

' =========================================================
' SAVE
' =========================================================
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm() = False Then Exit Sub

    If C.State = 0 Then Call CONN

    sql = "INSERT INTO COMPONENT_PROCESSING (p_id, t_id, com_nm, un_id, exp, vol, status) " & _
          "VALUES (" & _
          "'" & Esc(lblProcessID.Caption) & "', " & _
          "'" & Esc(cboTestingID.Text) & "', " & _
          "'" & Esc(cboComponent.Text) & "', " & _
          "'" & Esc(lblUnitID.Caption) & "', " & _
          "TO_DATE('" & Format(dtpExpiry.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          Val(txtVolume.Text) & ", " & _
          "'Active')"

    C.Execute sql

    MsgBox "Component processed successfully." & vbCrLf & "Unit ID: " & lblUnitID.Caption, vbInformation, "Success"

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' UPDATE
' =========================================================
Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(lblProcessID.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If ValidateForm(lblProcessID.Caption) = False Then Exit Sub

    If MsgBox("Are you sure you want to update this component details?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        Call CONN

        sql = "UPDATE COMPONENT_PROCESSING SET " & _
              "t_id='" & Esc(cboTestingID.Text) & "', " & _
              "com_nm='" & Esc(cboComponent.Text) & "', " & _
              "exp=TO_DATE('" & Format(dtpExpiry.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "vol=" & Val(txtVolume.Text) & " " & _
              "WHERE p_id='" & Esc(lblProcessID.Caption) & "'"

        C.Execute sql

        MsgBox "Component details updated successfully.", vbInformation, "Updated"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' DELETE
' =========================================================
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(lblProcessID.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to discard this component?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN
        sql = "UPDATE COMPONENT_PROCESSING SET status='Inactive' WHERE p_id='" & Esc(lblProcessID.Caption) & "'"
        C.Execute sql

        MsgBox "Component discarded successfully.", vbInformation, "Deleted"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' SEARCH
' =========================================================
Private Sub Combo1_Click()
    If Combo1.Text = "" Then Exit Sub

    Combo2.Clear
    Combo2.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo1.Text = "Process ID" Then sql = "SELECT p_id FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY p_id ASC"
    If Combo1.Text = "Test ID" Then sql = "SELECT DISTINCT t_id FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY t_id ASC"
    If Combo1.Text = "Unit ID" Then sql = "SELECT un_id FROM COMPONENT_PROCESSING WHERE status='Active' ORDER BY un_id ASC"

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo2_Click()
    If Trim(Combo2.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Error"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo1.Text = "Process ID" Then sql = "SELECT * FROM COMPONENT_PROCESSING WHERE p_id='" & Esc(Combo2.Text) & "' AND status='Active'"
    If Combo1.Text = "Test ID" Then sql = "SELECT * FROM COMPONENT_PROCESSING WHERE t_id='" & Esc(Combo2.Text) & "' AND status='Active'"
    If Combo1.Text = "Unit ID" Then sql = "SELECT * FROM COMPONENT_PROCESSING WHERE un_id='" & Esc(Combo2.Text) & "' AND status='Active'"

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        lblProcessID.Caption = rsSearch!p_id

        cboTestingID.Clear
        cboTestingID.AddItem rsSearch!t_id
        cboTestingID.Text = rsSearch!t_id

        cboComponent.Text = rsSearch!com_nm
        lblUnitID.Caption = rsSearch!un_id
        dtpExpiry.Value = rsSearch!Exp
        txtVolume.Text = rsSearch!vol

        MsgBox "Record found successfully.", vbInformation, "Search Result"
        Call SetEditMode
    Else
        MsgBox "Record not found or inactive.", vbCritical, "Not Found"
    End If

    rsSearch.Close
End Sub

' =========================================================
' GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsComponent_Processing.State = adStateOpen Then
        DataEnvironment1.rsComponent_Processing.Close
    End If

    DataEnvironment1.Commands("Component_Processing").CommandText = _
        "SELECT * FROM COMPONENT_PROCESSING WHERE (status='Active' OR status IS NULL) " & _
        "ORDER BY TO_NUMBER(SUBSTR(p_id,4)) ASC"

    DataEnvironment1.Component_Processing
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Component_Processing"
    DataGrid1.Refresh
End Sub

' =========================================================
' CLEAR
' =========================================================
Public Sub Clear_form()
    lblProcessID.Caption = ""
    cboTestingID.Clear
    cboTestingID.Text = ""
    cboComponent.ListIndex = -1
    cboComponent.Text = ""
    lblUnitID.Caption = ""
    txtVolume.Text = ""
    dtpExpiry.Value = Date + 1
    Combo1.ListIndex = -1
    Combo1.Text = ""
    Combo2.Clear
    Combo2.Text = ""
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

' =========================================================
' RESTRICTIONS / NAVIGATION
' =========================================================
Private Sub cboTestingID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboComponent.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cboComponent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtVolume.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub dtpExpiry_Change()
    If dtpExpiry.Value <= Date Then
        MsgBox "Expiry Date must be greater than the current date.", vbExclamation, "Date Validation"
        dtpExpiry.Value = Date + 1
    End If
End Sub

