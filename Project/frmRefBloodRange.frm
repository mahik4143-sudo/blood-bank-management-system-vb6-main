VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRefBloodRange 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REF BLOOD RANGE MASTER"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18180
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   18180
   Begin VB.Frame Frame4 
      Height          =   3855
      Left            =   13680
      TabIndex        =   23
      Top             =   1560
      Width           =   4335
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
         TabIndex        =   25
         Text            =   "Combo2"
         Top             =   1560
         Width           =   2055
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
         Left            =   2040
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Image cmdSearch 
         Height          =   495
         Left            =   1080
         Picture         =   "frmRefBloodRange.frx":0000
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2055
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
         TabIndex        =   27
         Top             =   1560
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
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1080
      TabIndex        =   20
      Top             =   7080
      Width           =   16935
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   14400
         Picture         =   "frmRefBloodRange.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   11520
         Picture         =   "frmRefBloodRange.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   8640
         Picture         =   "frmRefBloodRange.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   5760
         Picture         =   "frmRefBloodRange.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   3000
         Picture         =   "frmRefBloodRange.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   240
         Picture         =   "frmRefBloodRange.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRefBloodRange.frx":20B9B
      Height          =   2175
      Left            =   1080
      TabIndex        =   17
      Top             =   8160
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   3836
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
      DataMember      =   "Ref_Blood_Range"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "REF_ID"
         Caption         =   "REFRENCE ID"
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
         DataField       =   "CM_NM"
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
      BeginProperty Column02 
         DataField       =   "ME_UN"
         Caption         =   "UNIT"
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
         DataField       =   "MN_V_M"
         Caption         =   "MIN MALE RANGE"
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
         DataField       =   "MX_V_M"
         Caption         =   "MAX MALE RANGE"
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
         DataField       =   "MN_V_F"
         Caption         =   "MIN FEMALE RANGE"
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
         DataField       =   "MX_V_F"
         Caption         =   "MAX FEMALE RANGE"
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
      BeginProperty Column07 
         DataField       =   "AGE"
         Caption         =   "AGE"
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
      BeginProperty Column08 
         DataField       =   "DESCR"
         Caption         =   "DESCRIPTION"
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
      BeginProperty Column09 
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
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDesc 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   1080
      TabIndex        =   16
      Top             =   5520
      Width           =   12135
      Begin RichTextLib.RichTextBox txtDesc 
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmRefBloodRange.frx":20BBA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraFemale 
      Caption         =   "Female Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   7200
      TabIndex        =   7
      Top             =   3600
      Width           =   6015
      Begin VB.TextBox txtMxF 
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
         Left            =   2760
         TabIndex        =   15
         Text            =   "txtMxF"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtMnF 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2760
         TabIndex        =   14
         Text            =   "txtMnF"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Female Max Value"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Female Min Value"
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
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame fraMale 
      Caption         =   "Male Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   1080
      TabIndex        =   6
      Top             =   3600
      Width           =   5895
      Begin VB.TextBox txtMxM 
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
         Left            =   2520
         TabIndex        =   11
         Text            =   "txtMxM"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtMnM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2520
         TabIndex        =   10
         Text            =   "txtMnM"
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Male Max Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Male Min Value"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame fraBasic 
      Caption         =   "Basic Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   12135
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         TabIndex        =   22
         Text            =   "txtAge"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox cboUnit 
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
         TabIndex        =   5
         Text            =   "cboUnit"
         Top             =   1080
         Width           =   3015
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
         Left            =   8760
         TabIndex        =   4
         Text            =   "cboComponent"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label ref 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Refrences ID"
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
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Measuerment Unit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Age"
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
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   -120
      Picture         =   "frmRefBloodRange.frx":20C40
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmRefBloodRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmRefBloodRange
' FINAL CODE WITH FULL RESTRICTIONS
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Ref_blood_Report.Show
End Sub

' =========================================================
' FORM LOAD
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    cboComponent.Clear
    cboComponent.AddItem "Hemoglobin"
    cboComponent.AddItem "Temperature"
    cboComponent.AddItem "Pulse"
    cboComponent.AddItem "Weight"
    cboComponent.AddItem "Blood Pressure"

    cboUnit.Clear
    cboUnit.AddItem "g/dL"
    cboUnit.AddItem "Fahrenheit"
    cboUnit.AddItem "bpm"
    cboUnit.AddItem "Kg"
    cboUnit.AddItem "mmHg"

    Combo1.Clear
    Combo1.AddItem "Reference ID"
    Combo1.AddItem "Component Name"
    Combo2.Enabled = False

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

    txtAge.SetFocus
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

    txtAge.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    txtAge.Enabled = bFlag
    cboComponent.Enabled = bFlag
    cboUnit.Enabled = bFlag
    txtMnM.Enabled = bFlag
    txtMxM.Enabled = bFlag
    txtMnF.Enabled = bFlag
    txtMxF.Enabled = bFlag
    txtDesc.Enabled = bFlag
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

Private Function IsValidAgeText(ByVal s As String) As Boolean
    Dim i As Integer
    Dim ch As String

    s = Trim(s)
    If s = "" Then
        IsValidAgeText = False
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Then
            ' allowed
        Else
            IsValidAgeText = False
            Exit Function
        End If
    Next i

    IsValidAgeText = True
End Function

Private Function IsValidBP(ByVal sBP As String) As Boolean
    Dim arr() As String

    sBP = Trim(sBP)

    If InStr(1, sBP, "/") = 0 Then
        IsValidBP = False
        Exit Function
    End If

    arr = Split(sBP, "/")

    If UBound(arr) <> 1 Then
        IsValidBP = False
        Exit Function
    End If

    If arr(0) = "" Or arr(1) = "" Then
        IsValidBP = False
        Exit Function
    End If

    If Not IsNumeric(arr(0)) Then
        IsValidBP = False
        Exit Function
    End If

    If Not IsNumeric(arr(1)) Then
        IsValidBP = False
        Exit Function
    End If

    IsValidBP = True
End Function

Private Function IsValidRangeValue(ByVal sValue As String, ByVal sComponent As String) As Boolean
    sValue = Trim(sValue)

    If sValue = "" Then
        IsValidRangeValue = True
        Exit Function
    End If

    If UCase(Trim(sComponent)) = UCase("Blood Pressure") Then
        IsValidRangeValue = IsValidBP(sValue)
    Else
        IsValidRangeValue = IsDecimalValue(sValue)
    End If
End Function

Private Function UnitMatchesComponent(ByVal sComp As String, ByVal sUnit As String) As Boolean
    UnitMatchesComponent = False

    Select Case Trim(sComp)
        Case "Hemoglobin"
            If Trim(sUnit) = "g/dL" Then UnitMatchesComponent = True
        Case "Temperature"
            If Trim(sUnit) = "Fahrenheit" Then UnitMatchesComponent = True
        Case "Pulse"
            If Trim(sUnit) = "bpm" Then UnitMatchesComponent = True
        Case "Weight"
            If Trim(sUnit) = "Kg" Then UnitMatchesComponent = True
        Case "Blood Pressure"
            If Trim(sUnit) = "mmHg" Then UnitMatchesComponent = True
    End Select
End Function

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(ref.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Validation Error"
        Exit Function
    End If

    If Trim(txtAge.Text) = "" Then
        MsgBox "Please enter Age.", vbExclamation, "Validation Error"
        txtAge.SetFocus
        Exit Function
    End If

    If IsValidAgeText(txtAge.Text) = False Then
        MsgBox "Age must contain digits or '-' only, like 18 or 18-45.", vbExclamation, "Validation Error"
        txtAge.SetFocus
        Exit Function
    End If

    If Trim(cboComponent.Text) = "" Then
        MsgBox "Please select Component Name.", vbExclamation, "Validation Error"
        cboComponent.SetFocus
        Exit Function
    End If

    If Trim(cboUnit.Text) = "" Then
        MsgBox "Please select Measurement Unit.", vbExclamation, "Validation Error"
        cboUnit.SetFocus
        Exit Function
    End If

    If UnitMatchesComponent(cboComponent.Text, cboUnit.Text) = False Then
        MsgBox "Selected Measurement Unit does not match the selected Component.", vbExclamation, "Validation Error"
        cboUnit.SetFocus
        Exit Function
    End If

    If IsValidRangeValue(txtMnM.Text, cboComponent.Text) = False Then
        MsgBox "Invalid Male Min Value.", vbExclamation, "Validation Error"
        txtMnM.SetFocus
        Exit Function
    End If

    If IsValidRangeValue(txtMxM.Text, cboComponent.Text) = False Then
        MsgBox "Invalid Male Max Value.", vbExclamation, "Validation Error"
        txtMxM.SetFocus
        Exit Function
    End If

    If IsValidRangeValue(txtMnF.Text, cboComponent.Text) = False Then
        MsgBox "Invalid Female Min Value.", vbExclamation, "Validation Error"
        txtMnF.SetFocus
        Exit Function
    End If

    If IsValidRangeValue(txtMxF.Text, cboComponent.Text) = False Then
        MsgBox "Invalid Female Max Value.", vbExclamation, "Validation Error"
        txtMxF.SetFocus
        Exit Function
    End If

    If Trim(txtDesc.Text) <> "" And Len(Trim(txtDesc.Text)) > 500 Then
        MsgBox "Description is too long.", vbExclamation, "Validation Error"
        txtDesc.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

' =========================================================
' REFRESH GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsRef_Blood_Range.State = 1 Then
        DataEnvironment1.rsRef_Blood_Range.Requery
    Else
        DataEnvironment1.Ref_Blood_Range
    End If

    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Ref_Blood_Range"
    DataGrid1.Refresh
End Sub

' =========================================================
' AUTO ID
' =========================================================
Private Sub Auto_ReferenceID()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(ref_id, 4))), 0) FROM REF_BLOOD_RANGE"
    Set rsMax = C.Execute(sql)

    ref.Caption = "REF" & Format(rsMax.Fields(0).Value + 1, "0000")
    rsMax.Close
End Sub

' =========================================================
' NAVIGATION & KEY RESTRICTIONS
' =========================================================
Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboComponent.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub cboComponent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboUnit.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMnM.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtMnM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMxM.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If UCase(Trim(cboComponent.Text)) = UCase("Blood Pressure") Then
        If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Else
        If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub txtMxM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMnF.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If UCase(Trim(cboComponent.Text)) = UCase("Blood Pressure") Then
        If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Else
        If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub txtMnF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMxF.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If UCase(Trim(cboComponent.Text)) = UCase("Blood Pressure") Then
        If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Else
        If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub txtMxF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDesc.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If UCase(Trim(cboComponent.Text)) = UCase("Blood Pressure") Then
        If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Else
        If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

' =========================================================
' ADD / SAVE / UPDATE / DELETE
' =========================================================
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_ReferenceID
    Call SetAddMode
    txtAge.SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If C.State = 0 Then Call CONN

    Dim mM As String, xM As String, mF As String, xF As String
    If Trim(txtMnM.Text) = "" Then mM = "NULL" Else mM = "'" & Esc(txtMnM.Text) & "'"
    If Trim(txtMxM.Text) = "" Then xM = "NULL" Else xM = "'" & Esc(txtMxM.Text) & "'"
    If Trim(txtMnF.Text) = "" Then mF = "NULL" Else mF = "'" & Esc(txtMnF.Text) & "'"
    If Trim(txtMxF.Text) = "" Then xF = "NULL" Else xF = "'" & Esc(txtMxF.Text) & "'"

    sql = "INSERT INTO REF_BLOOD_RANGE (ref_id, cm_nm, me_un, mn_v_m, mx_v_m, mn_v_f, mx_v_f, age, descr, status) " & _
          "VALUES (" & _
          "'" & Esc(ref.Caption) & "', " & _
          "'" & Esc(cboComponent.Text) & "', " & _
          "'" & Esc(cboUnit.Text) & "', " & _
          mM & ", " & _
          xM & ", " & _
          mF & ", " & _
          xF & ", " & _
          "'" & Esc(txtAge.Text) & "', " & _
          "'" & Esc(txtDesc.Text) & "', " & _
          "'Active')"

    C.Execute sql

    MsgBox "Reference Range saved successfully.", vbInformation, "Red Cross Blood Bank"
    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical, "Red Cross Blood Bank - System Error"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    If Trim(ref.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If MsgBox("Do you want to update this Reference Range?", vbQuestion + vbYesNo, "Red Cross Blood Bank") = vbYes Then
        Call CONN

        Dim mM As String, xM As String, mF As String, xF As String
        If Trim(txtMnM.Text) = "" Then mM = "NULL" Else mM = "'" & Esc(txtMnM.Text) & "'"
        If Trim(txtMxM.Text) = "" Then xM = "NULL" Else xM = "'" & Esc(txtMxM.Text) & "'"
        If Trim(txtMnF.Text) = "" Then mF = "NULL" Else mF = "'" & Esc(txtMnF.Text) & "'"
        If Trim(txtMxF.Text) = "" Then xF = "NULL" Else xF = "'" & Esc(txtMxF.Text) & "'"

        sql = "UPDATE REF_BLOOD_RANGE SET " & _
              "cm_nm='" & Esc(cboComponent.Text) & "', " & _
              "me_un='" & Esc(cboUnit.Text) & "', " & _
              "mn_v_m=" & mM & ", " & _
              "mx_v_m=" & xM & ", " & _
              "mn_v_f=" & mF & ", " & _
              "mx_v_f=" & xF & ", " & _
              "age='" & Esc(txtAge.Text) & "', " & _
              "descr='" & Esc(txtDesc.Text) & "' " & _
              "WHERE ref_id='" & Esc(ref.Caption) & "'"

        C.Execute sql
        MsgBox "Record updated successfully.", vbInformation, "Red Cross Blood Bank"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "Red Cross Blood Bank - System Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    If Trim(ref.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    If MsgBox("Do you want to deactivate this Reference Range?", vbQuestion + vbYesNo, "Red Cross Blood Bank") = vbYes Then
        Call CONN
        sql = "UPDATE REF_BLOOD_RANGE SET status='Inactive' WHERE ref_id='" & Esc(ref.Caption) & "'"
        C.Execute sql
        MsgBox "Record deactivated successfully.", vbInformation, "Red Cross Blood Bank"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "Red Cross Blood Bank - System Error"
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

    If Combo1.Text = "Reference ID" Then sql = "SELECT ref_id FROM REF_BLOOD_RANGE WHERE status='Active' ORDER BY ref_id ASC"
    If Combo1.Text = "Component Name" Then sql = "SELECT DISTINCT cm_nm FROM REF_BLOOD_RANGE WHERE status='Active' ORDER BY cm_nm ASC"

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
        MsgBox "Please select search criteria.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo1.Text = "Reference ID" Then sql = "SELECT * FROM REF_BLOOD_RANGE WHERE ref_id='" & Esc(Combo2.Text) & "' AND status='Active'"
    If Combo1.Text = "Component Name" Then sql = "SELECT * FROM REF_BLOOD_RANGE WHERE cm_nm='" & Esc(Combo2.Text) & "' AND status='Active'"

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        ref.Caption = rsSearch!ref_id
        cboComponent.Text = rsSearch!CM_NM
        cboUnit.Text = rsSearch!me_un
        txtAge.Text = "" & rsSearch!age
        txtMnM.Text = "" & rsSearch!MN_V_M
        txtMxM.Text = "" & rsSearch!MX_V_M
        txtMnF.Text = "" & rsSearch!MN_V_F
        txtMxF.Text = "" & rsSearch!MX_V_F
        txtDesc.Text = "" & rsSearch!descr

        MsgBox "Record found successfully.", vbInformation, "Red Cross Blood Bank"
        Call SetEditMode
    Else
        MsgBox "Record not found. It may be inactive.", vbCritical, "Red Cross Blood Bank"
    End If
    rsSearch.Close
End Sub

' =========================================================
' CLEAR
' =========================================================
Private Sub Clear_form()
    ref.Caption = ""
    txtAge.Text = ""
    cboComponent.ListIndex = -1
    cboComponent.Text = ""
    cboUnit.ListIndex = -1
    cboUnit.Text = ""
    txtMnM.Text = ""
    txtMxM.Text = ""
    txtMnF.Text = ""
    txtMxF.Text = ""
    txtDesc.Text = ""
    Combo1.ListIndex = -1
    Combo1.Text = ""
    Combo2.Clear
    Combo2.Text = ""
    Combo2.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

' =========================================================
' AUTO UNIT SELECTION
' =========================================================
Private Sub cboComponent_Click()
    Select Case cboComponent.Text
        Case "Hemoglobin"
            cboUnit.Text = "g/dL"
        Case "Temperature"
            cboUnit.Text = "Fahrenheit"
        Case "Pulse"
            cboUnit.Text = "bpm"
        Case "Weight"
            cboUnit.Text = "Kg"
        Case "Blood Pressure"
            cboUnit.Text = "mmHg"
    End Select
End Sub

