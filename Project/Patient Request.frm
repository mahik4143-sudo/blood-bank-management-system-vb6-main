VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpatient_request 
   Caption         =   "PATIENT REQUEST"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   18315
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   720
      TabIndex        =   33
      Top             =   7920
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
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
   Begin VB.Frame Frame4 
      Height          =   4695
      Left            =   13920
      TabIndex        =   26
      Top             =   1560
      Width           =   4455
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
         Left            =   2160
         TabIndex        =   31
         Text            =   "Combo5"
         Top             =   2640
         Width           =   2055
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
         Left            =   2160
         TabIndex        =   28
         Text            =   "Combo2"
         Top             =   840
         Width           =   2055
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
         Left            =   2160
         TabIndex        =   27
         Text            =   "Combo3"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
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
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1080
         Picture         =   "Patient Request.frx":0000
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      TabIndex        =   22
      Top             =   6600
      Width           =   17655
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   15240
         Picture         =   "Patient Request.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   12360
         Picture         =   "Patient Request.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   9240
         Picture         =   "Patient Request.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2415
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   6240
         Picture         =   "Patient Request.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   3240
         Picture         =   "Patient Request.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   360
         Picture         =   "Patient Request.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   13095
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
         Left            =   3600
         TabIndex        =   25
         Text            =   "Combo4"
         Top             =   2520
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
         Left            =   9480
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text8 
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
         Left            =   9480
         TabIndex        =   21
         Text            =   "Text8"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox Text7 
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
         Left            =   9480
         TabIndex        =   20
         Text            =   "Text7"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox Text6 
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
         Left            =   9480
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text5 
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
         Left            =   9480
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   1320
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   3840
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   143261697
         CurrentDate     =   46061
      End
      Begin VB.TextBox Text4 
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
         Left            =   3600
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   3120
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
         Left            =   3600
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1800
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
         Left            =   3600
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1080
         Width           =   2895
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
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label13 
         Height          =   375
         Left            =   13560
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label12 
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Doctor Name"
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
         Left            =   7440
         TabIndex        =   11
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Platelets "
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
         Left            =   7440
         TabIndex        =   10
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "HB% "
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
         Left            =   7440
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Units "
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
         Left            =   7440
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Date "
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
         Left            =   1200
         TabIndex        =   7
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Pressure "
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
         Left            =   1200
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Component "
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
         Left            =   1200
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Ward "
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
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hospital "
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Patient ID "
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
         Left            =   7320
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Patient Request.frx":20B9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmpatient_request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    patient_request_report.Show
End Sub





' =========================================================
' FORM LOAD
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    Call Fill_PatientIDs

    Combo4.Clear
    Combo4.AddItem "Whole Blood"
    Combo4.AddItem "Packed Red Blood Cells (RBC)"
    Combo4.AddItem "Fresh Frozen Plasma (FFP)"
    Combo4.AddItem "Platelet Concentrate"

    Combo2.Clear
    Combo2.AddItem "Request ID"
    Combo2.AddItem "Patient ID"
    Combo2.AddItem "Hospital"
    Combo3.Enabled = False
    Combo5.Clear
    Combo5.Enabled = False
    Label15.Caption = ""

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

    Combo2.SetFocus
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
    Combo1.Enabled = bFlag
    Text1.Enabled = bFlag
    Text2.Enabled = bFlag
    Combo4.Enabled = bFlag
    Text4.Enabled = bFlag
    DTPicker1.Enabled = bFlag
    Text5.Enabled = bFlag
    Text6.Enabled = bFlag
    Text7.Enabled = bFlag
    Text8.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo2.Enabled = bFlag
    Combo3.Enabled = bFlag
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

Private Function LowerCaseText(ByVal s As String) As String
    LowerCaseText = LCase(Trim(s))
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Integer
    s = Trim(s)

    If s = "" Then
        IsDigitsOnly = False
        Exit Function
    End If

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i

    IsDigitsOnly = True
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

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(Label14.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Action Required"
        Exit Function
    End If

    If Trim(Combo1.Text) = "" Then
        MsgBox "Please select Patient ID.", vbExclamation, "Validation Error"
        Combo1.SetFocus
        Exit Function
    End If

    If Trim(Text1.Text) = "" Then
        MsgBox "Please enter Hospital Name.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Trim(Text2.Text) = "" Then
        MsgBox "Please enter Ward.", vbExclamation, "Validation Error"
        Text2.SetFocus
        Exit Function
    End If

    If Trim(Combo4.Text) = "" Then
        MsgBox "Please select Component.", vbExclamation, "Validation Error"
        Combo4.SetFocus
        Exit Function
    End If

    If Trim(Text4.Text) <> "" Then
        If IsValidBP(Text4.Text) = False Then
            MsgBox "Please enter Blood Pressure in proper format like 120/80.", vbExclamation, "Validation Error"
            Text4.SetFocus
            Exit Function
        End If
    End If

    If Trim(Text5.Text) = "" Then
        MsgBox "Please enter Units.", vbExclamation, "Validation Error"
        Text5.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(Text5.Text) Then
        MsgBox "Units must be numeric.", vbExclamation, "Validation Error"
        Text5.SetFocus
        Exit Function
    End If

    If Val(Text5.Text) <= 0 Then
        MsgBox "Units must be greater than zero.", vbExclamation, "Validation Error"
        Text5.SetFocus
        Exit Function
    End If

    If Trim(Text6.Text) <> "" Then
        If Not IsDecimalValue(Text6.Text) Then
            MsgBox "HB% must be numeric.", vbExclamation, "Validation Error"
            Text6.SetFocus
            Exit Function
        End If
    End If

    If Trim(Text7.Text) <> "" Then
        If Not IsDigitsOnly(Text7.Text) Then
            MsgBox "Platelets must be numeric.", vbExclamation, "Validation Error"
            Text7.SetFocus
            Exit Function
        End If
    End If

    If DTPicker1.Value > Date Then
        MsgBox "Date cannot be greater than the current date.", vbExclamation, "Date Validation"
        DTPicker1.SetFocus
        Exit Function
    End If

    If Trim(Text1.Text) <> "" And Len(Trim(Text1.Text)) > 100 Then
        MsgBox "Hospital Name is too long.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Trim(Text2.Text) <> "" And Len(Trim(Text2.Text)) > 50 Then
        MsgBox "Ward is too long.", vbExclamation, "Validation Error"
        Text2.SetFocus
        Exit Function
    End If

    If Trim(Text4.Text) <> "" And Len(Trim(Text4.Text)) > 7 Then
        MsgBox "Blood Pressure is too long.", vbExclamation, "Validation Error"
        Text4.SetFocus
        Exit Function
    End If

    If Trim(Text8.Text) <> "" And Len(Trim(Text8.Text)) > 100 Then
        MsgBox "Doctor Name is too long.", vbExclamation, "Validation Error"
        Text8.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

' =========================================================
' PATIENT ID LOADER
' =========================================================
Private Sub Fill_PatientIDs()
    Dim rsPat As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    Combo1.Clear

    sql = "SELECT pa_id FROM PATIENT WHERE status='Active' AND " & _
          "pa_id NOT IN (SELECT pa_id FROM PATIENT_REQUEST WHERE status='Active') " & _
          "ORDER BY pa_id ASC"

    rsPat.Open sql, C
    Do While Not rsPat.EOF
        Combo1.AddItem rsPat!pa_id
        rsPat.MoveNext
    Loop
    rsPat.Close
End Sub

' =========================================================
' GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    Dim rsGrid As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    rsGrid.CursorLocation = adUseClient
    sql = "SELECT * FROM PATIENT_REQUEST WHERE (status='Active' OR status IS NULL) ORDER BY TO_NUMBER(SUBSTR(req_id,4)) ASC"
    rsGrid.Open sql, C, adOpenStatic, adLockReadOnly

    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = rsGrid
    DataGrid1.Refresh
End Sub

' =========================================================
' AUTO ID
' =========================================================
Private Sub Auto_RequestID()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(req_id, 4))), 0) FROM PATIENT_REQUEST"
    Set rsMax = C.Execute(sql)

    Label14.Caption = "REQ" & Format(rsMax.Fields(0).Value + 1, "0000")
    rsMax.Close
End Sub

' =========================================================
' NAVIGATION & KEY RESTRICTIONS
' =========================================================
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
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
        Combo4.SetFocus
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Text4.Text) <> "" Then
            If IsValidBP(Text4.Text) = False Then
                MsgBox "Please enter Blood Pressure in proper format like 120/80.", vbExclamation, "Validation Error"
                Text4.SetFocus
                KeyAscii = 0
                Exit Sub
            End If
        End If

        DTPicker1.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Text5.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
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
    Call ClearForm
    Call Auto_RequestID
    Call Fill_PatientIDs
    Call SetAddMode
    Combo1.SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If C.State = 0 Then Call CONN

    Dim valHB As String, valPlat As String
    If Trim(Text6.Text) = "" Then
        valHB = "NULL"
    Else
        valHB = CStr(Val(Text6.Text))
    End If

    If Trim(Text7.Text) = "" Then
        valPlat = "NULL"
    Else
        valPlat = CStr(Val(Text7.Text))
    End If

    sql = "INSERT INTO PATIENT_REQUEST (req_id, pa_id, hos_nm, wrd, bed, units, comp, hb_p, pre, plat, dt, d_nm, status) " & _
          "VALUES (" & _
          "'" & Esc(Label14.Caption) & "', " & _
          "'" & Esc(Combo1.Text) & "', " & _
          "'" & Esc(Text1.Text) & "', " & _
          "'" & Esc(Text2.Text) & "', " & _
          "NULL, " & _
          Val(Text5.Text) & ", " & _
          "'" & Esc(Combo4.Text) & "', " & _
          valHB & ", " & _
          "'" & Esc(Text4.Text) & "', " & _
          valPlat & ", " & _
          "TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "'" & Esc(Text8.Text) & "', " & _
          "'Active')"

    C.Execute sql

    MsgBox "Patient request saved successfully.", vbInformation, "Success"

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

    If MsgBox("Do you want to update this request?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        Call CONN

        Dim valHB As String, valPlat As String
        If Trim(Text6.Text) = "" Then
            valHB = "NULL"
        Else
            valHB = CStr(Val(Text6.Text))
        End If

        If Trim(Text7.Text) = "" Then
            valPlat = "NULL"
        Else
            valPlat = CStr(Val(Text7.Text))
        End If

        sql = "UPDATE PATIENT_REQUEST SET " & _
              "pa_id='" & Esc(Combo1.Text) & "', " & _
              "hos_nm='" & Esc(Text1.Text) & "', " & _
              "wrd='" & Esc(Text2.Text) & "', " & _
              "units=" & Val(Text5.Text) & ", " & _
              "comp='" & Esc(Combo4.Text) & "', " & _
              "hb_p=" & valHB & ", " & _
              "pre='" & Esc(Text4.Text) & "', " & _
              "plat=" & valPlat & ", " & _
              "dt=TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "d_nm='" & Esc(Text8.Text) & "' " & _
              "WHERE req_id='" & Esc(Label14.Caption) & "'"

        C.Execute sql

        MsgBox "Request updated successfully.", vbInformation, "Updated"
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

    If MsgBox("Do you want to deactivate this request?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN
        sql = "UPDATE PATIENT_REQUEST SET status='Inactive' WHERE req_id='" & Esc(Label14.Caption) & "'"
        C.Execute sql
        MsgBox "Request deactivated successfully.", vbInformation, "Deleted"
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
Private Sub Combo2_Click()
    If Combo2.Text = "" Then Exit Sub

    Combo3.Clear
    Combo3.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo2.Text = "Request ID" Then sql = "SELECT req_id FROM PATIENT_REQUEST WHERE status='Active' ORDER BY req_id ASC"
    If Combo2.Text = "Patient ID" Then sql = "SELECT DISTINCT pa_id FROM PATIENT_REQUEST WHERE status='Active' ORDER BY pa_id ASC"
    If Combo2.Text = "Hospital" Then sql = "SELECT DISTINCT hos_nm FROM PATIENT_REQUEST WHERE status='Active' ORDER BY hos_nm ASC"

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo3.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo3_Click()
    Dim rsReq As New ADODB.Recordset

    If Trim(Combo3.Text) = "" Then Exit Sub

    cmdSearch.Enabled = True

    Combo5.Clear
    Combo5.Enabled = False
    Label15.Caption = ""

    If Combo2.Text = "Hospital" Then
        Label15.Caption = "Request ID"
        Combo5.Enabled = True

        If C.State = 0 Then Call CONN

        sql = "SELECT req_id FROM PATIENT_REQUEST " & _
              "WHERE hos_nm='" & Esc(Combo3.Text) & "' AND status='Active' " & _
              "ORDER BY req_id ASC"

        rsReq.Open sql, C

        Do While Not rsReq.EOF
            Combo5.AddItem rsReq!req_id
            rsReq.MoveNext
        Loop

        rsReq.Close
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo2.Text = "" Or Combo3.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Error"
        Exit Sub
    End If

    If Combo2.Text = "Hospital" Then
        If Trim(Combo5.Text) = "" Then
            MsgBox "Please select Request ID.", vbExclamation, "Search Error"
            Combo5.SetFocus
            Exit Sub
        End If
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo2.Text = "Request ID" Then
        sql = "SELECT * FROM PATIENT_REQUEST WHERE req_id='" & Esc(Combo3.Text) & "' AND status='Active'"
    End If

    If Combo2.Text = "Patient ID" Then
        sql = "SELECT * FROM PATIENT_REQUEST WHERE pa_id='" & Esc(Combo3.Text) & "' AND status='Active'"
    End If

    If Combo2.Text = "Hospital" Then
        sql = "SELECT * FROM PATIENT_REQUEST WHERE req_id='" & Esc(Combo5.Text) & "' AND status='Active'"
    End If

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label14.Caption = rsSearch!req_id

        Combo1.Clear
        Combo1.AddItem rsSearch!pa_id
        Combo1.Text = rsSearch!pa_id

        Text1.Text = "" & rsSearch!hos_nm
        Text2.Text = "" & rsSearch!wrd
        Combo4.Text = "" & rsSearch!comp
        Text4.Text = "" & rsSearch!pre
        If Not IsNull(rsSearch!dt) Then DTPicker1.Value = rsSearch!dt
        Text5.Text = "" & rsSearch!units
        If Not IsNull(rsSearch!hb_p) Then Text6.Text = "" & rsSearch!hb_p Else Text6.Text = ""
        If Not IsNull(rsSearch!plat) Then Text7.Text = "" & rsSearch!plat Else Text7.Text = ""
        Text8.Text = "" & rsSearch!d_nm

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
    Combo1.ListIndex = -1
    Combo1.Text = ""
    Combo4.ListIndex = -1
    Combo4.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    DTPicker1.Value = Date
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    
    
    Combo3.Enabled = False
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

Private Sub Text4_LostFocus()
    If Trim(Text4.Text) = "" Then Exit Sub

    If IsValidBP(Text4.Text) = False Then
        MsgBox "Please enter Blood Pressure in proper format like 120/80.", vbExclamation, "Validation Error"
        Text4.SetFocus
    End If
End Sub

Private Sub Text8_LostFocus()
    Text8.Text = LowerCaseText(Text8.Text)
End Sub

