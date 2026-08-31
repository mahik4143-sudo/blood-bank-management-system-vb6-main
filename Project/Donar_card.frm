VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDonarCard 
   Caption         =   "Donar Card"
   ClientHeight    =   10050
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   17820
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   11040
      TabIndex        =   11
      Top             =   1560
      Width           =   6135
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
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   600
         Width           =   3375
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
         TabIndex        =   12
         Text            =   "Combo3"
         Top             =   1560
         Width           =   3375
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
         TabIndex        =   15
         Top             =   600
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1920
         Picture         =   "Donar_card.frx":0000
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1800
      TabIndex        =   8
      Top             =   6000
      Width           =   15375
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   13200
         Picture         =   "Donar_card.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   10680
         Picture         =   "Donar_card.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   8160
         Picture         =   "Donar_card.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   5400
         Picture         =   "Donar_card.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   2760
         Picture         =   "Donar_card.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   120
         Picture         =   "Donar_card.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   8775
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
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   3120
         Width           =   3855
         _ExtentX        =   6800
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
         Format          =   137232385
         CurrentDate     =   46053
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3840
         TabIndex        =   5
         Top             =   2160
         Width           =   3855
         _ExtentX        =   6800
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
         Format          =   137232385
         CurrentDate     =   46053
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Valid Until "
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
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   1200
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Donor ID"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Card Number "
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
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Donar_card.frx":20B9B
      Height          =   2175
      Left            =   1800
      TabIndex        =   10
      Top             =   7320
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Command3"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "C_NO"
         Caption         =   "CARD NO."
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
         DataField       =   "D_ID"
         Caption         =   "DONAR ID"
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
         DataField       =   "ISS_DT"
         Caption         =   "ISSUE DATE"
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
         DataField       =   "VALID"
         Caption         =   "VALID UPTO"
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
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Donar_card.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmDonarCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmDonarCard (FINAL CODE WITH DATE VALIDATION)
' =========================================================

Option Explicit

Dim sql As String
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Donar_card_Report.Show
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Label6.Caption = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MsgBox "Please select Donor ID.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If DTPicker1.Value > Date Then
        MsgBox "Issue Date cannot be greater than the current date.", vbExclamation, "Date Validation"
        DTPicker1.Value = Date
        DTPicker1.SetFocus
        Exit Sub
    End If

    If DTPicker2.Value < Date Then
        MsgBox "Valid Until cannot be less than the current date.", vbExclamation, "Date Validation"
        DTPicker2.SetFocus
        Exit Sub
    End If

    If DTPicker2.Value < DTPicker1.Value Then
        MsgBox "Valid Until cannot be earlier than Issue Date.", vbExclamation, "Date Validation"
        DTPicker2.SetFocus
        Exit Sub
    End If

    If MsgBox("Do you want to update this card record?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then

        If C.State = 0 Then Call CONN

        sql = "UPDATE DONAR_CARD SET " & _
              "d_id='" & Combo1.Text & "', " & _
              "iss_dt=TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "valid=TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "','DD-MM-YYYY') " & _
              "WHERE c_no=" & Val(Label6.Caption)

        C.Execute sql

        MsgBox "Card details updated successfully.", vbInformation, "Update Success"

        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "System Error"
End Sub



Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    Call Fill_DonorIDs

    Combo2.Clear
    Combo2.AddItem "Card No"
    Combo2.AddItem "Donor ID"
    Combo3.Enabled = False

    DTPicker1.Value = Date
    DTPicker2.Value = DateAdd("yyyy", 1, Date)

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
    DTPicker1.Enabled = bFlag
    DTPicker2.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo2.Enabled = bFlag
    Combo3.Enabled = bFlag
End Sub

' =========================================================================
' 1. SMART DONOR ID LOADER (100% BUG FREE LOGIC)
' =========================================================================
' =========================================================================
' 1. SMART DONOR ID LOADER (Using Oracle MINUS - 100% Bug Free)
' =========================================================================
Private Sub Fill_DonorIDs()
    Dim rsD As New ADODB.Recordset
    
    ' Pehle Combo1 ko poori tarah clear karo
    Combo1.Clear
    Combo1.Text = ""

    ' Database Connection Check
    If C.State = 0 Then Call CONN

    ' ---> ORACLE 'MINUS' QUERY: Sabse safe tarika <---
    ' Yeh un sabhi Active Donors ki ID layega jinka Card abhi tak nahi bana hai
    sql = "SELECT D_ID FROM DONAR_DETAIL WHERE STATUS='ACTIVE' " & _
          "MINUS " & _
          "SELECT D_ID FROM DONAR_CARD WHERE STATUS='Active'"

    ' Safe cursor location for RecordCount
    rsD.CursorLocation = adUseClient
    rsD.Open sql, C, adOpenStatic, adLockReadOnly

    ' Check karna ki database se koi data aaya bhi hai ya nahi
    If rsD.RecordCount > 0 Then
        rsD.MoveFirst
        Do While Not rsD.EOF
            ' NULL check aur Trim ka use taaki extra space problem na kare
            If Not IsNull(rsD.Fields(0).Value) Then
                Combo1.AddItem Trim(rsD.Fields(0).Value & "")
            End If
            rsD.MoveNext
        Loop
    Else
        ' Agar sabhi active donors ka card ban chuka hai
        MsgBox "All active donors have already been issued a Donor Card!", vbInformation, "Red Cross Blood Bank"
    End If

    ' Recordset close karna
    rsD.Close
    Set rsD = Nothing
End Sub

' --- 2. AUTO CARD NO ---
Private Sub Auto_CardNo()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT MAX(c_no) FROM DONAR_CARD"
    Set rsMax = C.Execute(sql)

    If IsNull(rsMax.Fields(0).Value) Then
        Label6.Caption = "1001"
    Else
        Label6.Caption = rsMax.Fields(0).Value + 1
    End If

    rsMax.Close
End Sub

' --- 3. REFRESH GRID ---
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsCommand3.State = adStateOpen Then
        DataEnvironment1.rsCommand3.Close
    End If

    DataEnvironment1.Commands("Command3").CommandText = _
        "SELECT * FROM DONAR_CARD WHERE (status = 'Active' OR status IS NULL) ORDER BY c_no ASC"

    DataEnvironment1.Command3

    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Command3"
    DataGrid1.Refresh
End Sub

' --- 4. SAVE BUTTON ---
Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If Label6.Caption = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Missing Card Number"
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MsgBox "Please select Donor ID.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If DTPicker1.Value > Date Then
        MsgBox "Issue Date cannot be greater than the current date.", vbExclamation, "Date Validation"
        DTPicker1.Value = Date
        DTPicker1.SetFocus
        Exit Sub
    End If

    If DTPicker2.Value < Date Then
        MsgBox "Valid Until cannot be less than the current date.", vbExclamation, "Date Validation"
        DTPicker2.SetFocus
        Exit Sub
    End If

    If DTPicker2.Value < DTPicker1.Value Then
        MsgBox "Valid Until cannot be earlier than Issue Date.", vbExclamation, "Date Validation"
        DTPicker2.SetFocus
        Exit Sub
    End If

    If C.State = 0 Then Call CONN

    sql = "INSERT INTO DONAR_CARD (c_no, d_id, iss_dt, valid, status) " & _
          "VALUES (" & _
          "" & Val(Label6.Caption) & ", " & _
          "'" & Combo1.Text & "', " & _
          "TO_DATE('" & Format(DTPicker1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "TO_DATE('" & Format(DTPicker2.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "'Active')"

    C.Execute sql
    MsgBox "Card issued successfully.", vbInformation, "Save Success"

    Call Clear_form
    Call RefreshGrid
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save Error: " & Err.Description, vbCritical, "System Error"
End Sub

' --- 5. SOFT DELETE ---
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Label6.Caption = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If MsgBox("Do you want to deactivate this card?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN
        sql = "UPDATE DONAR_CARD SET status = 'Inactive' WHERE c_no = " & Val(Label6.Caption)
        C.Execute sql

        MsgBox "Card deactivated successfully.", vbInformation, "Delete Success"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "System Error"
End Sub

' --- 6. SEARCH LOGIC ---
Private Sub Combo2_Click()
    If Combo2.Text = "" Then Exit Sub

    Combo3.Clear
    Combo3.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    sql = ""
    If Combo2.Text = "Card No" Then
        sql = "SELECT c_no FROM DONAR_CARD WHERE status='Active' ORDER BY c_no ASC"
    ElseIf Combo2.Text = "Donor ID" Then
        sql = "SELECT DISTINCT d_id FROM DONAR_CARD WHERE status='Active' ORDER BY d_id ASC"
    End If

    If sql = "" Then Exit Sub

    rsS.Open sql, C

    Do While Not rsS.EOF
        Combo3.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
    Set rsS = Nothing
End Sub

Private Sub Combo3_Click()
    If Trim(Combo3.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    If Combo2.Text = "" Or Combo3.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Validation"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo2.Text = "Card No" Then
        sql = "SELECT * FROM DONAR_CARD WHERE c_no=" & Val(Combo3.Text) & " AND status='Active'"
    ElseIf Combo2.Text = "Donor ID" Then
        sql = "SELECT * FROM DONAR_CARD WHERE d_id='" & Combo3.Text & "' AND status='Active'"
    End If

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label6.Caption = rsSearch!c_no

        Combo1.Clear
        Combo1.AddItem rsSearch!D_ID
        Combo1.Text = rsSearch!D_ID

        If Not IsNull(rsSearch!iss_dt) Then DTPicker1.Value = rsSearch!iss_dt
        If Not IsNull(rsSearch!valid) Then DTPicker2.Value = rsSearch!valid

        MsgBox "Card found successfully.", vbInformation, "Search Success"
        Call SetEditMode
    Else
        MsgBox "Card not found or inactive.", vbCritical, "Search Result"
    End If

    rsSearch.Close
End Sub

' --- 7. HELPER FUNCTIONS ---
Public Sub Clear_form()
    Label6.Caption = ""
    Combo1.Clear
    Combo1.Text = ""
    DTPicker1.Value = Date
    DTPicker2.Value = DateAdd("yyyy", 1, Date)
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    Combo3.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_CardNo
    Call Fill_DonorIDs

    DTPicker1.Value = Date
    DTPicker2.Value = DateAdd("yyyy", 1, Date)

    Call SetAddMode
    Combo1.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

Private Sub DTPicker1_Change()
    If DTPicker1.Value > Date Then
        MsgBox "Issue Date cannot be greater than the current date.", vbExclamation, "Date Validation"
        DTPicker1.Value = Date
    End If

    DTPicker2.Value = DateAdd("yyyy", 1, DTPicker1.Value)
End Sub

Private Sub DTPicker2_Change()
    If DTPicker2.Value < Date Then
        MsgBox "Valid Until cannot be less than the current date.", vbExclamation, "Date Validation"
        DTPicker2.Value = DateAdd("yyyy", 1, DTPicker1.Value)
    End If

    If DTPicker2.Value < DTPicker1.Value Then
        MsgBox "Valid Until cannot be earlier than Issue Date.", vbExclamation, "Date Validation"
        DTPicker2.Value = DateAdd("yyyy", 1, DTPicker1.Value)
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

