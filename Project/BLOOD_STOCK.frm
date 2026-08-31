VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBloodstock 
   Caption         =   "BLOOD STOCK"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   18240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   18240
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   12600
      TabIndex        =   11
      Top             =   1680
      Width           =   5535
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
         Left            =   2400
         TabIndex        =   13
         Text            =   "Combo3"
         Top             =   720
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
         Left            =   2400
         TabIndex        =   12
         Text            =   "Combo4"
         Top             =   1560
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
         TabIndex        =   15
         Top             =   720
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1800
         Picture         =   "BLOOD_STOCK.frx":0000
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3000
      TabIndex        =   8
      Top             =   5880
      Width           =   13095
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   11160
         Picture         =   "BLOOD_STOCK.frx":3CDF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   9000
         Picture         =   "BLOOD_STOCK.frx":F28C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   6720
         Picture         =   "BLOOD_STOCK.frx":12E35
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   4560
         Picture         =   "BLOOD_STOCK.frx":1660C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   2400
         Picture         =   "BLOOD_STOCK.frx":19C83
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   240
         Picture         =   "BLOOD_STOCK.frx":1D6D3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   11175
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
         Left            =   4560
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   2880
         Width           =   3735
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
         Left            =   4560
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   1080
         Width           =   3735
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
         Left            =   4560
         TabIndex        =   6
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label5 
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
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Availability "
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
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Rack No "
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Processing ID"
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
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Stock ID "
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
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BLOOD_STOCK.frx":20B9B
      Height          =   2655
      Left            =   840
      TabIndex        =   10
      Top             =   7080
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   4683
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
      DataMember      =   "Blood_Stock"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ST_ID"
         Caption         =   "STOCK ID"
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
         DataField       =   "P_ID"
         Caption         =   "PROCESSING ID"
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
         DataField       =   "RACK"
         Caption         =   "RACK NO."
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
         DataField       =   "AVAIL"
         Caption         =   "AVAILABLITY"
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
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "BLOOD_STOCK.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmBloodstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmBloodstock
' FINAL CODE WITH BUTTON RESTRICTIONS + VALIDATIONS
' =========================================================

Option Explicit

Dim sql As String
Dim rs As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

Private Sub cmdreport_Click()
    Blood_Stock_Report.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call CONN

    ' 1. Availability Dropdown
    Combo2.Clear
    Combo2.AddItem "AVAILABLE"
    Combo2.AddItem "ISSUED"
    Combo2.AddItem "EXPIRED"
    ' DISCARDED removed because table check constraint does not allow it

    ' 2. Search Combos
    Combo3.Clear
    Combo3.AddItem "Stock ID"
    Combo3.AddItem "Processing ID"
    Combo3.AddItem "Rack No"
    Combo4.Enabled = False

    ' 3. Processing IDs Load
    Call Fill_ProcessingIDs

    ' 4. Initial Setup
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

    Combo3.SetFocus
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
    Combo2.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo3.Enabled = bFlag
    Combo4.Enabled = bFlag
End Sub

' =========================================================
' HELPERS
' =========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function IsRackValid(ByVal s As String) As Boolean
    Dim i As Integer
    Dim ch As String

    s = Trim(s)

    If s = "" Then
        IsRackValid = False
        Exit Function
    End If

    If Len(s) > 10 Then
        IsRackValid = False
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "-" Or ch = "/" Then
            ' allowed
        Else
            IsRackValid = False
            Exit Function
        End If
    Next i

    IsRackValid = True
End Function

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(Label5.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Action Required"
        Exit Function
    End If

    If Trim(Combo1.Text) = "" Then
        MsgBox "Please select Processing ID.", vbExclamation, "Validation Error"
        Combo1.SetFocus
        Exit Function
    End If

    If Trim(Text1.Text) = "" Then
        MsgBox "Please enter Rack No.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If IsRackValid(Text1.Text) = False Then
        MsgBox "Rack No must be up to 10 characters and can contain letters, digits, '-' or '/'.", vbExclamation, "Validation Error"
        Text1.SetFocus
        Exit Function
    End If

    If Trim(Combo2.Text) = "" Then
        MsgBox "Please select Availability.", vbExclamation, "Validation Error"
        Combo2.SetFocus
        Exit Function
    End If

    If UCase(Trim(Combo2.Text)) <> "AVAILABLE" And _
       UCase(Trim(Combo2.Text)) <> "ISSUED" And _
       UCase(Trim(Combo2.Text)) <> "EXPIRED" Then
        MsgBox "Availability must be AVAILABLE, ISSUED, or EXPIRED.", vbExclamation, "Validation Error"
        Combo2.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

' =========================================================
' SMART PROCESSING ID LOADER
' =========================================================
Private Sub Fill_ProcessingIDs()
    Dim rsProc As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    Combo1.Clear

    sql = "SELECT p_id FROM COMPONENT_PROCESSING WHERE status='Active' AND " & _
          "p_id NOT IN (SELECT p_id FROM BLOOD_STOCK WHERE status='Active') " & _
          "ORDER BY p_id ASC"

    rsProc.Open sql, C
    Do While Not rsProc.EOF
        Combo1.AddItem rsProc!p_id
        rsProc.MoveNext
    Loop
    rsProc.Close
End Sub

' =========================================================
' AUTO STOCK ID
' =========================================================
Private Sub Auto_StockID()
    Dim rsMax As New ADODB.Recordset

    If C.State = 0 Then Call CONN

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(st_id, 4))), 0) FROM BLOOD_STOCK"
    Set rsMax = C.Execute(sql)

    Label5.Caption = "STK" & Format(rsMax.Fields(0).Value + 1, "0000")

    rsMax.Close
End Sub

' =========================================================
' ADD NEW
' =========================================================
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_StockID
    Call Fill_ProcessingIDs
    Call SetAddMode
    Combo1.SetFocus
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

    sql = "INSERT INTO BLOOD_STOCK (st_id, p_id, rack, avail, status) " & _
          "VALUES (" & _
          "'" & Esc(Label5.Caption) & "', " & _
          "'" & Esc(Combo1.Text) & "', " & _
          "'" & Esc(Text1.Text) & "', " & _
          "'" & Esc(UCase(Combo2.Text)) & "', " & _
          "'Active')"

    C.Execute sql
    MsgBox "Stock added successfully.", vbInformation, "Success"

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

    If Trim(Label5.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If ValidateForm() = False Then Exit Sub

    If MsgBox("Do you want to update this stock record?", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        Call CONN

        sql = "UPDATE BLOOD_STOCK SET " & _
              "rack='" & Esc(Text1.Text) & "', " & _
              "avail='" & Esc(UCase(Combo2.Text)) & "' " & _
              "WHERE st_id='" & Esc(Label5.Caption) & "'"

        C.Execute sql

        MsgBox "Stock record updated successfully.", vbInformation, "Success"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

UpErr:
    MsgBox "Update Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' DELETE (SOFT DELETE)
' =========================================================
Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(Label5.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If UCase(Trim(Combo2.Text)) = "ISSUED" Then
        MsgBox "Cannot delete. This stock is marked as ISSUED.", vbCritical, "Restricted Action"
        Exit Sub
    End If

    If MsgBox("Do you want to remove this stock record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        Call CONN

        sql = "UPDATE BLOOD_STOCK SET status='Inactive' WHERE st_id='" & Esc(Label5.Caption) & "'"
        C.Execute sql

        MsgBox "Stock removed successfully.", vbInformation, "Deleted"
        Call RefreshGrid
        Call Clear_form
        Call SetStartupMode
    End If
    Exit Sub

DelErr:
    MsgBox "Delete Error: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================
' SEARCH LOAD
' =========================================================
Private Sub Combo3_Click()
    If Combo3.Text = "" Then Exit Sub

    Combo4.Clear
    Combo4.Enabled = True
    cmdSearch.Enabled = False

    Dim rsS As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    If Combo3.Text = "Stock ID" Then sql = "SELECT st_id FROM BLOOD_STOCK WHERE status='Active' ORDER BY st_id ASC"
    If Combo3.Text = "Processing ID" Then sql = "SELECT DISTINCT p_id FROM BLOOD_STOCK WHERE status='Active' ORDER BY p_id ASC"
    If Combo3.Text = "Rack No" Then sql = "SELECT DISTINCT rack FROM BLOOD_STOCK WHERE status='Active' ORDER BY rack ASC"

    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo4.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Combo4_Click()
    If Trim(Combo4.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

' =========================================================
' SEARCH
' =========================================================
Private Sub cmdSearch_Click()
    If Combo3.Text = "" Or Combo4.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Validation"
        Exit Sub
    End If

    Dim rsSearch As New ADODB.Recordset

    If Combo3.Text = "Stock ID" Then sql = "SELECT * FROM BLOOD_STOCK WHERE st_id='" & Esc(Combo4.Text) & "' AND status='Active'"
    If Combo3.Text = "Processing ID" Then sql = "SELECT * FROM BLOOD_STOCK WHERE p_id='" & Esc(Combo4.Text) & "' AND status='Active'"
    If Combo3.Text = "Rack No" Then sql = "SELECT * FROM BLOOD_STOCK WHERE rack='" & Esc(Combo4.Text) & "' AND status='Active'"

    If C.State = 0 Then Call CONN
    rsSearch.Open sql, C

    If Not rsSearch.EOF Then
        Label5.Caption = rsSearch!st_id

        Combo1.Clear
        Combo1.AddItem rsSearch!p_id
        Combo1.Text = rsSearch!p_id

        Text1.Text = rsSearch!rack
        Combo2.Text = rsSearch!avail

        MsgBox "Record found successfully.", vbInformation, "Success"
        Call SetEditMode
    Else
        MsgBox "Record not found or inactive.", vbCritical, "Error"
    End If

    rsSearch.Close
End Sub

' =========================================================
' REFRESH GRID
' =========================================================
Private Sub RefreshGrid()
    On Error Resume Next

    If DataEnvironment1.rsBlood_Stock.State = adStateOpen Then
        DataEnvironment1.rsBlood_Stock.Close
    End If

    DataEnvironment1.Commands("Blood_Stock").CommandText = _
        "SELECT * FROM BLOOD_STOCK " & _
        "WHERE (status='Active' OR status IS NULL) " & _
        "ORDER BY TO_NUMBER(SUBSTR(st_id,4)) ASC"

    DataEnvironment1.Blood_Stock
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Blood_Stock"
    DataGrid1.Refresh
End Sub

' =========================================================
' CLEAR
' =========================================================
Public Sub Clear_form()
    Label5.Caption = ""
    Combo1.Clear
    Combo1.Text = ""
    Text1.Text = ""
    Combo2.ListIndex = -1
    Combo2.Text = ""
    Combo3.ListIndex = -1
    Combo3.Text = ""
    Combo4.Clear
    Combo4.Text = ""
    Combo4.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    Call SetStartupMode
End Sub

' =========================================================
' KEY NAVIGATION / RESTRICTIONS
' =========================================================
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Combo2.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub

    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       (KeyAscii >= 48 And KeyAscii <= 57) Or _
       KeyAscii = 45 Or KeyAscii = 47 Then
        If Len(Text1.Text) >= 10 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If CmdSave.Enabled = True Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled = True Then
            'CmdUpdate.SetFocus
        End If
    Else
        KeyAscii = 0
    End If
End Sub

