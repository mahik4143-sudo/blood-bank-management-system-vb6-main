VERSION 5.00
Begin VB.Form Consent_Report 
   Caption         =   "Consent Report"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   720
      TabIndex        =   0
      Top             =   240
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
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "PRINT ALL"
         Height          =   735
         Left            =   2400
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
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
         TabIndex        =   6
         Top             =   480
         Width           =   1695
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Consent_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: Consent_Report (Smart Filter Form)
' =========================================================

Private Sub Form_Load()
    ' Initialize Combo1 with search options
    Combo1.Clear
    Combo1.AddItem "Consent ID"
    Combo1.AddItem "Visit ID"
    
    ' Disable Combo2 until a selection is made in Combo1
    Combo2.Enabled = False
End Sub

Private Sub Combo1_Click()
    Combo2.Clear
    Combo2.Enabled = True
    Dim rsS As New ADODB.Recordset
    
    ' Check and open database connection if closed
    If C.State = 0 Then Call CONN
    
    ' Fetch IDs from the CONSENT table based on selection
    If Combo1.Text = "Consent ID" Then
        sql = "SELECT co_id FROM CONSENT ORDER BY co_id ASC"
    ElseIf Combo1.Text = "Visit ID" Then
        sql = "SELECT DISTINCT v_id FROM CONSENT ORDER BY v_id ASC"
    End If
    
    ' Execute query and populate Combo2
    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

Private Sub Command1_Click()
    Dim sql As String
    
    ' Validation: Ensure a value is selected before printing
    If Combo2.Text = "" Then
        MsgBox "Please select an ID to print!", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If

    ' Build the filter query based on the selected ID
    If Combo1.Text = "Consent ID" Then
        sql = "SELECT * FROM CONSENT WHERE co_id='" & Combo2.Text & "'"
    ElseIf Combo1.Text = "Visit ID" Then
        sql = "SELECT * FROM CONSENT WHERE v_id='" & Combo2.Text & "'"
    End If

    ' =========================================================
    ' ?? PRO-TIP (Error Prevention for DataEnvironment) ??
    ' In VB6 DataEnvironment, you must close the previous Recordset
    ' before executing a new Command, otherwise it throws an
    ' "Operation is not allowed when the object is open" error.
    ' =========================================================
    
    ' NOTE: Match "rsConsent_Report" with your DataEnvironment Command name.
    ' Just add "rs" in front of your command name.
    ' (e.g., if your command is named "Consent", this should be "rsConsent").
    On Error Resume Next
    If DataEnvironment1.rsConsent_Report.State = 1 Then
        DataEnvironment1.rsConsent_Report.Close
    End If
    On Error GoTo 0

    ' Send the new SQL query to the DataEnvironment Command
    ' NOTE: Replace "Consent_Report" with the actual Command name inside DataEnvironment1
    DataEnvironment1.Commands("Consent_Report").CommandText = sql
    DataEnvironment1.Commands("Consent_Report").Execute
    
    ' Show the generated report
    Consent_Report1.Show
End Sub
