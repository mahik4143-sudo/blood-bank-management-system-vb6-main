VERSION 5.00
Begin VB.Form Ref_blood_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REFERENCE BLOOD RANGE REPORT"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   2775
      Left            =   1320
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
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "Ref_blood_Report.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Ref_blood_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: REFERENCE BLOOD RANGE REPORT (Filter Form)
' =========================================================

Private Sub Form_Load()
    ' Search options add karna
    Combo1.Clear
    Combo1.AddItem "Reference ID"
    Combo1.AddItem "Component Name"
    
    ' VALIDATION 1: Combo2 ko disable rakhna
    Combo2.Text = ""
    Combo2.Clear
    Combo2.Enabled = False
    
    ' VALIDATION 2: PRINT (Selective) Button ko disable rakhna
    Command1.Enabled = False
End Sub

' =========================================================
' COMBO1 SELECTION (Enable Combo2)
' =========================================================
Private Sub Combo1_Click()
    ' Naya option select karte hi Combo2 ko clear aur enable karna
    Combo2.Clear
    Combo2.Enabled = True
    
    ' Print button ko wapas disable karna
    Command1.Enabled = False
    
    Dim rsS As New ADODB.Recordset
    Dim sql As String
    
    If C.State = 0 Then Call CONN
    
    ' Database se IDs laana
    If Combo1.Text = "Reference ID" Then
        sql = "SELECT ref_id FROM REF_BLOOD_RANGE ORDER BY ref_id ASC"
    ElseIf Combo1.Text = "Component Name" Then
        sql = "SELECT DISTINCT cm_nm FROM REF_BLOOD_RANGE ORDER BY cm_nm ASC"
    End If
    
    ' Combo2 ko bharna
    rsS.Open sql, C
    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
End Sub

' =========================================================
' COMBO2 SELECTION (Enable PRINT Button)
' =========================================================
Private Sub Combo2_Click()
    ' Jaise hi Combo2 mein kuch select hoga, PRINT button ON ho jayega
    If Combo2.Text <> "" Then
        Command1.Enabled = True
    End If
End Sub

' =========================================================
' PRINT BUTTON (Selective Print - Command1)
' =========================================================
Private Sub Command1_Click()
    Dim sql As String
    
    ' Filter Query
    If Combo1.Text = "Reference ID" Then
        sql = "SELECT * FROM REF_BLOOD_RANGE WHERE ref_id='" & Combo2.Text & "'"
    ElseIf Combo1.Text = "Component Name" Then
        sql = "SELECT * FROM REF_BLOOD_RANGE WHERE cm_nm='" & Combo2.Text & "'"
    End If

    ' Error Prevention (Command name: Ref_Blood_Range)
    On Error Resume Next
    If DataEnvironment1.rsRef_Blood_Range.State = 1 Then
        DataEnvironment1.rsRef_Blood_Range.Close
    End If
    On Error GoTo 0

    ' Query Execute karna
    DataEnvironment1.Commands("Ref_Blood_Range").CommandText = sql
    DataEnvironment1.Commands("Ref_Blood_Range").Execute
    
    ' Report Show karna (Aapki teesri image ke hisaab se Report ka naam)
    REF_BLOOD_RANGE_REPORT.Show
End Sub

' =========================================================
' PRINT ALL BUTTON (Saare Records - Command2)
' =========================================================
Private Sub Command2_Click()
    Dim sql As String
    
    ' Bina condition ki query
    sql = "SELECT * FROM REF_BLOOD_RANGE ORDER BY ref_id ASC"
    
    ' Error Prevention
    On Error Resume Next
    If DataEnvironment1.rsRef_Blood_Range.State = 1 Then
        DataEnvironment1.rsRef_Blood_Range.Close
    End If
    On Error GoTo 0

    ' Query Execute karna
    DataEnvironment1.Commands("Ref_Blood_Range").CommandText = sql
    DataEnvironment1.Commands("Ref_Blood_Range").Execute
    
    ' Report Show karna
    Reference_Blood_Report1.Show
End Sub
