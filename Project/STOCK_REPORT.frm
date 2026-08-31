VERSION 5.00
Begin VB.Form STOCK_REPORT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock_Report"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PRINT ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   2
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   3240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "STOCK_REPORT.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "STOCK_REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataEnvironment1.Command1 Combo1.Text
STOCK_REPOTY.Show

Set DataEnvironment1 = Nothing
End Sub


Private Sub Form_Load()
Call CONN
Call Fill_StockCombo
End Sub
Private Sub Fill_StockCombo()
    On Error GoTo errh
    Dim rsDonor As New ADODB.Recordset
    
    ' Connection check karein
    If C Is Nothing Then Call CONN
    
    ' ComboBox ko clear karein (Name: cmbDonarID)
    Combo1.Clear
    
    ' Database se IDs uthayein
    Set rsDonor = C.Execute("SELECT ST_ID FROM BLOOD_STOCK ORDER BY ST_ID ASC")
    
    ' Loop se IDs Combo mein bharein
    Do While Not rsstock.EOF
        Combo1.AddItem stock_id
        rsstock.MoveNext
    Loop
    
    rsDonor.Close
    Set rsDonor = Nothing
    Exit Sub
errh:
    MsgBox "ID Load Error: " & Err.Description, vbCritical
End Sub
