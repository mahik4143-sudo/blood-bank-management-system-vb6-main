VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmDashboard 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   3255
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3255
      Left            =   600
      OleObjectBlob   =   "frmDashboard.frx":0000
      TabIndex        =   0
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Live Blood Stock Inventory"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmDashboard (LIVE ANALYTICS CHART)
' Developer: PRJ2531G
' Project: Red Cross Blood Bank Management System
' =========================================================

Dim sql As String
Dim rsChart As New ADODB.Recordset

' =========================================================
' FORM LOAD & BUTTON EVENTS
' =========================================================
Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = "Red Cross Blood Bank - Live Analytics Dashboard"
    
    ' Form khulte hi chart load ho jayega
    Call LoadChartData
End Sub

Private Sub cmdRefresh_Click()
    ' Refresh dabane par latest stock aayega
    Call LoadChartData
    MsgBox "Dashboard Updated with Live Database Stock.", vbInformation, "Red Cross Blood Bank"
End Sub

' Agar aapne Close button (cmdClose) banaya hai toh ye use karein
Private Sub cmdClose_Click()
    Unload Me
End Sub

' =========================================================
' ?? GRAPH BHARNE KA MAIN LOGIC (WITH PERFECT JOINS) ??
' =========================================================
Private Sub LoadChartData()
    On Error GoTo ChartErr
    
    ' 1. Database connection check
    If C.State = 0 Then Call CONN
    
    ' 2. ?? 6-Table Master Join Query ??
    ' Ye aapke schema ke hisaab se Blood Stock se lekar Donor ke Blood Group tak jayegi
    sql = "SELECT D.bl_grp, COUNT(S.st_id) as TotalBags " & _
          "FROM BLOOD_STOCK S, COMPONENT_PROCESSING P, LAB_TESTING L, SAMPLE SM, DONATION_VISIT V, DONAR_DETAIL D " & _
          "WHERE S.p_id = P.p_id " & _
          "AND P.t_id = L.t_id " & _
          "AND L.s_id = SM.s_id " & _
          "AND SM.v_id = V.v_id " & _
          "AND V.d_id = D.d_id " & _
          "AND S.avail = 'AVAILABLE' " & _
          "GROUP BY D.bl_grp"
          
    ' 3. Recordset open karna
    If rsChart.State = 1 Then rsChart.Close
    rsChart.CursorLocation = adUseClient
    rsChart.Open sql, C, adOpenStatic, adLockReadOnly
    
    ' 4. Agar stock khali hai toh message dikhana
    If rsChart.RecordCount = 0 Then
        MsgBox "No AVAILABLE blood stock to display right now.", vbExclamation, "Red Cross Blood Bank"
        Exit Sub
    End If
    
    ' 5. Chart ki basic setting (Bar Chart setup)
    MSChart1.chartType = VtChChartType2dBar
    MSChart1.ColumnCount = 1
    MSChart1.RowCount = rsChart.RecordCount
    
    Dim row As Integer
    row = 1
    
    ' 6. Data ko graph mein dalna (Loop)
    Do While Not rsChart.EOF
        MSChart1.row = row
        
        ' X-Axis (Neeche ki line) - Blood Group ka naam
        MSChart1.RowLabel = rsChart.Fields(0).Value & ""
        
        ' Y-Axis (Lambi line) - Total Bags ki ginti
        MSChart1.Data = rsChart.Fields(1).Value
        
        row = row + 1
        rsChart.MoveNext
    Loop
    
    ' 7. Chart ko final touch dena
    MSChart1.ShowLegend = True
    MSChart1.Refresh
    Exit Sub
    
ChartErr:
    MsgBox "Chart Error: " & Err.Description, vbCritical, "Red Cross Blood Bank - System Error"
End Sub

