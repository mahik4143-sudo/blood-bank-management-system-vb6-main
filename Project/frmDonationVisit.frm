VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDonationVisit 
   Caption         =   " Donation Visit Form"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   18975
   Begin VB.Frame Frame4 
      Height          =   3855
      Left            =   14040
      TabIndex        =   30
      Top             =   1680
      Width           =   4335
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
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   720
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
         Left            =   2040
         TabIndex        =   31
         Text            =   "Combo2"
         Top             =   1800
         Width           =   2055
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
         TabIndex        =   34
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
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Image cmdSearch 
         Height          =   735
         Left            =   1080
         Picture         =   "frmDonationVisit.frx":0000
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDonationVisit.frx":3CDF
      Height          =   1455
      Left            =   720
      TabIndex        =   29
      Top             =   8520
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Donation_Visit"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "V_ID"
         Caption         =   "VISIT ID"
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
         DataField       =   "DON_DT"
         Caption         =   "DATE"
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
         DataField       =   "REF_NM"
         Caption         =   "REFRENCE NAME"
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
         DataField       =   "OPD"
         Caption         =   "OPD"
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
         DataField       =   "BED"
         Caption         =   "BED"
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
         DataField       =   "WGT"
         Caption         =   "WEIGHT"
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
         DataField       =   "BP"
         Caption         =   "BP"
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
         DataField       =   "PULSE"
         Caption         =   "PULSE"
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
         DataField       =   "TEM"
         Caption         =   "TEMP"
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
      BeginProperty Column10 
         DataField       =   "HB_LEV"
         Caption         =   "HB LEVEL"
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
      BeginProperty Column11 
         DataField       =   "FIN_STAT"
         Caption         =   "FINAL STATUS"
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
      BeginProperty Column12 
         DataField       =   "DEFER"
         Caption         =   "DEFER"
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
      BeginProperty Column13 
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker TO_DATE 
      Height          =   495
      Left            =   3240
      TabIndex        =   28
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
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
      Format          =   173998081
      CurrentDate     =   46052
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   720
      TabIndex        =   27
      Top             =   7320
      Width           =   17655
      Begin VB.Image CmdAddNew 
         Height          =   615
         Left            =   120
         Picture         =   "frmDonationVisit.frx":3CFE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image CmdUpdate 
         Height          =   615
         Left            =   3000
         Picture         =   "frmDonationVisit.frx":71C6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdSave 
         Height          =   615
         Left            =   6120
         Picture         =   "frmDonationVisit.frx":AC16
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdClear 
         Height          =   615
         Left            =   9120
         Picture         =   "frmDonationVisit.frx":E28D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image CmdDelete 
         Height          =   615
         Left            =   12240
         Picture         =   "frmDonationVisit.frx":11A64
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image cmdreport 
         Height          =   615
         Left            =   15240
         Picture         =   "frmDonationVisit.frx":1560D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.ComboBox cmbDonarID 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   26
      Text            =   "cmbDonarID"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ComboBox cmbFinalStatus 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9960
      TabIndex        =   19
      Text            =   "cmbFinalStatus"
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox txtBED 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      TabIndex        =   18
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Medical Status"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   5295
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
      Begin VB.TextBox txtDeferReason 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   4560
         Width           =   3375
      End
      Begin VB.TextBox txtHB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         TabIndex        =   24
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtTemp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         TabIndex        =   23
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtPulse 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
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
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtBP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   21
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtWeight 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "Defer Reason"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Final status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "HB Level"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Pulse"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "BP "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Weight (Kg) "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame fraVisit 
      Caption         =   "Visit Patient details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   5295
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   6495
      Begin VB.TextBox txtOPD 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2520
         TabIndex        =   17
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox txtRefPatient 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2520
         TabIndex        =   16
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblVisitID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   525
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "BED No"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "OPD No"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ref Patient Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Donation Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Donar ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Visit ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmDonationVisit.frx":20BBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
End
Attribute VB_Name = "frmDonationVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMode As String   ' STARTUP / ADD / EDIT
Private sql As String

'========================================================
' FORM LOAD
'========================================================
Private Sub Form_Load()
    On Error GoTo ErrHandler

    CONN

    LoadFinalStatus
    InitializeSearchSection
    ClearFormValues
    Fill_DonorCombo
    RefreshGrid
    SetStartupMode

    MsgBox "Donation Visit Form loaded successfully.", vbInformation, "Form Ready"
    Exit Sub

ErrHandler:
'    MsgBox "Form Load Error : " & Err.Description, vbCritical, "Error"
End Sub

'========================================================
' MODE CONTROL
'========================================================
Private Sub SetStartupMode()
    mMode = "STARTUP"

    EnableEntryControls False
    EnableSearchControls True

    CmdAddNew.Enabled = True
    cmdSearch.Enabled = False
    cmdreport.Enabled = True

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = False

    Combo1.SetFocus
End Sub

Private Sub SetAddMode()
    mMode = "ADD"

    EnableEntryControls True
    EnableSearchControls False

    CmdAddNew.Enabled = False
    cmdSearch.Enabled = False
    cmdreport.Enabled = False

    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = True

    cmbDonarID.SetFocus
End Sub

Private Sub SetEditMode()
    mMode = "EDIT"

    EnableEntryControls True
    EnableSearchControls True

    CmdAddNew.Enabled = True
    cmdSearch.Enabled = True
    cmdreport.Enabled = True

    CmdSave.Enabled = False
    CmdUpdate.Enabled = True
    CmdDelete.Enabled = True
    CmdClear.Enabled = True

    cmbDonarID.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    cmbDonarID.Enabled = bFlag
    TO_DATE.Enabled = bFlag
    txtRefPatient.Enabled = bFlag
    txtOPD.Enabled = bFlag
    txtBED.Enabled = bFlag
    txtWeight.Enabled = bFlag
    txtBP.Enabled = bFlag
    txtPulse.Enabled = bFlag
    txtTemp.Enabled = bFlag
    txtHB.Enabled = bFlag
    cmbFinalStatus.Enabled = bFlag

    If bFlag = False Then
        txtDeferReason.Enabled = False
    Else
        ToggleDeferReason
    End If
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag
    Combo2.Enabled = bFlag
End Sub

'========================================================
' INITIALIZATION
'========================================================
Private Sub LoadFinalStatus()
    cmbFinalStatus.Clear
    cmbFinalStatus.AddItem "FIT"
    cmbFinalStatus.AddItem "UNFIT"
    cmbFinalStatus.AddItem "DEFERRED"
End Sub

Private Sub InitializeSearchSection()
    Combo1.Clear
    Combo2.Clear

    Combo1.AddItem "Visit ID"
    Combo1.AddItem "Donar ID"
    Combo1.AddItem "Patient Name"

    Combo1.ListIndex = -1
    Combo2.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub ClearFormValues()
    lblVisitID.Caption = ""

    cmbDonarID.ListIndex = -1
    cmbDonarID.Text = ""
    TO_DATE.Value = Date
    txtRefPatient.Text = ""
    txtOPD.Text = ""
    txtBED.Text = ""
    txtWeight.Text = ""
    txtBP.Text = ""
    txtPulse.Text = ""
    txtTemp.Text = ""
    txtHB.Text = ""
    cmbFinalStatus.ListIndex = -1
    cmbFinalStatus.Text = ""
    txtDeferReason.Text = ""
    txtDeferReason.Enabled = False

    Combo1.ListIndex = -1
    Combo2.Clear
    Combo2.Text = ""
    Combo2.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub ResetForm()
    LoadFinalStatus
    InitializeSearchSection
    ClearFormValues
    Fill_DonorCombo
    RefreshGrid
    SetStartupMode

    MsgBox "Form reset successfully.", vbInformation, "Reset"
End Sub

'========================================================
' HELPERS
'========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsNull(v) Then
        NzText = ""
    Else
        NzText = Trim(CStr(v))
    End If
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
    Dim ch As String
    Dim dotCount As Integer

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

    If Not IsDigitsOnly(arr(0)) Then
        IsValidBP = False
        Exit Function
    End If

    If Not IsDigitsOnly(arr(1)) Then
        IsValidBP = False
        Exit Function
    End If

    IsValidBP = True
End Function

Private Function ProperCaseText(ByVal s As String) As String
    s = Trim(s)

    If s = "" Then
        ProperCaseText = ""
    Else
        ProperCaseText = UCase(Left$(s, 1)) & LCase(Mid$(s, 2))
    End If
End Function

Private Sub ToggleDeferReason()
    If UCase(Trim(cmbFinalStatus.Text)) = "UNFIT" Or UCase(Trim(cmbFinalStatus.Text)) = "DEFERRED" Then
        txtDeferReason.Enabled = True
    Else
        txtDeferReason.Enabled = False
        txtDeferReason.Text = ""
    End If
End Sub

'========================================================
' VISIT ID
'========================================================
Private Sub Auto_VisitID()
    On Error GoTo ErrHandler

    Dim rsMax As ADODB.Recordset
    Dim nVal As Long

    CONN

    Set rsMax = New ADODB.Recordset
    rsMax.Open "SELECT NVL(MAX(TO_NUMBER(SUBSTR(V_ID,4))),0)+1 AS NEXT_ID FROM DONATION_VISIT WHERE V_ID LIKE 'VIS%'", C, adOpenForwardOnly, adLockReadOnly

    If Not rsMax.EOF Then
        nVal = CLng(rsMax!NEXT_ID)
    Else
        nVal = 1
    End If

    lblVisitID.Caption = "VIS" & Format(nVal, "0000")

    rsMax.Close
    Set rsMax = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error generating Visit ID : " & Err.Description, vbCritical, "Visit ID Error"
End Sub

'========================================================
' DONOR COMBO
'========================================================
Private Sub Fill_DonorCombo()
    On Error GoTo ErrHandler

    Dim rsD As ADODB.Recordset

    CONN

    cmbDonarID.Clear
    Set rsD = New ADODB.Recordset

    sql = "SELECT D_ID FROM DONAR_DETAIL WHERE STATUS='ACTIVE' " & _
          "AND D_ID NOT IN (SELECT D_ID FROM DONATION_VISIT WHERE STATUS='ACTIVE') " & _
          "ORDER BY D_ID"

    rsD.Open sql, C, adOpenForwardOnly, adLockReadOnly

    Do While Not rsD.EOF
        cmbDonarID.AddItem rsD!D_ID
        rsD.MoveNext
    Loop

    rsD.Close
    Set rsD = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error loading donor list : " & Err.Description, vbCritical, "Donor Load Error"
End Sub

Private Sub Fill_DonorCombo_ForEdit(ByVal donorID As String)
    On Error Resume Next

    Fill_DonorCombo

    If Trim(donorID) <> "" Then
        cmbDonarID.AddItem donorID
        cmbDonarID.Text = donorID
    End If
End Sub

'========================================================
' GRID
'========================================================
'========================================================
' 1. GRID (FIXED: Ascending Order & Active Only)
'========================================================
'========================================================
' REFRESH GRID (100% WORKING - DATA ENVIRONMENT BINDING)
'========================================================
Private Sub RefreshGrid()
    On Error Resume Next
    
    ' Pehle purana connection close karein (agar open hai)
    If DataEnvironment1.rsDonation_Visit.State = 1 Then
        DataEnvironment1.rsDonation_Visit.Close
    End If
    
    ' Sahi SQL Query set karein (Sirf ACTIVE data aur Serial wise)
    DataEnvironment1.Commands("Donation_Visit").CommandText = "SELECT * FROM DONATION_VISIT WHERE UPPER(status)='ACTIVE' ORDER BY v_id ASC"
    
    ' Query ko execute karein
    DataEnvironment1.Donation_Visit
    
    ' Grid ko fresh DataEnvironment ke saath bind karein
    Set DataGrid1.DataSource = DataEnvironment1
    DataGrid1.DataMember = "Donation_Visit"
    DataGrid1.Refresh
End Sub

'========================================================
' MEDICAL VALIDATION
'========================================================
Private Function Check_Vitals_With_Reference() As Boolean
    On Error GoTo ErrHandler

    Dim rsDonor As ADODB.Recordset
    Dim rsRef As ADODB.Recordset
    Dim gender As String
    Dim failMsg As String
    Dim isFit As Boolean
    Dim comp As String
    Dim minStr As String
    Dim maxStr As String

    isFit = True
    failMsg = ""

    CONN

    Set rsDonor = New ADODB.Recordset
    sql = "SELECT GENDER FROM DONAR_DETAIL WHERE D_ID='" & Esc(cmbDonarID.Text) & "'"
    rsDonor.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If Not rsDonor.EOF Then
        gender = UCase(NzText(rsDonor!gender))
    Else
        gender = "MALE"
    End If

    rsDonor.Close
    Set rsDonor = Nothing

    Set rsRef = New ADODB.Recordset
    sql = "SELECT * FROM REF_BLOOD_RANGE WHERE STATUS='Active'"
    rsRef.Open sql, C, adOpenForwardOnly, adLockReadOnly

    Do While Not rsRef.EOF
        comp = UCase(NzText(rsRef!CM_NM))

        If gender = "MALE" Then
            minStr = NzText(rsRef!MN_V_M)
            maxStr = NzText(rsRef!MX_V_M)
        Else
            minStr = NzText(rsRef!MN_V_F)
            maxStr = NzText(rsRef!MX_V_F)
        End If

        Select Case comp
            Case "WEIGHT"
                If Val(txtWeight.Text) < Val(minStr) Or Val(txtWeight.Text) > Val(maxStr) Then
                    isFit = False
                    failMsg = failMsg & "Weight out of range (" & minStr & " - " & maxStr & " Kg)" & vbCrLf
                End If

            Case "PULSE"
                If Val(txtPulse.Text) < Val(minStr) Or Val(txtPulse.Text) > Val(maxStr) Then
                    isFit = False
                    failMsg = failMsg & "Pulse out of range (" & minStr & " - " & maxStr & " bpm)" & vbCrLf
                End If

            Case "TEMPERATURE"
                If Val(txtTemp.Text) < Val(minStr) Or Val(txtTemp.Text) > Val(maxStr) Then
                    isFit = False
                    failMsg = failMsg & "Temperature out of range (" & minStr & " - " & maxStr & ")" & vbCrLf
                End If

            Case "HEMOGLOBIN"
                If Val(txtHB.Text) < Val(minStr) Or Val(txtHB.Text) > Val(maxStr) Then
                    isFit = False
                    failMsg = failMsg & "Hemoglobin out of range (" & minStr & " - " & maxStr & " g/dL)" & vbCrLf
                End If

            Case "BLOOD PRESSURE"
                Dim bpIn() As String, minBP() As String, maxBP() As String

                If IsValidBP(txtBP.Text) Then
                    bpIn = Split(txtBP.Text, "/")
                    minBP = Split(minStr, "/")
                    maxBP = Split(maxStr, "/")

                    If UBound(bpIn) = 1 And UBound(minBP) = 1 And UBound(maxBP) = 1 Then
                        If Val(bpIn(0)) < Val(minBP(0)) Or Val(bpIn(0)) > Val(maxBP(0)) Or _
                           Val(bpIn(1)) < Val(minBP(1)) Or Val(bpIn(1)) > Val(maxBP(1)) Then
                            isFit = False
                            failMsg = failMsg & "Blood Pressure out of range (" & minStr & " to " & maxStr & ")" & vbCrLf
                        End If
                    End If
                End If
        End Select

        rsRef.MoveNext
    Loop

    rsRef.Close
    Set rsRef = Nothing

    If isFit = False Then
        cmbFinalStatus.Text = "UNFIT"
        txtDeferReason.Enabled = True
        txtDeferReason.Text = Replace(failMsg, vbCrLf, " | ")

        MsgBox "Medical validation failed." & vbCrLf & vbCrLf & failMsg, vbExclamation, "Medical Status"
        Check_Vitals_With_Reference = False
    Else
        cmbFinalStatus.Text = "FIT"
        txtDeferReason.Enabled = False
        txtDeferReason.Text = ""

        MsgBox "All vitals are within allowed range.", vbInformation, "Medical Status"
        Check_Vitals_With_Reference = True
    End If

    Exit Function

ErrHandler:
    MsgBox "Error during medical validation : " & Err.Description, vbCritical, "Validation Error"
    Check_Vitals_With_Reference = False
End Function

'========================================================
' FORM VALIDATION
'========================================================
Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(lblVisitID.Caption) = "" Then
        MsgBox "Visit ID is missing.", vbExclamation, "Validation Error"
        Exit Function
    End If

    If Trim(cmbDonarID.Text) = "" Then
        MsgBox "Please select Donar ID.", vbExclamation, "Validation Error"
        cmbDonarID.SetFocus
        Exit Function
    End If

    If Trim(txtWeight.Text) = "" Then
        MsgBox "Please enter weight.", vbExclamation, "Validation Error"
        txtWeight.SetFocus
        Exit Function
    End If

    If Not IsDecimalValue(txtWeight.Text) Then
        MsgBox "Weight must be numeric.", vbExclamation, "Validation Error"
        txtWeight.SetFocus
        Exit Function
    End If

    If Trim(txtBP.Text) = "" Then
        MsgBox "Please enter blood pressure.", vbExclamation, "Validation Error"
        txtBP.SetFocus
        Exit Function
    End If

    If Not IsValidBP(txtBP.Text) Then
        MsgBox "Blood pressure must be in format like 120/80.", vbExclamation, "Validation Error"
        txtBP.SetFocus
        Exit Function
    End If

    If Trim(txtPulse.Text) = "" Then
        MsgBox "Please enter pulse.", vbExclamation, "Validation Error"
        txtPulse.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(txtPulse.Text) Then
        MsgBox "Pulse must be numeric.", vbExclamation, "Validation Error"
        txtPulse.SetFocus
        Exit Function
    End If

    If Trim(txtTemp.Text) = "" Then
        MsgBox "Please enter temperature.", vbExclamation, "Validation Error"
        txtTemp.SetFocus
        Exit Function
    End If

    If Not IsDecimalValue(txtTemp.Text) Then
        MsgBox "Temperature must be numeric.", vbExclamation, "Validation Error"
        txtTemp.SetFocus
        Exit Function
    End If

    If Trim(txtHB.Text) = "" Then
        MsgBox "Please enter HB level.", vbExclamation, "Validation Error"
        txtHB.SetFocus
        Exit Function
    End If

    If Not IsDecimalValue(txtHB.Text) Then
        MsgBox "HB level must be numeric.", vbExclamation, "Validation Error"
        txtHB.SetFocus
        Exit Function
    End If

    If Trim(txtOPD.Text) <> "" Then
        If Not IsDigitsOnly(txtOPD.Text) Then
            MsgBox "OPD number must be numeric.", vbExclamation, "Validation Error"
            txtOPD.SetFocus
            Exit Function
        End If
    End If

    If Trim(cmbFinalStatus.Text) = "" Then
        MsgBox "Please select final status.", vbExclamation, "Validation Error"
        cmbFinalStatus.SetFocus
        Exit Function
    End If

    If UCase(Trim(cmbFinalStatus.Text)) = "UNFIT" Or UCase(Trim(cmbFinalStatus.Text)) = "DEFERRED" Then
        If Trim(txtDeferReason.Text) = "" Then
            MsgBox "Please enter defer reason.", vbExclamation, "Validation Error"
            txtDeferReason.SetFocus
            Exit Function
        End If
    End If

    ValidateForm = True
End Function

Private Function ActiveVisitExists(ByVal donorID As String, Optional ByVal excludeVisitID As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim rsChk As ADODB.Recordset

    CONN
    Set rsChk = New ADODB.Recordset

    sql = "SELECT V_ID FROM DONATION_VISIT WHERE D_ID='" & Esc(donorID) & "' AND STATUS='Active'"

    If Trim(excludeVisitID) <> "" Then
        sql = sql & " AND V_ID<>'" & Esc(excludeVisitID) & "'"
    End If

    rsChk.Open sql, C, adOpenForwardOnly, adLockReadOnly

    ActiveVisitExists = Not rsChk.EOF

    rsChk.Close
    Set rsChk = Nothing
    Exit Function

ErrHandler:
    MsgBox "Error checking duplicate active visit : " & Err.Description, vbCritical, "Duplicate Check Error"
    ActiveVisitExists = True
End Function

'========================================================
' BUTTONS
'========================================================
Private Sub cmdAddNew_Click()
    On Error GoTo ErrHandler

    ClearFormValues
    Fill_DonorCombo
    Auto_VisitID
    SetAddMode

    MsgBox "New visit entry mode started.", vbInformation, "Add New"
    Exit Sub

ErrHandler:
    MsgBox "Add New Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If ActiveVisitExists(cmbDonarID.Text) = True Then
        MsgBox "This donor already has an active visit record.", vbExclamation, "Duplicate Active Visit"
        cmbDonarID.SetFocus
        Exit Sub
    End If

    Call Check_Vitals_With_Reference

    If UCase(Trim(cmbFinalStatus.Text)) = "UNFIT" Or UCase(Trim(cmbFinalStatus.Text)) = "DEFERRED" Then
        If Trim(txtDeferReason.Text) = "" Then
            MsgBox "Defer reason is required.", vbExclamation, "Validation Error"
            txtDeferReason.SetFocus
            Exit Sub
        End If
    End If

    CONN

    sql = "INSERT INTO DONATION_VISIT (" & _
          "V_ID, D_ID, DON_DT, REF_NM, OPD, BED, WGT, BP, PULSE, TEM, HB_LEV, FIN_STAT, DEFER, STATUS" & _
          ") VALUES (" & _
          "'" & Esc(lblVisitID.Caption) & "'," & _
          "'" & Esc(cmbDonarID.Text) & "'," & _
          "TO_DATE('" & Format(TO_DATE.Value, "DD-MM-YYYY") & "','DD-MM-YYYY')," & _
          "'" & Esc(txtRefPatient.Text) & "'," & _
          IIf(Trim(txtOPD.Text) = "", "NULL", Trim(txtOPD.Text)) & "," & _
          "'" & Esc(txtBED.Text) & "'," & _
          Val(txtWeight.Text) & "," & _
          "'" & Esc(txtBP.Text) & "'," & _
          Val(txtPulse.Text) & "," & _
          Val(txtTemp.Text) & "," & _
          Val(txtHB.Text) & "," & _
          "'" & Esc(UCase(cmbFinalStatus.Text)) & "'," & _
          "'" & Esc(txtDeferReason.Text) & "'," & _
          "'Active'" & _
          ")"

    C.Execute sql

    C.Execute "UPDATE DONAR_DETAIL SET " & _
              "LAST_DON = TO_DATE('" & Format(TO_DATE.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
              "TOT_DON = NVL(TOT_DON,0) + 1 " & _
              "WHERE D_ID='" & Esc(cmbDonarID.Text) & "' AND STATUS='ACTIVE'"

    MsgBox "Donation Visit record saved successfully.", vbInformation, "Save Success"
    ResetForm
    Exit Sub

SaveErr:
    MsgBox "Save Error : " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search a record first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If Trim(lblVisitID.Caption) = "" Then
        MsgBox "No visit selected for update.", vbExclamation, "Update Error"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If ActiveVisitExists(cmbDonarID.Text, lblVisitID.Caption) = True Then
        MsgBox "This donor already has another active visit record.", vbExclamation, "Duplicate Active Visit"
        cmbDonarID.SetFocus
        Exit Sub
    End If

    Call Check_Vitals_With_Reference

    If MsgBox("Do you want to update this visit record?", vbQuestion + vbYesNo, "Confirm Update") = vbNo Then
        MsgBox "Update cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    CONN

    sql = "UPDATE DONATION_VISIT SET " & _
          "D_ID='" & Esc(cmbDonarID.Text) & "'," & _
          "DON_DT=TO_DATE('" & Format(TO_DATE.Value, "DD-MM-YYYY") & "','DD-MM-YYYY')," & _
          "REF_NM='" & Esc(txtRefPatient.Text) & "'," & _
          "OPD=" & IIf(Trim(txtOPD.Text) = "", "NULL", Trim(txtOPD.Text)) & "," & _
          "BED='" & Esc(txtBED.Text) & "'," & _
          "WGT=" & Val(txtWeight.Text) & "," & _
          "BP='" & Esc(txtBP.Text) & "'," & _
          "PULSE=" & Val(txtPulse.Text) & "," & _
          "TEM=" & Val(txtTemp.Text) & "," & _
          "HB_LEV=" & Val(txtHB.Text) & "," & _
          "FIN_STAT='" & Esc(UCase(cmbFinalStatus.Text)) & "'," & _
          "DEFER='" & Esc(txtDeferReason.Text) & "' " & _
          "WHERE V_ID='" & Esc(lblVisitID.Caption) & "' AND STATUS='Active'"

    C.Execute sql

    MsgBox "Donation Visit record updated successfully.", vbInformation, "Update Success"
    ResetForm
    Exit Sub

UpErr:
    MsgBox "Update Error : " & Err.Description, vbCritical, "Update Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search a record first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If Trim(lblVisitID.Caption) = "" Then
        MsgBox "No visit selected.", vbExclamation, "Delete Error"
        Exit Sub
    End If

    If MsgBox("Do you want to deactivate this visit record?" & vbCrLf & _
              "Visit ID : " & lblVisitID.Caption, vbYesNo + vbQuestion, "Confirm Inactive") = vbNo Then
        MsgBox "Deactivate cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    CONN
    C.Execute "UPDATE DONATION_VISIT SET STATUS='Inactive' WHERE V_ID='" & Esc(lblVisitID.Caption) & "'"

    MsgBox "Visit record marked as INACTIVE successfully.", vbInformation, "Deactivate Success"
    ResetForm
    Exit Sub

DelErr:
    MsgBox "Delete/Inactive Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdClear_Click()
    ResetForm
End Sub

Private Sub cmdreport_Click()
    On Error GoTo ErrHandler

    Donation_Visit_Report.Show
    Exit Sub

ErrHandler:
    MsgBox "Report Error : " & Err.Description, vbCritical, "Error"
End Sub

'========================================================
' SEARCH
'========================================================
Private Sub Combo1_Click()
    On Error GoTo ErrHandler

    Dim rsS As ADODB.Recordset

    Combo2.Clear
    Combo2.Text = ""
    Combo2.Enabled = False
    cmdSearch.Enabled = False

    If Trim(Combo1.Text) = "" Then Exit Sub

    Set rsS = New ADODB.Recordset
    CONN

    If Combo1.Text = "Visit ID" Then
        sql = "SELECT V_ID FROM DONATION_VISIT WHERE STATUS='Active' ORDER BY V_ID"
    ElseIf Combo1.Text = "Donar ID" Then
        sql = "SELECT DISTINCT D_ID FROM DONATION_VISIT WHERE STATUS='Active' ORDER BY D_ID"
    ElseIf Combo1.Text = "Patient Name" Then
        sql = "SELECT DISTINCT REF_NM FROM DONATION_VISIT WHERE STATUS='Active' AND REF_NM IS NOT NULL ORDER BY REF_NM"
    End If

    rsS.Open sql, C, adOpenForwardOnly, adLockReadOnly

    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop

    rsS.Close
    Set rsS = Nothing

    Combo2.Enabled = True
    MsgBox "Search values loaded successfully.", vbInformation, "Search Ready"
    Exit Sub

ErrHandler:
    MsgBox "Search Load Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Combo2_Click()
    If Trim(Combo2.Text) <> "" Then
        cmdSearch.Enabled = True
        MsgBox "Search value selected. Now click Search.", vbInformation, "Search"
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler

    Dim rsSearch As ADODB.Recordset

    If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Validation"
        Exit Sub
    End If

    CONN
    Set rsSearch = New ADODB.Recordset

    If Combo1.Text = "Visit ID" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE V_ID='" & Esc(Combo2.Text) & "' AND STATUS='Active'"
    ElseIf Combo1.Text = "Donar ID" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE D_ID='" & Esc(Combo2.Text) & "' AND STATUS='Active'"
    ElseIf Combo1.Text = "Patient Name" Then
        sql = "SELECT * FROM DONATION_VISIT WHERE REF_NM='" & Esc(Combo2.Text) & "' AND STATUS='Active'"
    End If

    rsSearch.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If rsSearch.EOF Then
        rsSearch.Close
        Set rsSearch = Nothing
        MsgBox "Record not found.", vbExclamation, "Search Result"
        Exit Sub
    End If

    FillFormFromRecord rsSearch

    rsSearch.Close
    Set rsSearch = Nothing

    SetEditMode
    MsgBox "Record found successfully.", vbInformation, "Search Success"
    Exit Sub

ErrHandler:
    MsgBox "Search Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub FillFormFromRecord(ByVal rsSearch As ADODB.Recordset)
    lblVisitID.Caption = NzText(rsSearch!v_id)

    Fill_DonorCombo_ForEdit NzText(rsSearch!D_ID)
    cmbDonarID.Text = NzText(rsSearch!D_ID)

    If Not IsNull(rsSearch!DON_DT) Then
        TO_DATE.Value = rsSearch!DON_DT
    Else
        TO_DATE.Value = Date
    End If

    txtRefPatient.Text = NzText(rsSearch!REF_NM)
    txtOPD.Text = NzText(rsSearch!OPD)
    txtBED.Text = NzText(rsSearch!BED)
    txtWeight.Text = NzText(rsSearch!WGT)
    txtBP.Text = NzText(rsSearch!BP)
    txtPulse.Text = NzText(rsSearch!PULSE)
    txtTemp.Text = NzText(rsSearch!TEM)
    txtHB.Text = NzText(rsSearch!HB_LEV)
    cmbFinalStatus.Text = NzText(rsSearch!FIN_STAT)
    txtDeferReason.Text = NzText(rsSearch!DEFER)

    ToggleDeferReason
End Sub

'========================================================
' EXTRA LOST FOCUS EVENTS
'========================================================
Private Sub TO_DATE_LostFocus()
    On Error GoTo ErrHandler

    Dim rsLast As ADODB.Recordset
    Dim diffDays As Long

    If Trim(cmbDonarID.Text) = "" Then Exit Sub

    CONN

    Set rsLast = New ADODB.Recordset
    sql = "SELECT LAST_DON FROM DONAR_DETAIL WHERE D_ID='" & Esc(cmbDonarID.Text) & "' AND STATUS='ACTIVE'"
    rsLast.Open sql, C, adOpenForwardOnly, adLockReadOnly

    If Not rsLast.EOF Then
        If Not IsNull(rsLast!LAST_DON) Then
            diffDays = DateDiff("d", rsLast!LAST_DON, TO_DATE.Value)

            If diffDays < 90 Then
                MsgBox "Donor cannot donate now. At least 90 days are required after the last donation.", vbExclamation, "Donation Not Allowed"
                TO_DATE.SetFocus
                Exit Sub
            End If
        End If
    End If

    rsLast.Close
    Set rsLast = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error while checking last donation date : " & Err.Description, vbCritical, "Date Check Error"
End Sub

Private Sub txtRefPatient_LostFocus()
    txtRefPatient.Text = ProperCaseText(txtRefPatient.Text)
End Sub

'========================================================
' STATUS / FIELD EVENTS
'========================================================
Private Sub cmbFinalStatus_Click()
    ToggleDeferReason
    MsgBox "Final status changed.", vbInformation, "Status Update"
End Sub

'========================================================
' NAVIGATION + RESTRICTIONS
'========================================================
Private Sub cmbDonarID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        TO_DATE.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub TO_DATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtRefPatient.SetFocus
    End If
End Sub

Private Sub txtRefPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtOPD.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub

    KeyAscii = 0
End Sub

Private Sub txtOPD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtBED.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtBED_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtWeight.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Then Exit Sub

    KeyAscii = 0
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtBP.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtBP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPulse.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 47 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtPulse_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtTemp.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTemp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtHB.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtHB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Check_Vitals_With_Reference

        If CmdSave.Enabled Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled Then
            'CmdUpdate.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub cmbFinalStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0

        If txtDeferReason.Enabled = True Then
            txtDeferReason.SetFocus
        Else
            If CmdSave.Enabled Then
                'CmdSave.SetFocus
            ElseIf CmdUpdate.Enabled Then
                'CmdUpdate.SetFocus
            End If
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDeferReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If CmdSave.Enabled Then
            'CmdSave.SetFocus
        ElseIf CmdUpdate.Enabled Then
            'CmdUpdate.SetFocus
        End If
    End If
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

