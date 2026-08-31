VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DONOR_REGISTRATION_DETAIL 
   Caption         =   "DONER REGISTRATION FORM"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   5895
      Left            =   14280
      TabIndex        =   44
      Top             =   1320
      Width           =   4575
      Begin VB.ComboBox Combo3 
         Height          =   480
         Left            =   2160
         TabIndex        =   49
         Text            =   "Combo3"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   480
         Left            =   2160
         TabIndex        =   48
         Text            =   "Combo2"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label LabelMobile 
         Caption         =   "Mobile No"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label LabelSearchValue 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image cmdSearch 
         Height          =   855
         Left            =   1320
         Picture         =   "master form by rana.frx":0000
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label sea_val 
         Caption         =   "Search Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label sea_by 
         Caption         =   "Search By"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   1440
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "master form by rana.frx":3CDF
      Height          =   1935
      Left            =   360
      TabIndex        =   42
      Top             =   8280
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
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
      ColumnCount     =   21
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "NM"
         Caption         =   "NAME"
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
         DataField       =   "F_NM"
         Caption         =   "FATHER'S NAME"
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
         DataField       =   "DOB"
         Caption         =   "D.O.B"
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
         DataField       =   "GENDER"
         Caption         =   "GENDER"
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
         DataField       =   "BL_GRP"
         Caption         =   "BLOOD GROUP"
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
         DataField       =   "OCC"
         Caption         =   "OCCUPATION"
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
         DataField       =   "ORG"
         Caption         =   "ORGANISATION"
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
         DataField       =   "ADDR"
         Caption         =   "ADDRESS"
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
         DataField       =   "PIN"
         Caption         =   "PINCODE"
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
         DataField       =   "PO"
         Caption         =   "P.O"
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
         DataField       =   "DIST"
         Caption         =   "DISTRICT"
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
         DataField       =   "TELE"
         Caption         =   "TELEPHONE"
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
         DataField       =   "PHN"
         Caption         =   "PHONE NO."
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
      BeginProperty Column14 
         DataField       =   "CALL"
         Caption         =   "CALL"
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
      BeginProperty Column15 
         DataField       =   "FAX"
         Caption         =   "FAX"
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
      BeginProperty Column16 
         DataField       =   "EMAIL"
         Caption         =   "EMAIL"
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
      BeginProperty Column17 
         DataField       =   "D_PREV"
         Caption         =   "D_PREV"
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
      BeginProperty Column18 
         DataField       =   "TOT_DON"
         Caption         =   "TOT_DON"
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
      BeginProperty Column19 
         DataField       =   "LAST_DON"
         Caption         =   "LAST_DON"
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
      BeginProperty Column20 
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
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2445.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   3240
      Top             =   8400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      RecordSource    =   "SELECT * FROM DONAR_DETAIL"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   7080
      TabIndex        =   22
      Top             =   1320
      Width           =   6975
      Begin VB.ComboBox CmbCallable 
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
         Left            =   5400
         TabIndex        =   43
         Text            =   "CmbCallable"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtOccupation 
         Appearance      =   0  'Flat
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
         TabIndex        =   31
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtOrganization 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   30
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox CmbDistrict 
         Appearance      =   0  'Flat
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
         Left            =   2520
         TabIndex        =   29
         Text            =   "CmbDistrict"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
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
         TabIndex        =   28
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
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
         MaxLength       =   6
         TabIndex        =   27
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox txtPostOffice 
         Appearance      =   0  'Flat
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
         TabIndex        =   26
         Top             =   3360
         Width           =   4335
      End
      Begin VB.ComboBox cmbPrevDonation 
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
         Left            =   2520
         TabIndex        =   25
         Text            =   "cmbPrevDonation"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDonation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2520
         TabIndex        =   24
         Top             =   5160
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker DT2 
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   4560
         Width           =   4335
         _ExtentX        =   7646
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
         CurrentDate     =   46028
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   360
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Organization"
         Height          =   360
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1590
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Callable"
         Height          =   375
         Left            =   4320
         TabIndex        =   39
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   360
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIN Code"
         Height          =   360
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post Office"
         Height          =   360
         Left            =   120
         TabIndex        =   36
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         Height          =   360
         Left            =   120
         TabIndex        =   35
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Donation"
         Height          =   360
         Left            =   120
         TabIndex        =   34
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Donation"
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Donation"
         Height          =   360
         Left            =   120
         TabIndex        =   32
         Top             =   5160
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   6615
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   10
         Text            =   "cmbGender"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cmbBloodGroup 
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
         Left            =   5400
         TabIndex        =   9
         Text            =   "cmbBloodGroup"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtFather 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   6
         Top             =   3360
         Width           =   4215
      End
      Begin VB.TextBox txtTelephone 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   5
         Top             =   3960
         Width           =   4215
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   4
         Top             =   4560
         Width           =   4215
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Left            =   2280
         TabIndex        =   3
         Top             =   5160
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker D1 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   2280
         Width           =   4215
         _ExtentX        =   7435
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   126091267
         CurrentDate     =   46027
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Donar ID"
         Height          =   360
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   360
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name"
         Height          =   360
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOB"
         Height          =   360
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   360
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Group"
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label txtDonorID 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         Height          =   360
         Left            =   240
         TabIndex        =   13
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   360
         Left            =   240
         TabIndex        =   12
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   360
         Left            =   240
         TabIndex        =   11
         Top             =   4680
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   7440
      Width           =   14055
      Begin VB.Image cmdreport 
         Height          =   495
         Left            =   12000
         Picture         =   "master form by rana.frx":3CF4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image cmdSave 
         Height          =   495
         Left            =   9720
         Picture         =   "master form by rana.frx":F2A1
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image cmdDelete 
         Height          =   495
         Left            =   7440
         Picture         =   "master form by rana.frx":12918
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image cmdUpdate 
         Height          =   495
         Left            =   5040
         Picture         =   "master form by rana.frx":164C1
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image cmdClear 
         Height          =   495
         Left            =   2520
         Picture         =   "master form by rana.frx":19F11
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image cmdAddNew 
         Height          =   495
         Left            =   120
         Picture         =   "master form by rana.frx":1D6E8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "master form by rana.frx":20BB0
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   19695
   End
End
Attribute VB_Name = "DONOR_REGISTRATION_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMode As String
Private sql As String



'========================================================
' FORM LOAD
'========================================================
Private Sub Form_Load()
    On Error GoTo ErrHandler

    Call CONN
    LoadMasterCombos
    InitializeSearchSection
    ClearFormValues
    RefreshGrid
    SetStartupMode

    Exit Sub

ErrHandler:
    'MsgBox "Form Load Error : " & Err.Description, vbCritical, "Error"
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
    'EnableSearchControls False

    CmdAddNew.Enabled = False
    'cmdSearch.Enabled = False
    cmdreport.Enabled = False

    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    CmdDelete.Enabled = False
    CmdClear.Enabled = True

    txtName.SetFocus
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

    txtName.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    txtName.Enabled = bFlag
    txtFather.Enabled = bFlag
    D1.Enabled = bFlag
    cmbGender.Enabled = bFlag
    cmbBloodGroup.Enabled = bFlag
    txtMobile.Enabled = bFlag
    txtTelephone.Enabled = bFlag
    txtFax.Enabled = bFlag
    txtEmail.Enabled = bFlag
    txtOccupation.Enabled = bFlag
    txtOrganization.Enabled = bFlag
    CmbDistrict.Enabled = bFlag
    txtAddress.Enabled = bFlag
    txtPin.Enabled = bFlag
    txtPostOffice.Enabled = bFlag
    cmbPrevDonation.Enabled = bFlag
    CmbCallable.Enabled = bFlag

    If bFlag = False Then
        DT2.Enabled = False
        txtTotalDonation.Enabled = False
    Else
        TogglePreviousDonationControls
    End If
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag
    Combo2.Enabled = bFlag
    Combo3.Enabled = bFlag
End Sub

'========================================================
' INITIALIZATION
'========================================================
Private Sub LoadMasterCombos()
    cmbGender.Clear
    cmbGender.AddItem "Male"
    cmbGender.AddItem "Female"

    cmbBloodGroup.Clear
    cmbBloodGroup.AddItem "A+"
    cmbBloodGroup.AddItem "A-"
    cmbBloodGroup.AddItem "B+"
    cmbBloodGroup.AddItem "B-"
    cmbBloodGroup.AddItem "AB+"
    cmbBloodGroup.AddItem "AB-"
    cmbBloodGroup.AddItem "O+"
    cmbBloodGroup.AddItem "O-"

    cmbPrevDonation.Clear
    cmbPrevDonation.AddItem "YES"
    cmbPrevDonation.AddItem "NO"

    CmbCallable.Clear
    CmbCallable.AddItem "Yes"
    CmbCallable.AddItem "No"

    CmbDistrict.Clear
    CmbDistrict.AddItem "Araria"
    CmbDistrict.AddItem "Arwal"
    CmbDistrict.AddItem "Aurangabad"
    CmbDistrict.AddItem "Banka"
    CmbDistrict.AddItem "Begusarai"
    CmbDistrict.AddItem "Bhagalpur"
    CmbDistrict.AddItem "Bhojpur"
    CmbDistrict.AddItem "Buxar"
    CmbDistrict.AddItem "Darbhanga"
    CmbDistrict.AddItem "East Champaran"
    CmbDistrict.AddItem "Gaya"
    CmbDistrict.AddItem "Gopalganj"
    CmbDistrict.AddItem "Jamui"
    CmbDistrict.AddItem "Jehanabad"
    CmbDistrict.AddItem "Kaimur"
    CmbDistrict.AddItem "Katihar"
    CmbDistrict.AddItem "Khagaria"
    CmbDistrict.AddItem "Kishanganj"
    CmbDistrict.AddItem "Lakhisarai"
    CmbDistrict.AddItem "Madhepura"
    CmbDistrict.AddItem "Madhubani"
    CmbDistrict.AddItem "Munger"
    CmbDistrict.AddItem "Muzaffarpur"
    CmbDistrict.AddItem "Nalanda"
    CmbDistrict.AddItem "Nawada"
    CmbDistrict.AddItem "Patna"
    CmbDistrict.AddItem "Purnia"
    CmbDistrict.AddItem "Rohtas"
    CmbDistrict.AddItem "Saharsa"
    CmbDistrict.AddItem "Samastipur"
    CmbDistrict.AddItem "Saran"
    CmbDistrict.AddItem "Sheohar"
    CmbDistrict.AddItem "Sheikhpura"
    CmbDistrict.AddItem "Sitamarhi"
    CmbDistrict.AddItem "Supaul"
    CmbDistrict.AddItem "Siwan"
    CmbDistrict.AddItem "Vaishali"
    CmbDistrict.AddItem "West Champaran"
    CmbDistrict.AddItem "Other"
End Sub

Private Sub InitializeSearchSection()
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear

    Combo1.AddItem "ID"
    Combo1.AddItem "Name"

    Combo1.ListIndex = -1
    Combo2.Enabled = False
    Combo3.Enabled = False
End Sub

Private Sub ClearFormValues()
    txtDonorID.Caption = ""

    txtName.Text = ""
    txtFather.Text = ""
    txtMobile.Text = ""
    txtTelephone.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtOccupation.Text = ""
    txtOrganization.Text = ""
    txtAddress.Text = ""
    txtPin.Text = ""
    txtPostOffice.Text = ""
    txtTotalDonation.Text = "0"

    D1.Value = Date
    DT2.Value = Date

    cmbGender.ListIndex = -1
    cmbBloodGroup.ListIndex = -1
    cmbPrevDonation.ListIndex = -1
    CmbCallable.ListIndex = -1
    CmbDistrict.ListIndex = -1

    Combo1.ListIndex = -1
    Combo2.Clear
    Combo2.Text = ""
    Combo3.Clear
    Combo3.Text = ""
    Combo2.Enabled = False
    Combo3.Enabled = False

    DT2.Enabled = False
    txtTotalDonation.Enabled = False
End Sub

Private Sub ResetForm()
    LoadMasterCombos
    InitializeSearchSection
    ClearFormValues
    RefreshGrid
    SetStartupMode
End Sub

'========================================================
' GRID
'========================================================
'========================================================
' GRID (FIXED: Serial Wise & Active Only)
'========================================================
Private Sub RefreshGrid()
    On Error Resume Next
    
    Dim rsGrid As New ADODB.Recordset
    If C.State = 0 Then Call CONN

    ' Client cursor zaroori hai DataGrid mein data show karne ke liye
    rsGrid.CursorLocation = adUseClient
    
    ' Query: Sirf ACTIVE donors layega aur d_id ke hisaab se SERIAL wise set karega
    sql = "SELECT * FROM DONAR_DETAIL WHERE UPPER(status)='ACTIVE' ORDER BY d_id ASC"
    
    rsGrid.Open sql, C, adOpenStatic, adLockReadOnly

    ' Grid ko naye data ke sath bind karna
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = rsGrid
    DataGrid1.Refresh
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

Private Function GetAge(ByVal dtDOB As Date) As Integer
    GetAge = DateDiff("yyyy", dtDOB, Date)
    If Date < DateSerial(Year(Date), Month(dtDOB), Day(dtDOB)) Then
        GetAge = GetAge - 1
    End If
End Function

Private Function IsValidEmail(ByVal sEmail As String) As Boolean
    sEmail = Trim(sEmail)

    If sEmail = "" Then
        IsValidEmail = True
        Exit Function
    End If

    If InStr(1, sEmail, "@") = 0 Then
        IsValidEmail = False
        Exit Function
    End If

    If InStrRev(sEmail, ".") = 0 Then
        IsValidEmail = False
        Exit Function
    End If

    If Left$(sEmail, 1) = "@" Or Right$(sEmail, 1) = "." Then
        IsValidEmail = False
        Exit Function
    End If

    IsValidEmail = True
End Function

Private Function ProperCaseText(ByVal s As String) As String
    If Trim(s) = "" Then
        ProperCaseText = ""
    Else
        ProperCaseText = UCase(Left(Trim(s), 1)) & LCase(Mid(Trim(s), 2))
    End If
End Function

Private Sub TogglePreviousDonationControls()
    If UCase(Trim(cmbPrevDonation.Text)) = "YES" Then
        DT2.Enabled = True
        txtTotalDonation.Enabled = True
        If txtTotalDonation.Text = "0" Then txtTotalDonation.Text = ""
    Else
        DT2.Enabled = False
        DT2.Value = Date
        txtTotalDonation.Enabled = False
        txtTotalDonation.Text = "0"
    End If
End Sub

'========================================================
' DONOR ID GENERATION
'========================================================
Private Sub GenerateDonorID()
    On Error GoTo ErrHandler

    Dim rsMax As ADODB.Recordset
    Dim NextNo As Long

    Call CONN

    Set rsMax = New ADODB.Recordset
    rsMax.Open "SELECT NVL(MAX(TO_NUMBER(SUBSTR(D_ID,6))),0)+1 AS NEXT_ID FROM DONAR_DETAIL WHERE D_ID LIKE 'BDI00%'", C, adOpenForwardOnly, adLockReadOnly

    If Not rsMax.EOF Then
        NextNo = CLng(rsMax!NEXT_ID)
    Else
        NextNo = 1
    End If

    txtDonorID.Caption = "BDI00" & CStr(NextNo)

    rsMax.Close
    Set rsMax = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error generating Donor ID : " & Err.Description, vbCritical, "Donor ID Error"
End Sub

'========================================================
' BUTTONS
'========================================================
Private Sub cmdAddNew_Click()
    On Error GoTo ErrHandler

    ClearFormValues
    LoadMasterCombos
    GenerateDonorID
    SetAddMode

    MsgBox "New Donor ID Generated.", vbInformation, "Add New"
    Exit Sub

ErrHandler:
    MsgBox "Add New Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdClear_Click()
    ResetForm
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveErr

    Dim rsChk As ADODB.Recordset
    Dim LastDonSQL As String
    Dim TotalDonSQL As String
    Dim PinSQL As String
    Dim TeleSQL As String
    Dim FaxSQL As String
    Dim MobileNo As String

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    MobileNo = Trim(txtMobile.Text)

    Call CONN

    Set rsChk = New ADODB.Recordset
    rsChk.Open "SELECT D_ID FROM DONAR_DETAIL WHERE D_ID='" & Esc(Trim(txtDonorID.Caption)) & "'", C, adOpenForwardOnly, adLockReadOnly

    If Not rsChk.EOF Then
        rsChk.Close
        Set rsChk = Nothing
        MsgBox "This Donor ID already exists.", vbExclamation, "Duplicate ID"
        Exit Sub
    End If

    rsChk.Close
    Set rsChk = Nothing

    If UCase(Trim(cmbPrevDonation.Text)) = "YES" Then
        LastDonSQL = "TO_DATE('" & Format(DT2.Value, "DD-MM-YYYY") & "','DD-MM-YYYY')"
        TotalDonSQL = CStr(Val(txtTotalDonation.Text))
    Else
        LastDonSQL = "NULL"
        TotalDonSQL = "0"
    End If

    If Trim(txtPin.Text) = "" Then
        PinSQL = "NULL"
    Else
        PinSQL = Trim(txtPin.Text)
    End If

    If Trim(txtTelephone.Text) = "" Then
        TeleSQL = "NULL"
    Else
        TeleSQL = Trim(txtTelephone.Text)
    End If

    If Trim(txtFax.Text) = "" Then
        FaxSQL = "NULL"
    Else
        FaxSQL = "'" & Esc(Trim(txtFax.Text)) & "'"
    End If

    sql = "INSERT INTO DONAR_DETAIL (" & _
          "D_ID, NM, F_NM, DOB, GENDER, BL_GRP, OCC, ORG, ADDR, PIN, PO, DIST, TELE, PHN, CALL, FAX, EMAIL, D_PREV, TOT_DON, LAST_DON, STATUS" & _
          ") VALUES (" & _
          "'" & Esc(Trim(txtDonorID.Caption)) & "'," & _
          "'" & Esc(Trim(txtName.Text)) & "'," & _
          "'" & Esc(Trim(txtFather.Text)) & "'," & _
          "TO_DATE('" & Format(D1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY')," & _
          "'" & Esc(Trim(cmbGender.Text)) & "'," & _
          "'" & Esc(Trim(cmbBloodGroup.Text)) & "'," & _
          "'" & Esc(Trim(txtOccupation.Text)) & "'," & _
          "'" & Esc(Trim(txtOrganization.Text)) & "'," & _
          "'" & Esc(Trim(txtAddress.Text)) & "'," & _
          PinSQL & "," & _
          "'" & Esc(Trim(txtPostOffice.Text)) & "'," & _
          "'" & Esc(Trim(CmbDistrict.Text)) & "'," & _
          TeleSQL & "," & _
          MobileNo & "," & _
          "'" & Esc(Trim(CmbCallable.Text)) & "'," & _
          FaxSQL & "," & _
          "'" & Esc(Trim(txtEmail.Text)) & "'," & _
          "'" & Esc(Trim(cmbPrevDonation.Text)) & "'," & _
          TotalDonSQL & "," & _
          LastDonSQL & "," & _
          "'ACTIVE'" & _
          ")"

    C.Execute sql
    RefreshGrid

    MsgBox "Donor Saved Successfully.", vbInformation, "Save Success"
    ResetForm
    Exit Sub

SaveErr:
    MsgBox "Save Error : " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpdateErr

    Dim LastDonSQL As String
    Dim TotalDonSQL As String
    Dim PinSQL As String
    Dim TeleSQL As String
    Dim FaxSQL As String

    If mMode <> "EDIT" Then
        MsgBox "Please search a record first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If Trim(txtDonorID.Caption) = "" Then
        MsgBox "No donor selected for update.", vbExclamation, "Update Error"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If UCase(Trim(cmbPrevDonation.Text)) = "YES" Then
        LastDonSQL = "TO_DATE('" & Format(DT2.Value, "DD-MM-YYYY") & "','DD-MM-YYYY')"
        TotalDonSQL = CStr(Val(txtTotalDonation.Text))
    Else
        LastDonSQL = "NULL"
        TotalDonSQL = "0"
    End If

    If Trim(txtPin.Text) = "" Then
        PinSQL = "NULL"
    Else
        PinSQL = Trim(txtPin.Text)
    End If

    If Trim(txtTelephone.Text) = "" Then
        TeleSQL = "NULL"
    Else
        TeleSQL = Trim(txtTelephone.Text)
    End If

    If Trim(txtFax.Text) = "" Then
        FaxSQL = "NULL"
    Else
        FaxSQL = "'" & Esc(Trim(txtFax.Text)) & "'"
    End If

    sql = "UPDATE DONAR_DETAIL SET " & _
          "NM='" & Esc(Trim(txtName.Text)) & "', " & _
          "F_NM='" & Esc(Trim(txtFather.Text)) & "', " & _
          "DOB=TO_DATE('" & Format(D1.Value, "DD-MM-YYYY") & "','DD-MM-YYYY'), " & _
          "GENDER='" & Esc(Trim(cmbGender.Text)) & "', " & _
          "BL_GRP='" & Esc(Trim(cmbBloodGroup.Text)) & "', " & _
          "OCC='" & Esc(Trim(txtOccupation.Text)) & "', " & _
          "ORG='" & Esc(Trim(txtOrganization.Text)) & "', " & _
          "DIST='" & Esc(Trim(CmbDistrict.Text)) & "', " & _
          "ADDR='" & Esc(Trim(txtAddress.Text)) & "', " & _
          "PIN=" & PinSQL & ", " & _
          "PO='" & Esc(Trim(txtPostOffice.Text)) & "', " & _
          "TELE=" & TeleSQL & ", " & _
          "PHN=" & Trim(txtMobile.Text) & ", " & _
          "CALL='" & Esc(Trim(CmbCallable.Text)) & "', " & _
          "FAX=" & FaxSQL & ", " & _
          "EMAIL='" & Esc(Trim(txtEmail.Text)) & "', " & _
          "D_PREV='" & Esc(Trim(cmbPrevDonation.Text)) & "', " & _
          "TOT_DON=" & TotalDonSQL & ", " & _
          "LAST_DON=" & LastDonSQL & " " & _
          "WHERE D_ID='" & Esc(Trim(txtDonorID.Caption)) & "' AND STATUS='ACTIVE'"

    C.Execute sql

    MsgBox "Record Updated Successfully.", vbInformation, "Update Success"
    ResetForm
    Exit Sub

UpdateErr:
    MsgBox "Update Error : " & Err.Description, vbCritical, "Update Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo DelErr

    Dim response As Integer

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(txtDonorID.Caption) = "" Then
        MsgBox "No donor selected.", vbExclamation, "Delete Error"
        Exit Sub
    End If

    response = MsgBox("Do you want to deactivate this donor?" & vbCrLf & _
                      "Donor ID : " & txtDonorID.Caption & vbCrLf & _
                      "Name : " & txtName.Text & vbCrLf & vbCrLf & _
                      "This record will not be deleted permanently.", _
                      vbQuestion + vbYesNo, "Confirm Inactive")

    If response = vbYes Then
        Call CONN
        sql = "UPDATE DONAR_DETAIL SET STATUS='INACTIVE' WHERE D_ID='" & Esc(Trim(txtDonorID.Caption)) & "'"
        C.Execute sql

        MsgBox "Record marked as INACTIVE successfully.", vbInformation, "Inactive Success"
        ResetForm
    Else
        MsgBox "Deactivate cancelled.", vbInformation, "Cancelled"
    End If

    Exit Sub

DelErr:
    MsgBox "Delete Error : " & Err.Description, vbCritical, "Delete Error"
End Sub

Private Sub cmdreport_Click()
    RegistrationReport.Show
End Sub

'========================================================
' SEARCH
'========================================================
Private Sub Combo1_Click()
    On Error GoTo ErrHandler

    Dim rsS As ADODB.Recordset

    Combo2.Clear
    Combo3.Clear
    Combo2.Text = ""
    Combo3.Text = ""
    Combo2.Enabled = False
    Combo3.Enabled = False
    cmdSearch.Enabled = False

    If Trim(Combo1.Text) = "" Then Exit Sub

    Call CONN
    Set rsS = New ADODB.Recordset

    If Combo1.Text = "ID" Then
        sql = "SELECT D_ID FROM DONAR_DETAIL WHERE STATUS='ACTIVE' ORDER BY D_ID"
        rsS.Open sql, C, adOpenForwardOnly, adLockReadOnly

        Do While Not rsS.EOF
            Combo2.AddItem rsS.Fields(0).Value
            rsS.MoveNext
        Loop

        Combo2.Enabled = True

    ElseIf Combo1.Text = "Name" Then
        sql = "SELECT DISTINCT NM FROM DONAR_DETAIL WHERE STATUS='ACTIVE' ORDER BY NM"
        rsS.Open sql, C, adOpenForwardOnly, adLockReadOnly

        Do While Not rsS.EOF
            Combo2.AddItem rsS.Fields(0).Value
            rsS.MoveNext
        Loop

        Combo2.Enabled = True
    End If

    rsS.Close
    Set rsS = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Search Load Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Combo2_Click()
    On Error GoTo ErrHandler

    Dim rsM As ADODB.Recordset

    Combo3.Clear
    Combo3.Text = ""
    Combo3.Enabled = False
    cmdSearch.Enabled = False

    If Trim(Combo1.Text) = "" Then Exit Sub
    If Trim(Combo2.Text) = "" Then Exit Sub

    If Combo1.Text = "ID" Then
        cmdSearch.Enabled = True
        Exit Sub
    End If

    If Combo1.Text = "Name" Then
        Call CONN
        Set rsM = New ADODB.Recordset

        sql = "SELECT PHN FROM DONAR_DETAIL WHERE NM='" & Esc(Trim(Combo2.Text)) & "' AND STATUS='ACTIVE' ORDER BY PHN"
        rsM.Open sql, C, adOpenForwardOnly, adLockReadOnly

        Do While Not rsM.EOF
            Combo3.AddItem rsM.Fields(0).Value
            rsM.MoveNext
        Loop

        rsM.Close
        Set rsM = Nothing

        Combo3.Enabled = True
    End If

    Exit Sub

ErrHandler:
    MsgBox "Mobile List Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Combo3_Click()
    If Trim(Combo3.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo SearchErr

    Dim rsSearch As ADODB.Recordset

    If Trim(Combo1.Text) = "" Then
        MsgBox "Please select Search By.", vbExclamation, "Search"
        Combo1.SetFocus
        Exit Sub
    End If

    If Trim(Combo2.Text) = "" Then
        MsgBox "Please select Value.", vbExclamation, "Search"
        Combo2.SetFocus
        Exit Sub
    End If

    If Combo1.Text = "Name" And Trim(Combo3.Text) = "" Then
        MsgBox "Please select Mobile Number.", vbExclamation, "Search"
        Combo3.SetFocus
        Exit Sub
    End If

    Call CONN
    Set rsSearch = New ADODB.Recordset

    If Combo1.Text = "ID" Then
        sql = "SELECT * FROM DONAR_DETAIL WHERE D_ID='" & Esc(Trim(Combo2.Text)) & "' AND STATUS='ACTIVE'"
    ElseIf Combo1.Text = "Name" Then
        sql = "SELECT * FROM DONAR_DETAIL WHERE NM='" & Esc(Trim(Combo2.Text)) & "' AND PHN=" & Trim(Combo3.Text) & " AND STATUS='ACTIVE'"
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

SearchErr:
    MsgBox "Search Error : " & Err.Description, vbCritical, "Search Error"
End Sub

Private Sub FillFormFromRecord(ByVal rsSearch As ADODB.Recordset)
    txtDonorID.Caption = NzText(rsSearch!D_ID)
    txtName.Text = NzText(rsSearch!nm)
    txtFather.Text = NzText(rsSearch!F_NM)

    If Not IsNull(rsSearch!DOB) Then D1.Value = rsSearch!DOB

    cmbGender.Text = NzText(rsSearch!gender)
    cmbBloodGroup.Text = NzText(rsSearch!BL_GRP)
    txtOccupation.Text = NzText(rsSearch!OCC)
    txtOrganization.Text = NzText(rsSearch!ORG)
    txtAddress.Text = NzText(rsSearch!addr)
    txtPin.Text = NzText(rsSearch!PIN)
    txtPostOffice.Text = NzText(rsSearch!PO)
    CmbDistrict.Text = NzText(rsSearch!DIST)
    txtTelephone.Text = NzText(rsSearch!TELE)
    txtMobile.Text = NzText(rsSearch!PHN)
    CmbCallable.Text = NzText(rsSearch!Call)
    txtFax.Text = NzText(rsSearch!FAX)
    txtEmail.Text = NzText(rsSearch!EMAIL)
    cmbPrevDonation.Text = NzText(rsSearch!D_PREV)

    If IsNull(rsSearch!TOT_DON) Then
        txtTotalDonation.Text = "0"
    Else
        txtTotalDonation.Text = CStr(rsSearch!TOT_DON)
    End If

    If IsNull(rsSearch!LAST_DON) Then
        DT2.Value = Date
    Else
        DT2.Value = rsSearch!LAST_DON
    End If

    TogglePreviousDonationControls
End Sub

'========================================================
' VALIDATION
'========================================================
Private Function ValidateForm() As Boolean
    Dim ageVal As Integer

    ValidateForm = False

    If Trim(txtDonorID.Caption) = "" Then
        MsgBox "Donor ID is missing.", vbExclamation, "Validation Error"
        Exit Function
    End If

    If Trim(txtName.Text) = "" Then
        MsgBox "Please enter donor name.", vbExclamation, "Validation Error"
        txtName.SetFocus
        Exit Function
    End If

    If Trim(txtFather.Text) = "" Then
        MsgBox "Please enter father name.", vbExclamation, "Validation Error"
        txtFather.SetFocus
        Exit Function
    End If

    ageVal = GetAge(D1.Value)
    If ageVal < 18 Then
        MsgBox "Donor age must be at least 18 years.", vbExclamation, "Validation Error"
        D1.SetFocus
        Exit Function
    End If

    If Trim(cmbGender.Text) = "" Then
        MsgBox "Please select gender.", vbExclamation, "Validation Error"
        cmbGender.SetFocus
        Exit Function
    End If

    If Trim(cmbBloodGroup.Text) = "" Then
        MsgBox "Please select blood group.", vbExclamation, "Validation Error"
        cmbBloodGroup.SetFocus
        Exit Function
    End If

    If Trim(txtMobile.Text) = "" Then
        MsgBox "Please enter mobile number.", vbExclamation, "Validation Error"
        txtMobile.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(txtMobile.Text) Or Len(Trim(txtMobile.Text)) <> 10 Then
        MsgBox "Mobile number must be exactly 10 digits.", vbExclamation, "Validation Error"
        txtMobile.SetFocus
        Exit Function
    End If

    If Trim(txtTelephone.Text) <> "" Then
        If Not IsDigitsOnly(txtTelephone.Text) Or Len(Trim(txtTelephone.Text)) > 15 Then
            MsgBox "Telephone number must be numeric and maximum 15 digits.", vbExclamation, "Validation Error"
            txtTelephone.SetFocus
            Exit Function
        End If
    End If

    If Trim(txtFax.Text) <> "" Then
        If Not IsDigitsOnly(txtFax.Text) Then
            MsgBox "Fax number must contain digits only.", vbExclamation, "Validation Error"
            txtFax.SetFocus
            Exit Function
        End If
    End If

    If Trim(txtPin.Text) = "" Then
        MsgBox "Please enter PIN Code.", vbExclamation, "Validation Error"
        txtPin.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(txtPin.Text) Or Len(Trim(txtPin.Text)) <> 6 Then
        MsgBox "PIN Code must be exactly 6 digits.", vbExclamation, "Validation Error"
        txtPin.SetFocus
        Exit Function
    End If

    If Trim(CmbDistrict.Text) = "" Then
        MsgBox "Please select district.", vbExclamation, "Validation Error"
        CmbDistrict.SetFocus
        Exit Function
    End If

    If Trim(txtAddress.Text) = "" Then
        MsgBox "Please enter address.", vbExclamation, "Validation Error"
        txtAddress.SetFocus
        Exit Function
    End If

    If Trim(txtPostOffice.Text) = "" Then
        MsgBox "Please enter Post Office.", vbExclamation, "Validation Error"
        txtPostOffice.SetFocus
        Exit Function
    End If

    If Trim(cmbPrevDonation.Text) = "" Then
        MsgBox "Please select previous donation status.", vbExclamation, "Validation Error"
        cmbPrevDonation.SetFocus
        Exit Function
    End If

    If Trim(CmbCallable.Text) = "" Then
        MsgBox "Please select callable status.", vbExclamation, "Validation Error"
        CmbCallable.SetFocus
        Exit Function
    End If

    If UCase(Trim(cmbPrevDonation.Text)) = "YES" Then
        If Trim(txtTotalDonation.Text) = "" Then
            MsgBox "Please enter total donation.", vbExclamation, "Validation Error"
            txtTotalDonation.SetFocus
            Exit Function
        End If

        If Not IsDigitsOnly(txtTotalDonation.Text) Then
            MsgBox "Total donation must be numeric.", vbExclamation, "Validation Error"
            txtTotalDonation.SetFocus
            Exit Function
        End If

        If Val(txtTotalDonation.Text) <= 0 Then
            MsgBox "Total donation must be greater than 0.", vbExclamation, "Validation Error"
            txtTotalDonation.SetFocus
            Exit Function
        End If

        If DT2.Value >= Date Then
            MsgBox "Last donation date must be less than today.", vbExclamation, "Validation Error"
            DT2.SetFocus
            Exit Function
        End If
    Else
        txtTotalDonation.Text = "0"
    End If

    If IsValidEmail(txtEmail.Text) = False Then
        MsgBox "Please enter a valid email address.", vbExclamation, "Validation Error"
        txtEmail.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

'========================================================
' FIELD EVENTS
'========================================================
Private Sub cmbPrevDonation_Click()
    TogglePreviousDonationControls
End Sub

Private Sub txtEmail_LostFocus()
    If Trim(txtEmail.Text) = "" Then Exit Sub

    If IsValidEmail(txtEmail.Text) = False Then
        MsgBox "Invalid email address.", vbExclamation, "Email Error"
        txtEmail.SetFocus
    End If
End Sub

'========================================================
' KEY PRESS + AUTO FORMAT
'========================================================
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtName.Text = ProperCaseText(txtName.Text)
        KeyAscii = 0
        txtFather.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtFather_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFather.Text = ProperCaseText(txtFather.Text)
        KeyAscii = 0
        D1.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub D1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmbGender.SetFocus
    End If
End Sub

Private Sub D1_LostFocus()
    Dim ageVal As Integer

    ageVal = DateDiff("yyyy", D1.Value, Date)

    If Date < DateSerial(Year(Date), Month(D1.Value), Day(D1.Value)) Then
        ageVal = ageVal - 1
    End If

    If ageVal < 18 Then
        MsgBox "Donor age must be at least 18 years.", vbExclamation, "Age Validation"
        D1.SetFocus
    End If
End Sub

Private Sub cmbGender_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmbBloodGroup.SetFocus
    End If
End Sub

Private Sub cmbBloodGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtMobile.SetFocus
    End If
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(Trim(txtMobile.Text)) <> 10 Then
            MsgBox "Mobile number must be exactly 10 digits.", vbExclamation, "Input Error"
            txtMobile.SetFocus
        Else
            txtTelephone.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    If Len(txtMobile.Text) >= 10 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtFax.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    If Len(txtTelephone.Text) >= 15 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtEmail.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtOccupation.SetFocus
    End If
End Sub

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOccupation.Text = ProperCaseText(txtOccupation.Text)
        KeyAscii = 0
        txtOrganization.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtOrganization_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOrganization.Text = ProperCaseText(txtOrganization.Text)
        KeyAscii = 0
        CmbDistrict.SetFocus
        Exit Sub
    End If
End Sub

Private Sub CmbDistrict_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAddress.Text = ProperCaseText(txtAddress.Text)
        KeyAscii = 0
        txtPin.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(Trim(txtPin.Text)) <> 6 Then
            MsgBox "PIN Code must be exactly 6 digits.", vbExclamation, "Input Error"
            txtPin.SetFocus
        Else
            txtPostOffice.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    If Len(txtPin.Text) >= 6 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPostOffice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPostOffice.Text = ProperCaseText(txtPostOffice.Text)
        KeyAscii = 0
        cmbPrevDonation.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmbPrevDonation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        CmbCallable.SetFocus
    End If
End Sub

Private Sub CmbCallable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If UCase(Trim(cmbPrevDonation.Text)) = "YES" Then
            DT2.SetFocus
        Else
            If CmdSave.Enabled Then
                'cmdSave.SetFocus
            ElseIf CmdUpdate.Enabled Then
                'cmdUpdate.SetFocus
            End If
        End If
    End If
End Sub

Private Sub DT2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        txtTotalDonation.SetFocus
    End If
End Sub

Private Sub txtTotalDonation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If CmdSave.Enabled Then
            'cmdSave.SetFocus
        ElseIf CmdUpdate.Enabled Then
            'cmdUpdate.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

