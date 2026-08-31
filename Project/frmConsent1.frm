VERSION 5.00
Begin VB.Form frmConsent 
   Caption         =   "Consent Form"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   18855
   Begin VB.Frame Frame15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   13080
      TabIndex        =   55
      Top             =   8520
      Width           =   5535
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
         TabIndex        =   57
         Text            =   "Combo4"
         Top             =   840
         Width           =   2535
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
         Left            =   2400
         TabIndex        =   56
         Text            =   "Combo3"
         Top             =   240
         Width           =   2535
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
         TabIndex        =   59
         Top             =   840
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
         Left            =   480
         TabIndex        =   58
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame14 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   495
      Left            =   15720
      TabIndex        =   50
      Top             =   7560
      Width           =   2175
      Begin VB.OptionButton Option15 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   52
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option16 
         Caption         =   "No"
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
         Left            =   1200
         TabIndex        =   51
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   495
      Left            =   15720
      TabIndex        =   47
      Top             =   6480
      Width           =   2175
      Begin VB.OptionButton Option14 
         Caption         =   "No"
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
         Left            =   1200
         TabIndex        =   49
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   48
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   495
      Left            =   15840
      TabIndex        =   44
      Top             =   5880
      Width           =   2175
      Begin VB.OptionButton Option12 
         Caption         =   "No"
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
         Left            =   1080
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   495
      Left            =   15720
      TabIndex        =   41
      Top             =   5430
      Width           =   2175
      Begin VB.OptionButton Option9 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   43
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option10 
         Caption         =   "No"
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
         Left            =   1200
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   600
      TabIndex        =   39
      Top             =   8760
      Width           =   12135
      Begin VB.Image CmdUpdate 
         Height          =   735
         Left            =   7320
         Picture         =   "frmConsent1.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image CmdSearch 
         Height          =   735
         Left            =   9840
         Picture         =   "frmConsent1.frx":3A50
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image CmdAddNew 
         Height          =   735
         Left            =   0
         Picture         =   "frmConsent1.frx":772F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image CmdSave 
         Height          =   735
         Left            =   4920
         Picture         =   "frmConsent1.frx":ABF7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image CmdClear 
         Height          =   735
         Left            =   2280
         Picture         =   "frmConsent1.frx":E26E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   18255
      Begin VB.Frame Frame4 
         Caption         =   "For Women Donors : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3375
         Left            =   9600
         TabIndex        =   19
         Top             =   3720
         Width           =   8415
         Begin VB.Label Label13 
            Caption         =   "Would you like to be informed about any abnormal test result at the address furnished by you ?"
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
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   8055
         End
         Begin VB.Label Label12 
            Caption         =   "Do you have a child less than one year old ?"
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
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   5295
         End
         Begin VB.Label Label11 
            Caption         =   "Have you had an abortion in the last 6 months ?"
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
            TabIndex        =   21
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label10 
            Caption         =   "Are you pregnant ?"
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
            TabIndex        =   20
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3495
         Left            =   9600
         TabIndex        =   10
         Top             =   150
         Width           =   8415
         Begin VB.CheckBox Check7 
            Caption         =   "Dental Extraction"
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
            Left            =   5520
            TabIndex        =   18
            Top             =   3000
            Width           =   2535
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Ear Piercing"
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
            Left            =   2880
            TabIndex        =   17
            Top             =   3000
            Width           =   1815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Tattooing"
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
            TabIndex        =   16
            Top             =   3000
            Width           =   1575
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Continuous low grade fever"
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
            Left            =   4440
            TabIndex        =   14
            Top             =   840
            Width           =   3855
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Swollen Glands"
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
            Left            =   240
            TabIndex        =   13
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Repeated Diarrhoea"
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
            TabIndex        =   12
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Unexplained weight loss"
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
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label Label5 
            Caption         =   "In the last 6 months have you had any history"
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
            TabIndex        =   25
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label Label9 
            Caption         =   "In the last 6 months have you had any : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   2520
            Width           =   5415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Confidential (Tick wherever applicable)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4575
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   8655
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   615
            Left            =   5640
            TabIndex        =   32
            Top             =   3960
            Width           =   2895
            Begin VB.OptionButton Option8 
               Caption         =   "No"
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
               Left            =   1680
               TabIndex        =   38
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Yes"
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
               TabIndex        =   37
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   615
            Left            =   5640
            TabIndex        =   31
            Top             =   1920
            Width           =   2775
            Begin VB.OptionButton Option6 
               Caption         =   "No"
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
               Left            =   1680
               TabIndex        =   36
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Yes"
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
               TabIndex        =   35
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   855
            Left            =   5640
            TabIndex        =   30
            Top             =   1200
            Width           =   2775
            Begin VB.OptionButton Option4 
               Caption         =   "No"
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
               Left            =   1680
               TabIndex        =   34
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   615
            Left            =   5760
            TabIndex        =   27
            Top             =   360
            Width           =   2175
            Begin VB.OptionButton Option2 
               Caption         =   "No"
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
               Left            =   1440
               TabIndex        =   29
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Yes"
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
               Left            =   0
               TabIndex        =   28
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Do you have any reason to believe that you may be infected by either hepatittis ,Malaria , HIV/ AIDS and /or venereal disease ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   6
            Top             =   2520
            Width           =   8295
         End
         Begin VB.Label Label14 
            Caption         =   "Did you have any discomfort during / after donation "
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
            Left            =   600
            TabIndex        =   26
            Top             =   3480
            Width           =   7575
         End
         Begin VB.Label Label8 
            Caption         =   "Did you sleep well last night ?"
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
            Left            =   600
            TabIndex        =   9
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label Label7 
            Caption         =   "Have you eaten something in the  last 4 hours ?"
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
            TabIndex        =   8
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Do you feel well today ?"
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
            Left            =   600
            TabIndex        =   7
            Top             =   600
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   8655
         Begin VB.ComboBox cboVisitID 
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
            TabIndex        =   53
            Text            =   "Combo1"
            Top             =   1200
            Width           =   4215
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   360
            TabIndex        =   24
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label con_ID 
            Height          =   375
            Left            =   2400
            TabIndex        =   40
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label3 
            Caption         =   "Visit ID"
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
            Left            =   600
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Consent ID"
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
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   615
         Left            =   3960
         TabIndex        =   54
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Image Image3 
      Height          =   1215
      Index           =   1
      Left            =   0
      Picture         =   "frmConsent1.frx":11A45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmConsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim R As New ADODB.Recordset
Dim mMode As String   ' STARTUP / ADD / EDIT

' =========================================================
' FORM LOAD
' =========================================================
Private Sub Form_Load()
    On Error GoTo errh

    Call CONN

    Call Load_VisitIDs
    Call InitSearchSection
    Call Clear_form
    Call SetStartupMode

    Exit Sub

errh:
    'MsgBox "Error while loading Consent Form: " & Err.Description, vbCritical, "System Error"
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
    CmdClear.Enabled = False
    cmdSearch.Enabled = False

    Combo3.SetFocus
End Sub

Private Sub SetAddMode()
    mMode = "ADD"

    Call EnableEntryControls(True)
    'Call EnableSearchControls(False)

    CmdAddNew.Enabled = False
    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    CmdClear.Enabled = True
    'cmdSearch.Enabled = False

    cboVisitID.SetFocus
End Sub

Private Sub SetEditMode()
    mMode = "EDIT"

    Call EnableEntryControls(True)
    Call EnableSearchControls(True)

    CmdAddNew.Enabled = True
    CmdSave.Enabled = False
    CmdUpdate.Enabled = True
    CmdClear.Enabled = True
    cmdSearch.Enabled = True

    cboVisitID.SetFocus
End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    cboVisitID.Enabled = bFlag

    Option1.Enabled = bFlag
    Option2.Enabled = bFlag
    Option3.Enabled = bFlag
    Option4.Enabled = bFlag
    Option5.Enabled = bFlag
    Option6.Enabled = bFlag
    Option7.Enabled = bFlag
    Option8.Enabled = bFlag
    Option9.Enabled = bFlag
    Option10.Enabled = bFlag
    Option11.Enabled = bFlag
    Option12.Enabled = bFlag
    Option13.Enabled = bFlag
    Option14.Enabled = bFlag
    Option15.Enabled = bFlag
    Option16.Enabled = bFlag

    Check1.Enabled = bFlag
    Check2.Enabled = bFlag
    Check3.Enabled = bFlag
    Check4.Enabled = bFlag
    Check5.Enabled = bFlag
    Check6.Enabled = bFlag
    Check7.Enabled = bFlag
    Check8.Enabled = bFlag
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo3.Enabled = bFlag
    Combo4.Enabled = bFlag
End Sub

' =========================================================
' INITIALIZATION
' =========================================================
Private Sub InitSearchSection()
    Combo3.Clear
    Combo3.AddItem "Consent ID"
    Combo3.AddItem "Visit ID"

    Combo4.Clear
    Combo4.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub Load_VisitIDs()
    Dim rsVisit As New ADODB.Recordset

    Call CONN
    cboVisitID.Clear

    sql = "SELECT v_id FROM DONATION_VISIT ORDER BY TO_NUMBER(SUBSTR(v_id,4)) ASC"
    rsVisit.Open sql, C, adOpenStatic, adLockReadOnly

    Do While Not rsVisit.EOF
        cboVisitID.AddItem rsVisit!v_id
        rsVisit.MoveNext
    Loop

    rsVisit.Close
    Set rsVisit = Nothing
End Sub

' =========================================================
' AUTO CONSENT ID
' =========================================================
Private Sub Auto_ConsentID()
    Dim rsMax As New ADODB.Recordset
    Dim CC As String

    Call CONN
    CC = "CON"

    sql = "SELECT NVL(MAX(TO_NUMBER(SUBSTR(co_id,4))),0) FROM CONSENT"
    Set rsMax = C.Execute(sql)

    con_ID.Caption = CC & Format(rsMax.Fields(0).Value + 1, "0000")

    rsMax.Close
    Set rsMax = Nothing
End Sub

' =========================================================
' CLEAR FORM
' =========================================================
Private Sub Clear_form()

    con_ID.Caption = ""
    cboVisitID.ListIndex = -1
    cboVisitID.Text = ""

    ' Default No
    Option2.Value = True
    Option4.Value = True
    Option6.Value = True
    Option8.Value = True

    Option10.Value = True
    Option12.Value = True
    Option14.Value = True
    Option16.Value = True

    ' Checkboxes reset
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0

    Combo3.ListIndex = -1
    Combo3.Text = ""
    Combo4.Clear
    Combo4.Text = ""
    Combo4.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub cmdClear_Click()
    Call Clear_form
    If mMode = "ADD" Then
        Call Auto_ConsentID
        cboVisitID.SetFocus
    Else
        Call SetStartupMode
    End If
End Sub

' =========================================================
' ADD NEW
' =========================================================
Private Sub cmdAddNew_Click()
    Call Clear_form
    Call Auto_ConsentID
    Call Load_VisitIDs
    Call SetAddMode

    MsgBox "New Consent entry is ready.", vbInformation, "Ready"
End Sub

' =========================================================
' HELPERS
' =========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function ConsentExistsForVisit(ByVal vID As String, Optional ByVal excludeID As String = "") As Boolean
    Dim rsChk As New ADODB.Recordset

    sql = "SELECT COUNT(*) CNT FROM CONSENT WHERE v_id='" & Esc(vID) & "'"
    If Trim(excludeID) <> "" Then
        sql = sql & " AND co_id<>'" & Esc(excludeID) & "'"
    End If

    rsChk.Open sql, C, adOpenStatic, adLockReadOnly
    ConsentExistsForVisit = (rsChk!CNT > 0)
    rsChk.Close
    Set rsChk = Nothing
End Function

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Trim(con_ID.Caption) = "" Then
        MsgBox "Please click Add New first.", vbExclamation, "Validation Error"
        Exit Function
    End If

    If Trim(cboVisitID.Text) = "" Then
        MsgBox "Please select Visit ID.", vbExclamation, "Validation Error"
        cboVisitID.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

' =========================================================
' SAVE
' =========================================================
Private Sub cmdSave_Click()
    Dim vWell As String, vEat As String, vSleep As String, vInfect As String
    Dim vLast6 As String, vPreg As String, vAbor As String, vChild As String, vAbnor As String

    On Error GoTo errh

    Call CONN

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Invalid Action"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    If ConsentExistsForVisit(cboVisitID.Text) = True Then
        MsgBox "A consent record already exists for this Visit ID.", vbExclamation, "Duplicate Entry"
        cboVisitID.SetFocus
        Exit Sub
    End If

    vWell = IIf(Option1.Value = True, "Y", "N")
    vEat = IIf(Option3.Value = True, "Y", "N")
    vSleep = IIf(Option5.Value = True, "Y", "N")
    vInfect = IIf(Check8.Value = 1, "Y", "N")

    If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Or _
       Check4.Value = 1 Or Check5.Value = 1 Or Check6.Value = 1 Or _
       Check7.Value = 1 Then
        vLast6 = "Y"
    Else
        vLast6 = "N"
    End If

    vPreg = IIf(Option9.Value = True, "Y", "N")
    vAbor = IIf(Option11.Value = True, "Y", "N")
    vChild = IIf(Option13.Value = True, "Y", "N")
    vAbnor = IIf(Option15.Value = True, "Y", "N")

    sql = "INSERT INTO CONSENT (co_id, v_id, well, eat, sleep, infect, last_6, preg, abor, child, abnor) " & _
          "VALUES (" & _
          "'" & Esc(con_ID.Caption) & "', " & _
          "'" & Esc(cboVisitID.Text) & "', " & _
          "'" & vWell & "', " & _
          "'" & vEat & "', " & _
          "'" & vSleep & "', " & _
          "'" & vInfect & "', " & _
          "'" & vLast6 & "', " & _
          "'" & vPreg & "', " & _
          "'" & vAbor & "', " & _
          "'" & vChild & "', " & _
          "'" & vAbnor & "')"

    C.Execute sql

    MsgBox "Consent details saved successfully.", vbInformation, "Success"

    Call Clear_form
    Call SetStartupMode
    Exit Sub

errh:
    MsgBox "Save Error: " & Err.Description, vbCritical, "System Error"
End Sub

' =========================================================
' UPDATE
' =========================================================
Private Sub cmdUpdate_Click()
    Dim vWell As String, vEat As String, vSleep As String, vInfect As String
    Dim vLast6 As String, vPreg As String, vAbor As String, vChild As String, vAbnor As String

    On Error GoTo errh

    If mMode <> "EDIT" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If Trim(con_ID.Caption) = "" Then
        MsgBox "Please search for a record first.", vbExclamation, "Search Required"
        Exit Sub
    End If

    If ValidateForm = False Then Exit Sub

    Call CONN

    If ConsentExistsForVisit(cboVisitID.Text, con_ID.Caption) = True Then
        MsgBox "Another consent record already exists for this Visit ID.", vbExclamation, "Duplicate Entry"
        cboVisitID.SetFocus
        Exit Sub
    End If

    vWell = IIf(Option1.Value = True, "Y", "N")
    vEat = IIf(Option3.Value = True, "Y", "N")
    vSleep = IIf(Option5.Value = True, "Y", "N")
    vInfect = IIf(Check8.Value = 1, "Y", "N")

    If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Or _
       Check4.Value = 1 Or Check5.Value = 1 Or Check6.Value = 1 Or _
       Check7.Value = 1 Then
        vLast6 = "Y"
    Else
        vLast6 = "N"
    End If

    vPreg = IIf(Option9.Value = True, "Y", "N")
    vAbor = IIf(Option11.Value = True, "Y", "N")
    vChild = IIf(Option13.Value = True, "Y", "N")
    vAbnor = IIf(Option15.Value = True, "Y", "N")

    If MsgBox("Do you want to update this consent record?", vbQuestion + vbYesNo, "Confirm Update") = vbNo Then Exit Sub

    sql = "UPDATE CONSENT SET " & _
          "v_id='" & Esc(cboVisitID.Text) & "', " & _
          "well='" & vWell & "', " & _
          "eat='" & vEat & "', " & _
          "sleep='" & vSleep & "', " & _
          "infect='" & vInfect & "', " & _
          "last_6='" & vLast6 & "', " & _
          "preg='" & vPreg & "', " & _
          "abor='" & vAbor & "', " & _
          "child='" & vChild & "', " & _
          "abnor='" & vAbnor & "' " & _
          "WHERE co_id='" & Esc(con_ID.Caption) & "'"

    C.Execute sql

    MsgBox "Consent details updated successfully.", vbInformation, "Success"
    Call Clear_form
    Call SetStartupMode
    Exit Sub

errh:
    MsgBox "Update Error: " & Err.Description, vbCritical, "System Error"
End Sub

' =========================================================
' SEARCH SECTION
' =========================================================
Private Sub Combo3_Click()
    Dim rsS As New ADODB.Recordset

    If Combo3.Text = "" Then Exit Sub

    Combo4.Clear
    Combo4.Enabled = True
    cmdSearch.Enabled = False

    Call CONN

    If Combo3.Text = "Consent ID" Then
        sql = "SELECT co_id FROM CONSENT ORDER BY TO_NUMBER(SUBSTR(co_id,4)) ASC"
    ElseIf Combo3.Text = "Visit ID" Then
        sql = "SELECT v_id FROM CONSENT ORDER BY TO_NUMBER(SUBSTR(v_id,4)) ASC"
    End If

    rsS.Open sql, C, adOpenStatic, adLockReadOnly
    Do While Not rsS.EOF
        Combo4.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop
    rsS.Close
    Set rsS = Nothing
End Sub

Private Sub Combo4_Click()
    If Trim(Combo4.Text) <> "" Then
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim rsSearch As New ADODB.Recordset

    If Combo3.Text = "" Or Combo4.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Error"
        Exit Sub
    End If

    Call CONN

    If Combo3.Text = "Consent ID" Then
        sql = "SELECT * FROM CONSENT WHERE co_id='" & Esc(Combo4.Text) & "'"
    ElseIf Combo3.Text = "Visit ID" Then
        sql = "SELECT * FROM CONSENT WHERE v_id='" & Esc(Combo4.Text) & "'"
    End If

    rsSearch.Open sql, C, adOpenStatic, adLockReadOnly

    If rsSearch.EOF Then
        MsgBox "Record not found.", vbExclamation, "Search Result"
        rsSearch.Close
        Set rsSearch = Nothing
        Exit Sub
    End If

    con_ID.Caption = "" & rsSearch!co_id

    cboVisitID.Clear
    cboVisitID.AddItem rsSearch!v_id
    cboVisitID.Text = rsSearch!v_id

    Option1.Value = (rsSearch!well = "Y")
    Option2.Value = (rsSearch!well <> "Y")

    Option3.Value = (rsSearch!eat = "Y")
    Option4.Value = (rsSearch!eat <> "Y")

    Option5.Value = (rsSearch!sleep = "Y")
    Option6.Value = (rsSearch!sleep <> "Y")

    Check8.Value = IIf(rsSearch!infect = "Y", 1, 0)

    If rsSearch!last_6 = "Y" Then
        ' exact checkbox details table me nahi hain, isliye sab blank hi rahenge
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
    Else
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
    End If

    Option9.Value = (rsSearch!preg = "Y")
    Option10.Value = (rsSearch!preg <> "Y")

    Option11.Value = (rsSearch!abor = "Y")
    Option12.Value = (rsSearch!abor <> "Y")

    Option13.Value = (rsSearch!child = "Y")
    Option14.Value = (rsSearch!child <> "Y")

    Option15.Value = (rsSearch!abnor = "Y")
    Option16.Value = (rsSearch!abnor <> "Y")

    rsSearch.Close
    Set rsSearch = Nothing

    MsgBox "Record found successfully.", vbInformation, "Search Success"
    Call SetEditMode
End Sub

