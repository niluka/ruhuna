VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCancellationOfRepaymentsReports 
   Caption         =   "Cancellation of Refunds - Reports"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   6015
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Selected Day"
      TabPicture(0)   =   "frmCancellationOfRepaymentsReports.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MonthView1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Period"
      TabPicture(1)   =   "frmCancellationOfRepaymentsReports.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MonthView3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MonthView2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   56426497
         CurrentDate     =   39534
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   56426497
         CurrentDate     =   39534
      End
      Begin MSComCtl2.MonthView MonthView3 
         Height          =   2370
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   56426497
         CurrentDate     =   39534
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCancellationOfRepaymentsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

