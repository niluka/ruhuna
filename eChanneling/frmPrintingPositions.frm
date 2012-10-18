VERSION 5.00
Begin VB.Form frmPrintingPositions 
   Caption         =   "Printing Positions"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   13470
   WindowState     =   2  'Maximized
   Begin VB.Frame FramePrintingPositions 
      BackColor       =   &H80000009&
      Caption         =   "Printing Position Arrangements"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.TextBox txtPhone2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5880
         TabIndex        =   34
         Text            =   "Phone 2"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtPhone1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   720
         TabIndex        =   33
         Text            =   "Phone 1"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtTime2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   7200
         TabIndex        =   32
         Text            =   "Time 2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtTime1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   1920
         TabIndex        =   31
         Text            =   "Time 1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtMsg2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   6960
         TabIndex        =   30
         Text            =   "Message2"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtMsg1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   1800
         TabIndex        =   29
         Text            =   "Message1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtRoomNo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3000
         TabIndex        =   28
         Text            =   "Room No1"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtRoomNo2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8280
         TabIndex        =   27
         Text            =   "Room No2"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtAgentRefNo2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8040
         TabIndex        =   26
         Text            =   "Agent Ref. No 1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtAgentCode2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5880
         TabIndex        =   25
         Text            =   "Agent Code 1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtAgentRefNo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   2880
         TabIndex        =   24
         Text            =   "Agent Ref. No 1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtAgentCode1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   720
         TabIndex        =   23
         Text            =   "Agent Code 1"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtReceptionist2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8040
         TabIndex        =   22
         Text            =   "Receptionist 2"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5880
         TabIndex        =   21
         Text            =   "Total 2"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtHospChg2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8280
         TabIndex        =   20
         Text            =   "Hosp. Chg. 2"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtDrsFee2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5760
         TabIndex        =   19
         Text            =   "Dr's Fee 2"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtAt2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8280
         TabIndex        =   18
         Text            =   "At 2"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtAppointOn2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5760
         TabIndex        =   17
         Text            =   "Appoint. On 2"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtPatient2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5760
         TabIndex        =   16
         Text            =   "Patient 2"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtAppoNo2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8280
         TabIndex        =   15
         Text            =   "Appo. No. 2"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtConsultant2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5760
         TabIndex        =   14
         Text            =   "Consultant 2"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtRefNo2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   8280
         TabIndex        =   13
         Text            =   "Ref. No. 2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDate2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   5760
         TabIndex        =   12
         Text            =   "Date 2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtReceptionist1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   2880
         TabIndex        =   11
         Text            =   "Receptionist 1"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   720
         TabIndex        =   10
         Text            =   "Total 1"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtHospChg1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3000
         TabIndex        =   9
         Text            =   "Hosp. Chg. 1"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtDrsFee1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   480
         TabIndex        =   8
         Text            =   "Dr's Fee 1"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtAt1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3000
         TabIndex        =   7
         Text            =   "At 1"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtAppointOn1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   480
         TabIndex        =   6
         Text            =   "Appoint. On 1"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtPatient1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   480
         TabIndex        =   5
         Text            =   "Patient 1"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtAppoNo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3000
         TabIndex        =   4
         Text            =   "Appo. No. 1"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtConsultant1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   480
         TabIndex        =   3
         Text            =   "Consultant 1"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtRefNo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3000
         TabIndex        =   2
         Text            =   "Ref. No. 1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDate1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   480
         TabIndex        =   1
         Text            =   "Date1"
         Top             =   2160
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPrintingPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MoveConstant As Integer


Private Sub Form_Load()
    Call GetPreferances
    MoveConstant = 20
End Sub

Private Sub GetPreferances()
With DataEnvironment1.rssqlTem15
    If .State = 1 Then .Open
    .Source = "SELECT * from tblchannellingPrintingPreferances"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    
    txtDate1.Top = !date1y * FramePrintingPositions.Height
    txtDate1.Left = !date1x * FramePrintingPositions.Width
    txtRefNo1.Top = !refno1y * FramePrintingPositions.Height
    txtRefNo1.Left = !refno1x * FramePrintingPositions.Width
    txtAppoNo1.Top = !appono1y * FramePrintingPositions.Height
    txtAppoNo1.Left = !appono1x * FramePrintingPositions.Width
    txtConsultant1.Top = !consultant1y * FramePrintingPositions.Height
    txtConsultant1.Left = !consultant1x * FramePrintingPositions.Width
    txtPatient1.Top = !patient1y * FramePrintingPositions.Height
    txtPatient1.Left = !patient1x * FramePrintingPositions.Width
    txtPhone1.Top = !phone1y * FramePrintingPositions.Height
    txtPhone1.Left = !phone1x * FramePrintingPositions.Width
    
    txtAppointOn1.Top = !appointon1y * FramePrintingPositions.Height
    txtAppointOn1.Left = !appointon1x * FramePrintingPositions.Width
    txtAt1.Top = !at1y * FramePrintingPositions.Height
    txtAt1.Left = !at1x * FramePrintingPositions.Width
    txtDrsFee1.Top = !drsfee1y * FramePrintingPositions.Height
    txtDrsFee1.Left = !drsfee1x * FramePrintingPositions.Width
    txtHospChg1.Top = !hospchg1y * FramePrintingPositions.Height
    txtHospChg1.Left = !hospchg1x * FramePrintingPositions.Width
    txtTotal1.Top = !total1y * FramePrintingPositions.Height
    txtTotal1.Left = !total1x * FramePrintingPositions.Width
    txtReceptionist1.Top = !receptionist1y * FramePrintingPositions.Height
    txtReceptionist1.Left = !receptionist1x * FramePrintingPositions.Width
    txtAgentCode1.Top = !agentcode1y * FramePrintingPositions.Height
    txtAgentCode1.Left = !agentcode1x * FramePrintingPositions.Width
    txtAgentRefNo1.Top = !agentrefno1y * FramePrintingPositions.Height
    txtAgentRefNo1.Left = !agentrefno1x * FramePrintingPositions.Width
    txtRoomNo1.Top = !roomno1y * FramePrintingPositions.Height
    txtRoomNo1.Left = !roomno1x * FramePrintingPositions.Width
    txtMsg1.Top = !msg1y * FramePrintingPositions.Height
    txtMsg1.Left = !msg1x * FramePrintingPositions.Width
    txtTime1.Top = !time1y * FramePrintingPositions.Height
    txtTime1.Left = !time1X * FramePrintingPositions.Width

    txtDate2.Top = !date2y * FramePrintingPositions.Height
    txtDate2.Left = !date2x * FramePrintingPositions.Width
    txtRefNo2.Top = !refno2y * FramePrintingPositions.Height
    txtRefNo2.Left = !refno2x * FramePrintingPositions.Width
    txtAppoNo2.Top = !appono2y * FramePrintingPositions.Height
    txtAppoNo2.Left = !appono2x * FramePrintingPositions.Width
    txtConsultant2.Top = !consultant2y * FramePrintingPositions.Height
    txtConsultant2.Left = !consultant2x * FramePrintingPositions.Width
    txtPatient2.Top = !patient2y * FramePrintingPositions.Height
    txtPatient2.Left = !patient2x * FramePrintingPositions.Width
    txtPhone2.Top = !phone2y * FramePrintingPositions.Height
    txtPhone2.Left = !phone2x * FramePrintingPositions.Width
    
    txtAppointOn2.Top = !appointon2y * FramePrintingPositions.Height
    txtAppointOn2.Left = !appointon2x * FramePrintingPositions.Width
    txtAt2.Top = !at2y * FramePrintingPositions.Height
    txtAt2.Left = !at2x * FramePrintingPositions.Width
    txtDrsFee2.Top = !drsfee2y * FramePrintingPositions.Height
    txtDrsFee2.Left = !drsfee2x * FramePrintingPositions.Width
    txtHospChg2.Top = !hospchg2y * FramePrintingPositions.Height
    txtHospChg2.Left = !hospchg2x * FramePrintingPositions.Width
    txtTotal2.Top = !total2y * FramePrintingPositions.Height
    txtTotal2.Left = !total2x * FramePrintingPositions.Width
    txtReceptionist2.Top = !receptionist2y * FramePrintingPositions.Height
    txtReceptionist2.Left = !receptionist2x * FramePrintingPositions.Width
    txtAgentCode2.Top = !agentcode2y * FramePrintingPositions.Height
    txtAgentCode2.Left = !agentcode2x * FramePrintingPositions.Width
    txtAgentRefNo2.Top = !agentrefno2y * FramePrintingPositions.Height
    txtAgentRefNo2.Left = !agentrefno2x * FramePrintingPositions.Width
    txtRoomNo2.Top = !roomno2y * FramePrintingPositions.Height
    txtRoomNo2.Left = !roomno2x * FramePrintingPositions.Width
    txtMsg2.Top = !msg2y * FramePrintingPositions.Height
    txtMsg2.Left = !msg2x * FramePrintingPositions.Width
    txtTime2.Top = !time2y * FramePrintingPositions.Height
    txtTime2.Left = !time2X * FramePrintingPositions.Width


    .Close
End With
End Sub
Private Sub SavePositions()
With DataEnvironment1.rssqlTem15
    If .State = 1 Then .Open
    .Source = "SELECT * from tblchannellingPrintingPreferances"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    !date1y = txtDate1.Top / FramePrintingPositions.Height
    !date1x = txtDate1.Left / FramePrintingPositions.Width
    !refno1y = txtRefNo1.Top / FramePrintingPositions.Height
    !refno1x = txtRefNo1.Left / FramePrintingPositions.Width
    !appono1y = txtAppoNo1.Top / FramePrintingPositions.Height
    !appono1x = txtAppoNo1.Left / FramePrintingPositions.Width
    !consultant1y = txtConsultant1.Top / FramePrintingPositions.Height
    !consultant1x = txtConsultant1.Left / FramePrintingPositions.Width
    !patient1y = txtPatient1.Top / FramePrintingPositions.Height
    !patient1x = txtPatient1.Left / FramePrintingPositions.Width
    !phone1y = txtPhone1.Top / FramePrintingPositions.Height
    !phone1x = txtPhone1.Left / FramePrintingPositions.Width
    
    
    
    !appointon1y = txtAppointOn1.Top / FramePrintingPositions.Height
    !appointon1x = txtAppointOn1.Left / FramePrintingPositions.Width
    !at1y = txtAt1.Top / FramePrintingPositions.Height
    !at1x = txtAt1.Left / FramePrintingPositions.Width
    !drsfee1y = txtDrsFee1.Top / FramePrintingPositions.Height
    !drsfee1x = txtDrsFee1.Left / FramePrintingPositions.Width
    !hospchg1y = txtHospChg1.Top / FramePrintingPositions.Height
    !hospchg1x = txtHospChg1.Left / FramePrintingPositions.Width
    !total1y = txtTotal1.Top / FramePrintingPositions.Height
    !total1x = txtTotal1.Left / FramePrintingPositions.Width
    !receptionist1y = txtReceptionist1.Top / FramePrintingPositions.Height
    !receptionist1x = txtReceptionist1.Left / FramePrintingPositions.Width
    !agentcode1y = txtAgentCode1.Top / FramePrintingPositions.Height
    !agentcode1x = txtAgentCode1.Left / FramePrintingPositions.Width
    !agentrefno1y = txtAgentRefNo1.Top / FramePrintingPositions.Height
    !agentrefno1x = txtAgentRefNo1.Left / FramePrintingPositions.Width
    !roomno1y = txtRoomNo1.Top / FramePrintingPositions.Height
    !roomno1x = txtRoomNo1.Left / FramePrintingPositions.Width
    !msg1y = txtMsg1.Top / FramePrintingPositions.Height
    !msg1x = txtMsg1.Left / FramePrintingPositions.Width
    !time1y = txtTime1.Top / FramePrintingPositions.Height
    !time1X = txtTime1.Left / FramePrintingPositions.Width
    
    
    !date2y = txtDate2.Top / FramePrintingPositions.Height
    !date2x = txtDate2.Left / FramePrintingPositions.Width
    !refno2y = txtRefNo2.Top / FramePrintingPositions.Height
    !refno2x = txtRefNo2.Left / FramePrintingPositions.Width
    !appono2y = txtAppoNo2.Top / FramePrintingPositions.Height
    !appono2x = txtAppoNo2.Left / FramePrintingPositions.Width
    !consultant2y = txtConsultant2.Top / FramePrintingPositions.Height
    !consultant2x = txtConsultant2.Left / FramePrintingPositions.Width
    !patient2y = txtPatient2.Top / FramePrintingPositions.Height
    !patient2x = txtPatient2.Left / FramePrintingPositions.Width
    !phone2y = txtPhone2.Top / FramePrintingPositions.Height
    !phone2x = txtPhone2.Left / FramePrintingPositions.Width
    
    !appointon2y = txtAppointOn2.Top / FramePrintingPositions.Height
    !appointon2x = txtAppointOn2.Left / FramePrintingPositions.Width
    !at2y = txtAt2.Top / FramePrintingPositions.Height
    !at2x = txtAt2.Left / FramePrintingPositions.Width
    !drsfee2y = txtDrsFee2.Top / FramePrintingPositions.Height
    !drsfee2x = txtDrsFee2.Left / FramePrintingPositions.Width
    !hospchg2y = txtHospChg2.Top / FramePrintingPositions.Height
    !hospchg2x = txtHospChg2.Left / FramePrintingPositions.Width
    !total2y = txtTotal2.Top / FramePrintingPositions.Height
    !total2x = txtTotal2.Left / FramePrintingPositions.Width
    !receptionist2y = txtReceptionist2.Top / FramePrintingPositions.Height
    !receptionist2x = txtReceptionist2.Left / FramePrintingPositions.Width
    !agentcode2y = txtAgentCode2.Top / FramePrintingPositions.Height
    !agentcode2x = txtAgentCode2.Left / FramePrintingPositions.Width
    !agentrefno2y = txtAgentRefNo2.Top / FramePrintingPositions.Height
    !agentrefno2x = txtAgentRefNo2.Left / FramePrintingPositions.Width
    !roomno2y = txtRoomNo2.Top / FramePrintingPositions.Height
    !roomno2x = txtRoomNo2.Left / FramePrintingPositions.Width
    !msg2y = txtMsg2.Top / FramePrintingPositions.Height
    !msg2x = txtMsg2.Left / FramePrintingPositions.Width
    !time2y = txtTime2.Top / FramePrintingPositions.Height
    !time2X = txtTime2.Left / FramePrintingPositions.Width
    
    .Update
    .Close
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePositions
End Sub


Private Sub txtAppointOn1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAppointOn1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txttime1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtTime1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txtmsg1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtMsg1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txttime2_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtTime2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txtmsg2_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtMsg2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txtroomno1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtRoomNo1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub


Private Sub txtroomno2_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtRoomNo2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub


Private Sub txtAppointOn2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAppointOn2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtAppoNo1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAppoNo1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtAppoNo2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAppoNo2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtAt1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAt1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtAt2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAt2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtConsultant1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtConsultant1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtConsultant2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtConsultant2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtDate1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtDate2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtDrsFee1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtDrsFee1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtDrsFee2_KeyDown(KeyCode As Integer, Shift As Integer)

   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtDrsFee2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtHospChg1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtHospChg1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtHospChg2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtHospChg2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtPatient1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtPatient1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtPatient2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtPatient2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub









Private Sub txtPhone1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtPhone1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtPhone2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtPhone2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub



Private Sub txtReceptionist1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtReceptionist1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtReceptionist2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtReceptionist2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtRefNo1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtRefNo1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtRefNo2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtRefNo2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtTotal1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtTotal1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

Private Sub txtTotal2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtTotal2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub




Private Sub txtagentcode1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAgentCode1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub



Private Sub txtagentcode2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAgentCode2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub


Private Sub txtagentrefno1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAgentRefNo1
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub


Private Sub txtagentrefno2_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   Dim ThisTextBox As TextBox
   Set ThisTextBox = txtAgentRefNo2
    Select Case KeyCode
        Case vbKeyUp
            ThisTextBox.Top = ThisTextBox.Top - MoveConstant
        Case vbKeyDown
            ThisTextBox.Top = ThisTextBox.Top + MoveConstant
        Case vbKeyLeft
            ThisTextBox.Left = ThisTextBox.Left - MoveConstant
        Case vbKeyRight
            ThisTextBox.Left = ThisTextBox.Left + MoveConstant
    End Select
    KeyCode = Empty

End Sub

