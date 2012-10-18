VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStoresPreferance 
   Caption         =   "Store Preferances"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramePastDurations 
      Caption         =   "Default Past Duration"
      Height          =   1815
      Left            =   3600
      TabIndex        =   18
      Top             =   120
      Width           =   4095
      Begin VB.TextBox PriceDays 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1560
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox OrderingDays 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1560
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox UsageDays 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Prices for                            days"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "Ordering for                        days"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Usage for                            days"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame FrameExpected 
      Caption         =   "Expected Period of Ordering"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   3375
      Begin MSComCtl2.DTPicker dtpStime 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20774914
         CurrentDate     =   0.333333333333333
      End
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpETime 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20774914
         CurrentDate     =   0.666666666666667
      End
      Begin VB.Label Label4 
         Caption         =   "Ending Time"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Starting Time"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Up to                                  days"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Starting From                      days"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame FrameExpireTransfer 
      Caption         =   "Transfer"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
      Begin VB.OptionButton OptDoNotAllowExpireTransfer 
         Caption         =   "No not allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton OptAllowExpireTransfer 
         Caption         =   "Allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Frame FrameConsumptionExpire 
      Caption         =   "Consumption"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
      Begin VB.OptionButton OptDoNotAllowExpireConsumption 
         Caption         =   "No not allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton OptAllowExpireConsumption 
         Caption         =   "Allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Frame FrameSaleExpiary 
      Caption         =   "Sale"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton OptAllowExpireSale 
         Caption         =   "Allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optDoNotAllowExpireSale 
         Caption         =   "No not allow Expired items"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmStoresPreferance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetMemory()
    OptDoNotAllowExpireConsumption.Value = DoNotAllowExpireConsumption
    optDoNotAllowExpireSale.Value = DoNotAllowExpireSale
    OptDoNotAllowExpireTransfer.Value = DoNotAllowExpireTransfer
    dtpETime.Value = GetSetting(App.EXEName, "Options", "dtpETime", "16:00")
    dtpStime.Value = GetSetting(App.EXEName, "Options", "dtpStime", "16:00")
    txtFromDate.Text = GetSetting(App.EXEName, "Options", "txtFromDate", 3)
    txtToDate.Text = GetSetting(App.EXEName, "Options", "txtToDate", 5)
    UsageDays.Text = GetSetting(App.EXEName, "Options", "UsageDays", 30)
    OrderingDays.Text = GetSetting(App.EXEName, "Options", "OrderingDays", 30)
    PriceDays.Text = GetSetting(App.EXEName, "Options", "PriceDays", 30)
    
End Sub

Private Sub SetReg()
    SaveSetting App.EXEName, "Options", "DoNotAllowExpireConsumption", OptDoNotAllowExpireConsumption.Value
    SaveSetting App.EXEName, "Options", "DoNotAllowExpireSale", optDoNotAllowExpireSale.Value
    SaveSetting App.EXEName, "Options", "DoNotAllowExpireTransfer", OptDoNotAllowExpireTransfer.Value
    SaveSetting App.EXEName, "Options", "dtpETime", dtpETime.Value
    SaveSetting App.EXEName, "Options", "dtpStime", dtpStime.Value
    SaveSetting App.EXEName, "Options", "txtFromDate", Val(txtFromDate.Text)
    SaveSetting App.EXEName, "Options", "txtToDate", Val(txtToDate.Text)
    SaveSetting App.EXEName, "Options", "UsageDays", Val(UsageDays.Text)
    SaveSetting App.EXEName, "Options", "OrderingDays", Val(OrderingDays.Text)
    SaveSetting App.EXEName, "Options", "PriceDays", Val(PriceDays.Text)
End Sub

Private Sub SetMemory()
    DoNotAllowExpireConsumption = OptDoNotAllowExpireConsumption.Value
    DoNotAllowExpireSale = optDoNotAllowExpireSale.Value
    DoNotAllowExpireTransfer = OptDoNotAllowExpireTransfer.Value
End Sub

Private Sub Form_Load()
    Call GetMemory
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetMemory
    Call SetReg
End Sub
