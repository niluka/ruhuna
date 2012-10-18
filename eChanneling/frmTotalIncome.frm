VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTotalIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Total Income"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTotalIncome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10335
   Begin btButtonEx.ButtonEx bttnPrintReport 
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   8520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   9855
      Begin VB.Label lblProfitLoss 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   45
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Profit (Loss)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   9720
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label lblIncome 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   40
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label lblExpences 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   39
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label lblPatientRepayments 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   38
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label lblCashPurchases 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   37
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label lblCashPayments 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblCustomerPaymentCheque 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   35
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblAgentPaymentCheque 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   34
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblReceptionCheque 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   33
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblCustomerPaymentCard 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   32
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblAgentPaymentCard 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   31
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblReceptionCard 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   30
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblCustomerPaymentCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblAgentPaymentCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblReceptionCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   7560
         X2              =   7560
         Y1              =   240
         Y2              =   6600
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   5400
         Y1              =   240
         Y2              =   7080
      End
      Begin VB.Line Line8 
         X1              =   5400
         X2              =   9720
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Repayments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchases"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Income"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   6975
         Left            =   120
         Top             =   240
         Width           =   9615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Reception"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Reception"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Reception"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expences"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   4560
         Width           =   1335
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Close"
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
      Height          =   8295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   14631
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "frmTotalIncome.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DTPickerDate"
      Tab(0).Control(1)=   "Label1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmTotalIncome.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPickerSelectDate"
      Tab(1).Control(1)=   "Label36"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Selected Period"
      TabPicture(2)   =   "frmTotalIncome.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label37"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label38"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPickerSelectTo"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPickerSelectFrom"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPickerDate 
         Height          =   375
         Left            =   -70800
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MM yyyy"
         Format          =   58130435
         CurrentDate     =   39425
      End
      Begin MSComCtl2.DTPicker DTPickerSelectDate 
         Height          =   375
         Left            =   -70320
         TabIndex        =   3
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MM yyyy"
         Format          =   58130435
         CurrentDate     =   39425
      End
      Begin MSComCtl2.DTPicker DTPickerSelectFrom 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MM yyyy"
         Format          =   58130435
         CurrentDate     =   39425
      End
      Begin MSComCtl2.DTPicker DTPickerSelectTo 
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MM yyyy"
         Format          =   58130435
         CurrentDate     =   39425
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   41
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71760
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmTotalIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemReceptionCash As Double
Dim TemAgentPaymentCash As Double
Dim TemCustomerPaymentCash As Double
Dim TemReceptionCard As Double
Dim TemAgentPaymentCard As Double
Dim TemCustomerPaymentCard As Double
Dim TemReceptionCheque As Double
Dim TemCustomerPaymentChaeque As Double
Dim TemAgentPaymentCheque As Double

Dim TemPaymentCash As Double
Dim TemPurchasesCash As Double
Dim TemDealerPaymentCash As Double
Dim TemBankDepositCash As Double

Dim TemStaffPayment As Double

Dim TotalIncome As Double
Dim TotalExpence As Double
Private Sub SetColour()



Select Case ColourScheme

Case 1:

BttnBackColour = 5341695
BttnForeColour = 1314458
FrmBackColour = 11066623
FrmForeColour = 1314458
FrameBackColour = 11066623
FrameForeColour = 1314458
TxtBackColour = 9881851
TxtForeColour = 1314458
LblBackColour = 11066623
LblForeColour = 1314458



GridBackColor = 9881855
GridBackColorBkg = 10474239
GridBackColorFixed = 8566015
GridBackColorSel = 5341695

GridForeColor = 1314458
GridForeColorFixed = 11944
GridForeColorSel = 3014824

'GridCellBackColor = 5853695
'GridCellForeColor = 658120


Case 2:

BttnBackColour = 14803300
BttnForeColour = 5539362
FrmBackColour = 16766120
FrmForeColour = 5539362
FrameBackColour = 16766120
FrameForeColour = 5539362
TxtBackColour = 16760450
TxtForeColour = 5539362
LblBackColour = 16766120
LblForeColour = 5539362

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588


Case 3:

BttnBackColour = 51455
BttnForeColour = 942490
FrmBackColour = 11070719
FrmForeColour = 942490
FrameBackColour = 11070719
FrameForeColour = 942490
TxtBackColour = 11528439
TxtForeColour = 1314458
LblBackColour = 11070719
LblForeColour = 942490

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588

End Select

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour
'
bttnPrintReport.BackColor = BttnBackColour
bttnPrintReport.ForeColor = BttnForeColour
frmTotalIncome.BackColor = FrameBackColour
frmTotalIncome.ForeColor = FrameForeColour
'
SSTab1.BackColor = FrameBackColour
SSTab1.ForeColor = FrameForeColour
End Sub

Private Sub CalculateTotals()
    TotalIncome = 0
    Call FindReception
    Call FindPatientCreditSettle
    Call FindAgentCreditSettle
    lblIncome.Caption = Format(TotalIncome, "#0.00")
    
    TotalExpence = 0
    
    Call FindPayment
    Call FindPatientRepayments
    
    lblExpences.Caption = Format(TotalExpence, "0.00")

    If TotalExpence > TotalIncome Then
        lblProfitLoss = "(" & Format(TotalExpence - TotalIncome, "0.00") & ")"
    Else
        lblProfitLoss = Format(TotalIncome - TotalExpence, "0.00")
    End If
    
End Sub


Private Sub FindPatientRepayments()
Dim TemPatientRepay As Double
With DataEnvironment1.rssqlTem6
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
        Case 0: .Source = "SELECT * from tblpatientrepay where (repayDate = #" & Date & "#)"
        Case 1: .Source = "SELECT * from tblpatientrepay where (repayDate = #" & DTPickerSelectDate.Value & "#)"
        Case 2: .Source = "SELECT * from tblpatientrepay where (repayDate between #" & DTPickerSelectFrom.Value & "# and #" & DTPickerSelectTo.Value & "#)"
    End Select
    If .State = 0 Then .Open
    TemPatientRepay = 0
    If .RecordCount <> 0 Then
    .MoveFirst
        While .EOF = False
            TemPatientRepay = TemPatientRepay + !TotalRepay
            .MoveNext
        Wend
    End If
    .Close
End With

    lblPatientRepayments.Caption = Format(TemPatientRepay, "#0.00")
    
    TotalExpence = TotalExpence + TemPatientRepay

End Sub

Private Sub FindPayment()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
        Case 0: .Source = "SELECT * from tblstaffpayment where (PaidDate = #" & Date & "#)"
        Case 1: .Source = "SELECT * from tblstaffpayment where (PaidDate = #" & DTPickerSelectDate.Value & "#)"
        Case 2: .Source = "SELECT * from tblstaffpayment where (PaidDate between #" & DTPickerSelectFrom.Value & "# and #" & DTPickerSelectTo.Value & "#)"
    End Select
    If .State = 0 Then .Open
    TemStaffPayment = 0
    If .RecordCount <> 0 Then
    .MoveFirst
        While .EOF = False
            TemStaffPayment = TemStaffPayment + !PaidAmount
            .MoveNext
        Wend
    End If
    .Close
End With

    lblCashPayments.Caption = Format(TemStaffPayment, "0.00")
    
    TotalExpence = TemStaffPayment
End Sub

Private Sub FindAgentCreditSettle()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        
        Select Case SSTab1.Tab
            Case 0: .Source = "SELECT tblAgentCashSettle.* from tblAgentCashSettle where (SettledDate = #" & Date & "#) "
            Case 1: .Source = "SELECT tblAgentCashSettle.* from tblAgentCashSettle where (settleddate = #" & DTPickerSelectDate.Value & "#)"
            Case 2: .Source = "SELECT tblAgentCashSettle.* from tblAgentCashSettle where (SettledDate between #" & DTPickerSelectFrom.Value & "# and #" & DTPickerSelectTo.Value & "#)"
        End Select
        
        If .State = 0 Then .Open
        
        TemAgentPaymentCash = 0
        TemAgentPaymentCard = 0
        TemAgentPaymentCheque = 0
        
        If .RecordCount <> 0 Then
            .MoveFirst
            While .EOF = False
                Select Case !SettleMethod
                    Case "Cash": TemAgentPaymentCash = TemAgentPaymentCash + !Cash
                    Case "CreditCard": TemAgentPaymentCard = TemAgentPaymentCard + !CreditCardAmount
                    Case "Cheque": TemAgentPaymentCheque = TemAgentPaymentCheque + !ChequeAmount
                End Select
                .MoveNext
            Wend
        End If
    End With
    
    lblAgentPaymentCash.Caption = Format(TemAgentPaymentCash, "#0.00")
    lblAgentPaymentCard.Caption = Format(TemAgentPaymentCard, "#0.00")
    lblAgentPaymentCheque.Caption = Format(TemAgentPaymentCheque, "#0.00")
    
    TotalIncome = TotalIncome + TemAgentPaymentCash + TemAgentPaymentCard + TemAgentPaymentCheque
   
    
End Sub



Private Sub FindPatientCreditSettle()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        
        Select Case SSTab1.Tab
            Case 0: .Source = "SELECT tblPatientCashSettle.* from tblPatientCashSettle where (SettledDate = #" & Date & "#)"
            Case 1: .Source = "SELECT tblPatientCashSettle.* from tblPatientCashSettle where (settleddate = #" & DTPickerSelectDate.Value & "#)"
            Case 2: .Source = "SELECT tblPatientCashSettle.* from tblPatientCashSettle where (SettledDate between #" & DTPickerSelectFrom.Value & "# and #" & DTPickerSelectTo.Value & "#)"
        End Select
        
        If .State = 0 Then .Open
        
        TemCustomerPaymentCash = 0
        TemCustomerPaymentCard = 0
        TemCustomerPaymentChaeque = 0
        
        If .RecordCount <> 0 Then
            .MoveFirst
            While .EOF = False
                Select Case !SettleMethod
                    Case "Cash": TemCustomerPaymentCash = TemCustomerPaymentCash + !Cash
                    Case "CreditCard": TemCustomerPaymentCard = TemCustomerPaymentCard + !CreditCardAmount
                    Case "Cheque": TemCustomerPaymentChaeque = TemCustomerPaymentChaeque + !ChequeAmount
                End Select
                .MoveNext
            Wend
        End If
    End With
    
    lblCustomerPaymentCash.Caption = Format(TemCustomerPaymentCash, "#0.00")
    lblCustomerPaymentCard.Caption = Format(TemCustomerPaymentCard, "#0.00")
    lblCustomerPaymentCheque.Caption = Format(TemCustomerPaymentChaeque, "#0.00")
    
    TotalIncome = TotalIncome + TemCustomerPaymentCash + TemCustomerPaymentCard + TemCustomerPaymentChaeque
        
End Sub


Private Sub FindReception()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
        Case 0: .Source = "SELECT * from tblpatientbill where (Date = #" & Date & "#) and billsuccess = true"
                DTPickerDate.Value = Date
        Case 1: .Source = "SELECT * from tblpatientbill where (Date = #" & DTPickerSelectDate.Value & "#) and billsuccess = true"
        Case 2: .Source = "SELECT * from tblPatientBill where (Date between #" & DTPickerSelectFrom.Value & "# and #" & DTPickerSelectTo.Value & "#) and billsuccess = true"
    End Select
    
    If .State = 0 Then .Open
    TemReceptionCash = 0
    TemReceptionCard = 0
    TemReceptionCheque = 0
    
    If .RecordCount <> 0 Then
    .MoveFirst
    While .EOF = False
        Select Case !paymentmethod
        Case "Cash": TemReceptionCash = TemReceptionCash + !Cash + !CreditCash
        Case "CreditCard": TemReceptionCard = TemReceptionCard + !CreditCardAmount
        Case "Cheque": TemReceptionCheque = TemReceptionCheque + !ChequeAmount
        End Select
        .MoveNext
    Wend
    
    End If
    
    .Close
End With

    lblReceptionCash.Caption = Format(TemReceptionCash, "0.00")
    lblReceptionCard.Caption = Format(TemReceptionCard, "#0.00")
    lblReceptionCheque.Caption = Format(TemReceptionCheque, "#0.00")
    
    TotalIncome = TotalIncome + TemReceptionCash + TemReceptionCard + TemReceptionCheque

    
End Sub



Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrintReport_Click()


With Printer
    .Font = "Bernard MT Condensed"
    Printer.Print
    .FontSize = 14
'    Printer.Print Tab(2); InstitutionName
'    .FontSize = 9
'    Printer.Print Tab(3); InstitutionAddress
'    Printer.Print Tab(3); InstitutionTelephone
    .FontName = "Arial"
    .FontSize = 10
    Printer.Print
    Printer.Print
    
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print


Dim TemTab1 As Long
Dim TemTab2 As Long
Dim TemTab3 As Long

TemTab1 = 5
TemTab2 = 15
TemTab3 = 45

    Printer.Print Tab(TemTab1);
Select Case SSTab1.Tab
    Case 0: Printer.Print "Totay's (" & Format(Date, "dd mmmm yyyy") & ") Income and Expences"
    Case 1: Printer.Print "Income and Expences on " & Format(DTPickerSelectDate.Value, "dd mmmm yyyy")
    Case 2: Printer.Print "Income and Expences from " & Format(DTPickerSelectFrom.Value, "dd mmmm yyyy") & " to " & Format(DTPickerSelectTo.Value, "dd mmmm yyyy")
End Select
Printer.Print

Printer.Print Tab(TemTab1);
Printer.Print "Cash"
    Printer.Print Tab(TemTab2);
    Printer.Print "Reception";
        Printer.Print Tab(TemTab3);
        Printer.Print lblReceptionCash.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Agent Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblAgentPaymentCash.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Customer Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblCustomerPaymentCash.Caption


Printer.Print Tab(TemTab1);
Printer.Print "Cheque"
    Printer.Print Tab(TemTab2);
    Printer.Print "Reception";
        Printer.Print Tab(TemTab3);
        Printer.Print lblReceptionCheque.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Agent Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblAgentPaymentCheque.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Customer Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblCustomerPaymentCheque.Caption

Printer.Print Tab(TemTab1);
Printer.Print "Credit Card"
    Printer.Print Tab(TemTab2);
    Printer.Print "Reception";
        Printer.Print Tab(TemTab3);
        Printer.Print lblReceptionCard.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Agent Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblReceptionCard.Caption
    Printer.Print Tab(TemTab2);
    Printer.Print "Customer Payment";
        Printer.Print Tab(TemTab3);
        Printer.Print lblReceptionCard.Caption

Printer.Print
Printer.Print Tab(TemTab1);
Printer.Print "Total Income";
Printer.Print Tab(TemTab3);
Printer.Print lblIncome.Caption

Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(TemTab1);
Printer.Print "Total Doctor's Payments";
Printer.Print Tab(TemTab3);
Printer.Print lblExpences.Caption

.EndDoc

End With
End Sub

Private Sub DTPickerSelectDate_Change()
    Call CalculateTotals
End Sub


Private Sub DTPickerSelectFrom_Change()
    Call CalculateTotals
End Sub


Private Sub DTPickerSelectTo_Change()
    Call CalculateTotals
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    DTPickerDate.MaxDate = Date
    DTPickerDate.MinDate = Date
    DTPickerDate.Enabled = False
    Call CalculateTotals
    Call SetColour
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call CalculateTotals
End Sub
