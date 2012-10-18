VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAllBHTSummeryF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All BHT Summeries F"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8220
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   8220
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "P&rocess"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6960
      TabIndex        =   55
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
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
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   50
      Top             =   7680
      Width           =   6135
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   52
         Top             =   240
         Width           =   4575
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   960
         TabIndex        =   51
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Office Copy"
      TabPicture(0)   =   "frmAllBHTSummeryF.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAdditionalCharge"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblN"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblMa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblPay"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblA"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblM"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblR"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblProfessionalCharges"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblNursingCharges"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblMaintainCharges"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblAdmissionFee"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblLinenCharge"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblServiceCharge"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblMedicineCharge"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblRoomCharge"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblBalance"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblPayments"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblTotalCharge"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblTot"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblDis"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblNetCharge"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label6"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblDiscount"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Patient Copy"
      TabPicture(1)   =   "frmAllBHTSummeryF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDiscountF"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblHospitalBillF"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblNetChargeF"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblTotalChargeF"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPaymentsF"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblBalanceF"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblRoomChargeF"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblMedicineChargeF"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblServiceChargeF"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLinenChargeF"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblAdmissionFeeF"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblMaintainChargesF"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblNursingChargesF"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblProfessionalChargesF"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblAdditionalChargeF"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblHCf"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label3"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label23"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label28"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label27"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label26"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label25"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label24"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label20"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label19"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label18"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label17"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label22"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lblBalF"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).ControlCount=   30
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   69
         Top             =   4200
         Width           =   525
      End
      Begin VB.Label lblDiscountF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69600
         TabIndex        =   68
         Top             =   4800
         Width           =   525
      End
      Begin VB.Label lblHospitalBillF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -69630
         TabIndex        =   63
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Net Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblNetCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   61
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label lblNetChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69630
         TabIndex        =   59
         Top             =   5280
         Width           =   525
      End
      Begin VB.Label lblDis 
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lblTotalChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69630
         TabIndex        =   49
         Top             =   4320
         Width           =   525
      End
      Begin VB.Label lblPaymentsF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69630
         TabIndex        =   48
         Top             =   5640
         Width           =   525
      End
      Begin VB.Label lblBalanceF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69630
         TabIndex        =   47
         Top             =   6000
         Width           =   525
      End
      Begin VB.Label lblRoomChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   46
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblMedicineChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   45
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblServiceChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   44
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblLinenChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   43
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblAdmissionFeeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   42
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblMaintainChargesF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   41
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblNursingChargesF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   40
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblProfessionalChargesF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -69630
         TabIndex        =   39
         Top             =   3840
         Width           =   525
      End
      Begin VB.Label lblAdditionalChargeF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   240
         Left            =   -69480
         TabIndex        =   27
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblTot 
         Caption         =   "Total Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblTotalCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label lblRoomCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblMedicineCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblServiceCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLinenCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblAdmissionFee 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblMaintainCharges 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblNursingCharges 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblProfessionalCharges 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblR 
         Caption         =   "Room Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblM 
         Caption         =   "Cost of Medicines"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblS 
         Caption         =   "Cost of Services"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblL 
         Caption         =   "Cost of Linen"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblA 
         Caption         =   "Admission Fee"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblPay 
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label lblBal 
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label lblMa 
         Caption         =   "Maintain Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblN 
         Caption         =   "Nursing Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblP 
         Caption         =   "Professional Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblAdditionalCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblAd 
         Caption         =   "Additional Charge"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblHCf 
         AutoSize        =   -1  'True
         Caption         =   "Total Hospital Charges"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74880
         TabIndex        =   62
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Net Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   58
         Top             =   5280
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   56
         Top             =   4800
         Width           =   1050
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Total Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   33
         Top             =   4320
         Width           =   1590
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Room Charges"
         Height          =   240
         Left            =   -74880
         TabIndex        =   38
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Medicines"
         Height          =   240
         Left            =   -74880
         TabIndex        =   37
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Services"
         Height          =   240
         Left            =   -74880
         TabIndex        =   36
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Linen"
         Height          =   240
         Left            =   -74880
         TabIndex        =   35
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Admission Fee"
         Height          =   240
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Maintain Charges"
         Height          =   240
         Left            =   -74880
         TabIndex        =   30
         Top             =   2280
         Width           =   1485
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nursing Charges"
         Height          =   240
         Left            =   -74880
         TabIndex        =   29
         Top             =   2640
         Width           =   1410
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Professional Charges"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74880
         TabIndex        =   28
         Top             =   3840
         Width           =   2385
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Additional Charge"
         Height          =   240
         Left            =   -74880
         TabIndex        =   26
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   32
         Top             =   5640
         Width           =   1200
      End
      Begin VB.Label lblBalF 
         AutoSize        =   -1  'True
         Caption         =   "Due Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   31
         Top             =   6000
         Width           =   1515
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   64
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   81199107
      CurrentDate     =   39960
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   66
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   81199107
      CurrentDate     =   39960
   End
   Begin btButtonEx.ButtonEx btnPrintPatientCopy 
      Height          =   375
      Left            =   6360
      TabIndex        =   70
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print Patient Copy"
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
   Begin btButtonEx.ButtonEx btnPrintOfficeCopy 
      Height          =   375
      Left            =   6360
      TabIndex        =   71
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print Office Copy"
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
   Begin btButtonEx.ButtonEx btnExcelPatientCopy 
      Height          =   375
      Left            =   6360
      TabIndex        =   72
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Excel Patient Copy"
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
   Begin btButtonEx.ButtonEx btnExcelOfficeCopy 
      Height          =   375
      Left            =   6360
      TabIndex        =   73
      Top             =   3120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Excel Office Copy"
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
   Begin VB.Label Label1 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllBHTSummeryF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    
Private Sub btnClose_Click()
    Unload Me
End Sub



Private Sub PrintingText(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, PrintText As String, PrintAlignment As TextAlignment, ReportPrintFont As ReportFont)
    
    If PrintAlignment = leftAlign Then
        Printer.CurrentX = X1
    ElseIf PrintAlignment = rightAlign Then
        Printer.CurrentX = X2 - Printer.TextWidth(PrintText)
    ElseIf PrintAlignment = CentreAlign Then
        Printer.CurrentX = (X1 + X2 / 2) - (Printer.TextWidth(PrintText) / 2)
    Else
        Printer.CurrentX = X1
    End If
    If Y1 <> 0 Then Printer.CurrentY = Y1
    Printer.Font.Name = ReportPrintFont.Name
    Printer.Font.Size = ReportPrintFont.Size
    Printer.Font.Italic = ReportPrintFont.Italic
    Printer.Font.Bold = ReportPrintFont.Bold
    Printer.Font.Underline = ReportPrintFont.Underline
    
    Printer.Print PrintText
End Sub


Private Sub btnExcelOfficeCopy_Click()
    Dim MyTotal As Double
    Dim MyHSS As New clsHSS
    Dim AllLines() As String
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim TemPath As String
    Dim FSys As New Scripting.FileSystemObject
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myworksheet = myworkbook.Worksheets.Item(1)
    
    myworksheet.Cells(1, 1) = HospitalName
    myworksheet.Cells(2, 1) = "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    
    myworksheet.Cells(4, 3) = lblAdmissionFeeF.Caption
    myworksheet.Cells(5, 3) = lblRoomChargeF.Caption
    myworksheet.Cells(6, 3) = lblMedicineChargeF.Caption
    myworksheet.Cells(7, 3) = lblServiceChargeF.Caption
    myworksheet.Cells(8, 3) = lblProfessionalChargesF.Caption
    myworksheet.Cells(9, 3) = lblAdditionalChargeF.Caption
    myworksheet.Cells(10, 3) = lblLinenChargeF.Caption
    myworksheet.Cells(11, 3) = lblNursingChargesF.Caption
    myworksheet.Cells(12, 3) = lblMaintainChargesF.Caption
    
    myworksheet.Cells(14, 3) = lblTotalChargeF.Caption
    myworksheet.Cells(15, 3) = lblPaymentsF.Caption
    myworksheet.Cells(16, 3) = lblDiscountF.Caption
    myworksheet.Cells(17, 3) = lblBalanceF.Caption

    myworksheet.Cells(4, 1) = "Admission Charges"
    myworksheet.Cells(5, 1) = "Room Charges"
    myworksheet.Cells(6, 1) = "Medicine Charges"
    myworksheet.Cells(7, 1) = "Service Charges"
    myworksheet.Cells(8, 1) = "Professional Charges"
    myworksheet.Cells(9, 1) = "Additional Charges"
    myworksheet.Cells(10, 1) = "Linen Charges"
    myworksheet.Cells(11, 1) = "Nursing Charges"
    myworksheet.Cells(12, 1) = "Maintain Charges"
    
    myworksheet.Cells(14, 1) = "Total Charges"
    myworksheet.Cells(15, 1) = "Payments"
    myworksheet.Cells(16, 1) = "Discounts"
    myworksheet.Cells(17, 1) = "Balance"
    
    myworkbook.SaveAs FSys.GetParentFolderName(Database) & "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & ".xls"
    
    myworkbook.Save
    
    ShellExecute 0&, "open", FSys.GetParentFolderName(Database) & "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & ".xls", "", "", vbMaximizedFocus
    

End Sub

Private Sub btnExcelPatientCopy_Click()
    Dim MyTotal As Double
    Dim MyHSS As New clsHSS
    Dim AllLines() As String
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim TemPath As String
    Dim FSys As New Scripting.FileSystemObject
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myworksheet = myworkbook.Worksheets.Item(1)
    
    myworksheet.Cells(1, 1) = HospitalName
    myworksheet.Cells(2, 1) = "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    
    myworksheet.Cells(4, 3) = lblAdmissionFee.Caption
    myworksheet.Cells(5, 3) = lblRoomCharge.Caption
    myworksheet.Cells(6, 3) = lblMedicineCharge.Caption
    myworksheet.Cells(7, 3) = lblServiceCharge.Caption
    myworksheet.Cells(8, 3) = lblProfessionalCharges.Caption
    myworksheet.Cells(9, 3) = lblAdditionalCharge.Caption
    myworksheet.Cells(10, 3) = lblLinenCharge.Caption
    myworksheet.Cells(11, 3) = lblNursingCharges.Caption
    myworksheet.Cells(12, 3) = lblMaintainCharges.Caption
    
    myworksheet.Cells(14, 3) = lblTotalCharge.Caption
    myworksheet.Cells(15, 3) = lblPayments.Caption
    myworksheet.Cells(16, 3) = lblDiscount.Caption
    myworksheet.Cells(17, 3) = lblBalance.Caption

    myworksheet.Cells(4, 1) = "Admission Charges"
    myworksheet.Cells(5, 1) = "Room Charges"
    myworksheet.Cells(6, 1) = "Medicine Charges"
    myworksheet.Cells(7, 1) = "Service Charges"
    myworksheet.Cells(8, 1) = "Professional Charges"
    myworksheet.Cells(9, 1) = "Additional Charges"
    myworksheet.Cells(10, 1) = "Linen Charges"
    myworksheet.Cells(11, 1) = "Nursing Charges"
    myworksheet.Cells(12, 1) = "Maintain Charges"
    
    myworksheet.Cells(14, 1) = "Total Charges"
    myworksheet.Cells(15, 1) = "Payments"
    myworksheet.Cells(16, 1) = "Discounts"
    myworksheet.Cells(17, 1) = "Balance"



'    myworksheet.Cells(1, 1) = lblAdmissionFeeF.Caption
'    myworksheet.Cells(1, 1) = lblRoomChargeF.Caption
'    myworksheet.Cells(1, 1) = lblMedicineChargeF.Caption
'    myworksheet.Cells(1, 1) = lblServiceChargeF.Caption
'    myworksheet.Cells(1, 1) = lblProfessionalChargesF.Caption
'    myworksheet.Cells(1, 1) = lblAdditionalChargeF.Caption
'    myworksheet.Cells(1, 1) = lblLinenChargeF.Caption
'    myworksheet.Cells(1, 1) = lblNursingChargesF.Caption
'    myworksheet.Cells(1, 1) = lblMaintainChargesF.Caption
'
'    myworksheet.Cells(1, 1) = lblTotalChargeF.Caption
'    myworksheet.Cells(1, 1) = lblPaymentsF.Caption
'    myworksheet.Cells(1, 1) = lblBalanceF.Caption
'    myworksheet.Cells(1, 1) = lblDiscountF.Caption
'    myworksheet.Cells(1, 1) = lblHospitalBillF.Caption
    
    
    myworkbook.SaveAs FSys.GetParentFolderName(Database) & "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & ".xls"
    
    myworkbook.Save
    
    ShellExecute 0&, "open", FSys.GetParentFolderName(Database) & "BHT Summery - All Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & ".xls", "", "", vbMaximizedFocus
    
End Sub

Private Sub btnPrintOfficeCopy_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim CenterX As Long
    Dim FieldX As Long
    Dim SubFieldX As Long
    Dim ValueX As Long
    Dim SubValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 12
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        CenterX = Printer.Width / 2
        FieldX = (1440) * 0.3
        SubFieldX = (1440) * 0.7
        ValueX = Printer.Width - (1440) * 1.3
        SubValueX = Printer.Width - 1440 * 2.7
        
        Printer.Print
        Printer.Print
        
        MyFOnt.Bold = True
        MyFOnt.Size = 13
        temY = Printer.CurrentY
        PrintingText 0, temY, Printer.Width, 0, "RUHUNU HOSPITALS PVT(LTD)", CentreAlign, MyFOnt

        MyFOnt.Bold = False
        MyFOnt.Size = 12

        temY = Printer.CurrentY
        PrintingText 0, temY, Printer.Width, 0, "Karapitiya, Galle", CentreAlign, MyFOnt
        temY = Printer.CurrentY
        
        PrintingText 0, temY, Printer.Width, 0, "Tel. 091 2234059/60, Fax. 091 2234061", CentreAlign, MyFOnt

        Printer.Print
        temY = Printer.CurrentY
        
        PrintingText 0, temY, Printer.Width, 0, "BHT Summery for Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy"), CentreAlign, MyFOnt
        
        MyFOnt.Size = 11
        
        Printer.Print

        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblA.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblAdmissionFee.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblR.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblRoomCharge.Caption, rightAlign, MyFOnt
    
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblM.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMedicineCharge.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblS.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblServiceCharge.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblL.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblLinenCharge.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblMa.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMaintainCharges.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblN.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblNursingCharges.Caption, rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "- - - - - - - -", rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblHCf.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblHospitalBillF.Caption, rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblP.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblProfessionalCharges.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "- - - - - - - -", rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblTot.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotalCharge.Caption, rightAlign, MyFOnt
        
       
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblDis.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblDiscount.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblPay.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblPayments.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "===============", rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblBal.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblBalance.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "===============", rightAlign, MyFOnt
        
        
        Printer.Print
        
        Printer.EndDoc
        
    End If
    


End Sub

Private Sub btnPrintPatientCopy_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim CenterX As Long
    Dim FieldX As Long
    Dim SubFieldX As Long
    Dim ValueX As Long
    Dim SubValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 12
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        CenterX = Printer.Width / 2
        FieldX = (1440) * 0.3
        SubFieldX = (1440) * 0.7
        ValueX = Printer.Width - (1440) * 1.3
        SubValueX = Printer.Width - 1440 * 2.7
        
        Printer.Print
        Printer.Print
        
        MyFOnt.Bold = True
        MyFOnt.Size = 13
        temY = Printer.CurrentY
        PrintingText 0, temY, Printer.Width, 0, "RUHUNU HOSPITALS PVT(LTD)", CentreAlign, MyFOnt

        MyFOnt.Bold = False
        MyFOnt.Size = 12

        temY = Printer.CurrentY
        PrintingText 0, temY, Printer.Width, 0, "Karapitiya, Galle", CentreAlign, MyFOnt
        temY = Printer.CurrentY
        
        PrintingText 0, temY, Printer.Width, 0, "Tel. 091 2234059/60, Fax. 091 2234061", CentreAlign, MyFOnt

        Printer.Print
        temY = Printer.CurrentY
        
        PrintingText 0, temY, Printer.Width, 0, "BHT Summery for Discharges from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy"), CentreAlign, MyFOnt
        
        MyFOnt.Size = 11
        
        Printer.Print
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblA.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblAdmissionFeeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblR.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblRoomChargeF.Caption, rightAlign, MyFOnt
    
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblM.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMedicineChargeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblS.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblServiceChargeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblL.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblLinenChargeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblMa.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMaintainChargesF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblN.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblNursingChargesF.Caption, rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "- - - - - - - -", rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblHCf.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblHospitalBillF.Caption, rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblP.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblProfessionalChargesF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "- - - - - - - -", rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblTot.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotalChargeF.Caption, rightAlign, MyFOnt
        
       
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblDis.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblDiscountF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblPay.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblPaymentsF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "===============", rightAlign, MyFOnt
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblBal.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblBalanceF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "===============", rightAlign, MyFOnt
        
        
        Printer.Print
        
        Printer.EndDoc
        
    End If
    

End Sub

Public Sub btnProcess_Click()
    Call ClearValues
    Call ClearDisplayValues
    Call DisplayDetails
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim Price As Double
    Dim Discount As Double
    Dim NetPrice As Double
    Dim Balance As Double
    Dim AdmissionCharge As Double
    Dim LinanCharge As Double
    Dim RoomCharge As Double
    Dim ServicesCharge As Double
    Dim MaintananceCharge As Double
    Dim NursingCharge As Double
    Dim ProfessionalCharge As Double
    Dim AdditionalCharge As Double
    Dim MedicineCharge As Double
    Dim TotalCharge As Double
    Dim Payments As Double
    Dim FAdmissionRate As Double
    Dim FInitialLinanRate As Double
    Dim FLaterLinanRate As Double
    Dim FMaintananceRate As Double
    Dim FMaintainaceCashDiscountRate As Double
    Dim FNursingRate As Double
    Dim FICUNursingRate As Double
    Dim FAdmissionFee As Double
    Dim FAdmissionCharge As Double
    Dim FLinanCharge As Double
    Dim FRoomCharge As Double
    Dim FServicesCharge As Double
    Dim FMaintananceCharge As Double
    Dim FNursingCharge As Double
    Dim FProfessionalCharge As Double
    Dim FMedicineCharge As Double
    Dim FAdditionalCharge As Double
    Dim FTotalCharge As Double
    Dim FPayments As Double
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select  Sum(Price) as SumOfPrice,  Sum(Discount) as SumOfDiscount,  Sum(NetPrice) as SumOfNetPrice,  Sum(Balance) as SumOfBalance,  Sum(AdmissionCharge) as SumOfAdmissionCharge,  Sum(LinanCharge) as SumOfLinanCharge,  Sum(RoomCharge) as SumOfRoomCharge,  Sum(ServicesCharge) as SumOfServicesCharge,  Sum(MaintananceCharge) as SumOfMaintananceCharge,  Sum(NursingCharge) as SumOfNursingCharge,  Sum(ProfessionalCharge) as SumOfProfessionalCharge,  Sum(AdditionalCharge) as SumOfAdditionalCharge,  Sum(MedicineCharge) as SumOfMedicineCharge,  Sum(TotalCharge) as SumOfTotalCharge,  Sum(Payments) as SumOfPayments,  Sum(FAdmissionRate) as SumOfFAdmissionRate,  Sum(FInitialLinanRate) as SumOfFInitialLinanRate,  Sum(FLaterLinanRate) as SumOfFLaterLinanRate,  Sum(FMaintananceRate) as SumOfFMaintananceRate,  Sum(FMaintainaceCashDiscountRate) as SumOfFMaintainaceCashDiscountRate, " & _
                    "Sum(FNursingRate) as SumOfFNursingRate,  Sum(FICUNursingRate) as SumOfFICUNursingRate,  Sum(FAdmissionFee) as SumOfFAdmissionFee, " & _
                    "Sum(FAdmissionCharge) as SumOfFAdmissionCharge,  Sum(FLinanCharge) as SumOfFLinanCharge,  Sum(FRoomCharge) as SumOfFRoomCharge,  Sum(FServicesCharge) as SumOfFServicesCharge,  Sum(FMaintananceCharge) as SumOfFMaintananceCharge,  Sum(FNursingCharge) as SumOfFNursingCharge,  Sum(FProfessionalCharge) as SumOfFProfessionalCharge,  Sum(FMedicineCharge) as SumOfFMedicineCharge,  Sum(FAdditionalCharge) as SumOfFAdditionalCharge,  Sum(FTotalCharge) as SumOfFTotalCharge,  Sum(FPayments) as SumOfFPayments " & _
                    "from tblBHT " & _
                    "Where IsBHT = 1 AND DOD between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfPrice) = False Then Price = !SumOfPrice
            If IsNull(!SumOfDiscount) = False Then Discount = !SumOfDiscount
            If IsNull(!SumOfNetPrice) = False Then NetPrice = !SumOfNetPrice
            If IsNull(!SumOfBalance) = False Then Balance = !SumOfBalance
            If IsNull(!SumOfAdmissionCharge) = False Then AdmissionCharge = !SumOfAdmissionCharge
            If IsNull(!SumOfLinanCharge) = False Then LinanCharge = !SumOfLinanCharge
            If IsNull(!SumOfRoomCharge) = False Then RoomCharge = !SumOfRoomCharge
            If IsNull(!SumOfServicesCharge) = False Then ServicesCharge = !SumOfServicesCharge
            If IsNull(!SumOfMaintananceCharge) = False Then MaintananceCharge = !SumOfMaintananceCharge
            If IsNull(!SumOfNursingCharge) = False Then NursingCharge = !SumOfNursingCharge
            If IsNull(!SumOfProfessionalCharge) = False Then ProfessionalCharge = !SumOfProfessionalCharge
            If IsNull(!SumOfAdditionalCharge) = False Then AdditionalCharge = !SumOfAdditionalCharge
            If IsNull(!SumOfMedicineCharge) = False Then MedicineCharge = !SumOfMedicineCharge
            If IsNull(!SumOfTotalCharge) = False Then TotalCharge = !SumOfTotalCharge
            If IsNull(!SumOfPayments) = False Then Payments = !SumOfPayments
            If IsNull(!SumOfFAdmissionRate) = False Then FAdmissionRate = !SumOfFAdmissionRate
            If IsNull(!SumOfFInitialLinanRate) = False Then FInitialLinanRate = !SumOfFInitialLinanRate
            If IsNull(!SumOfFLaterLinanRate) = False Then FLaterLinanRate = !SumOfFLaterLinanRate
            If IsNull(!SumOfFMaintananceRate) = False Then FMaintananceRate = !SumOfFMaintananceRate
            If IsNull(!SumOfFMaintainaceCashDiscountRate) = False Then FMaintainaceCashDiscountRate = !SumOfFMaintainaceCashDiscountRate
            If IsNull(!SumOfFNursingRate) = False Then FNursingRate = !SumOfFNursingRate
            If IsNull(!SumOfFICUNursingRate) = False Then FICUNursingRate = !SumOfFICUNursingRate
            If IsNull(!SumOfFAdmissionFee) = False Then FAdmissionFee = !SumOfFAdmissionFee
            If IsNull(!SumOfFAdmissionCharge) = False Then FAdmissionCharge = !SumOfFAdmissionCharge
            If IsNull(!SumOfFLinanCharge) = False Then FLinanCharge = !SumOfFLinanCharge
            If IsNull(!SumOfFRoomCharge) = False Then FRoomCharge = !SumOfFRoomCharge
            If IsNull(!SumOfFServicesCharge) = False Then FServicesCharge = !SumOfFServicesCharge
            If IsNull(!SumOfFMaintananceCharge) = False Then FMaintananceCharge = !SumOfFMaintananceCharge
            If IsNull(!SumOfFNursingCharge) = False Then FNursingCharge = !SumOfFNursingCharge
            If IsNull(!SumOfFProfessionalCharge) = False Then FProfessionalCharge = !SumOfFProfessionalCharge
            If IsNull(!SumOfFMedicineCharge) = False Then FMedicineCharge = !SumOfFMedicineCharge
            If IsNull(!SumOfFAdditionalCharge) = False Then FAdditionalCharge = !SumOfFAdditionalCharge
            If IsNull(!SumOfFTotalCharge) = False Then FTotalCharge = !SumOfFTotalCharge
            If IsNull(!SumOfFPayments) = False Then FPayments = !SumOfFPayments
        End If
    End With
    lblAdmissionFee.Caption = Format(AdmissionCharge, "#,##0.00")
    lblRoomCharge.Caption = Format(RoomCharge, "#,##0.00")
    lblMedicineCharge.Caption = Format(MedicineCharge, "#,##0.00")
    lblServiceCharge.Caption = Format(ServicesCharge, "#,##0.00")
    lblProfessionalCharges.Caption = Format(ProfessionalCharge, "#,##0.00")
    lblAdditionalCharge.Caption = Format(AdditionalCharge, "#,##0.00")
    lblLinenCharge.Caption = Format(LinanCharge, "#,##0.00")
    lblNursingCharges.Caption = Format(NursingCharge, "#,##0.00")
    lblMaintainCharges.Caption = Format(MaintananceCharge, "#,##0.00")
    
    lblTotalCharge.Caption = Format(TotalCharge, "#,##0.00")
    lblPayments.Caption = Format(Payments, "#,##0.00")
    lblDiscount.Caption = Format(Discount, "#,##0.00")
    lblBalance.Caption = Format(Balance, "#,##0.00")

    lblAdmissionFeeF.Caption = Format(FAdmissionCharge, "#,##0.00")
    lblRoomChargeF.Caption = Format(FRoomCharge, "#,##0.00")
    lblMedicineChargeF.Caption = Format(FMedicineCharge, "#,##0.00")
    lblServiceChargeF.Caption = Format(FServicesCharge, "#,##0.00")
    lblProfessionalChargesF.Caption = Format(FProfessionalCharge, "#,##0.00")
    lblAdditionalChargeF.Caption = Format(FAdditionalCharge, "#,##0.00")
    lblLinenChargeF.Caption = Format(FLinanCharge, "#,##0.00")
    lblNursingChargesF.Caption = Format(FNursingCharge, "#,##0.00")
    lblMaintainChargesF.Caption = Format(FMaintananceCharge, "#,##0.00")
    
    lblTotalChargeF.Caption = Format(FTotalCharge, "#,##0.00")
    lblPaymentsF.Caption = Format(FPayments, "#,##0.00")
    lblBalanceF.Caption = Format(Balance, "#,##0.00")
    lblDiscountF.Caption = Format(Discount, "#,##0.00")
    lblHospitalBillF.Caption = Format(FTotalCharge - FProfessionalCharge, "#,##0.00")
    
End Sub



Private Sub ClearDisplayValues()
    lblAdmissionFee.Caption = "0.00"
    lblBalance.Caption = "0.00"
    lblLinenCharge.Caption = "0.00"
    lblMedicineCharge.Caption = "0.00"
    lblPayments.Caption = "0.00"
    lblRoomCharge.Caption = "0.00"
    lblServiceCharge.Caption = "0.00"
    lblAdditionalCharge.Caption = "0.00"
    lblTotalCharge.Caption = "0.00"
    lblAdmissionFeeF.Caption = "0.00"
    lblBalanceF.Caption = "0.00"
    lblLinenChargeF.Caption = "0.00"
    lblMedicineChargeF.Caption = "0.00"
    lblPaymentsF.Caption = "0.00"
    lblRoomChargeF.Caption = "0.00"
    lblServiceChargeF.Caption = "0.00"
    lblAdditionalChargeF.Caption = "0.00"
    lblTotalChargeF.Caption = "0.00"
End Sub


Private Sub ClearValues()
    lblAdmissionFee.Caption = "0.00"
    lblRoomCharge.Caption = "0.00"
    lblMedicineCharge.Caption = "0.00"
    lblServiceCharge.Caption = "0.00"
    lblProfessionalCharges.Caption = "0.00"
    lblAdditionalCharge.Caption = "0.00"
    lblLinenCharge.Caption = "0.00"
    lblNursingCharges.Caption = "0.00"
    lblMaintainCharges.Caption = "0.00"
    
    lblTotalCharge.Caption = "0.00"
    lblPayments.Caption = "0.00"
    lblBalance.Caption = "0.00"

    lblAdmissionFeeF.Caption = "0.00"
    lblRoomChargeF.Caption = "0.00"
    lblMedicineChargeF.Caption = "0.00"
    lblServiceChargeF.Caption = "0.00"
    lblProfessionalChargesF.Caption = "0.00"
    lblAdditionalChargeF.Caption = "0.00"
    lblLinenChargeF.Caption = "0.00"
    lblNursingChargesF.Caption = "0.00"
    lblMaintainChargesF.Caption = "0.00"
    
    lblTotalChargeF.Caption = "0.00"
    lblPaymentsF.Caption = "0.00"
    lblBalanceF.Caption = "0.00"
    
End Sub


Private Sub Form_Load()
    Call PopulatePrinters
    Call PopulatePapers
    Call GetSettings
End Sub


Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    SSTab1.Tab = GetSetting(App.EXEName, Me.Name, SSTab1.Name, "0")
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
    SaveCommonSettings Me
End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

