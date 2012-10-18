VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGSBSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Sheet Bill Summery"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12990
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
   ScaleHeight     =   9720
   ScaleWidth      =   12990
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   7080
      TabIndex        =   2
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
      Left            =   11640
      TabIndex        =   46
      Top             =   8760
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
      Height          =   855
      Left            =   120
      TabIndex        =   41
      Top             =   8760
      Width           =   9735
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   43
         Top             =   240
         Width           =   4575
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   6360
         TabIndex        =   42
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   5640
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Office Copy"
      TabPicture(0)   =   "frmGSBSummery.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblP"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPay"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblM"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProfessionalCharges"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblServiceCharge"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblMedicineCharge"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblBalance"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPayments"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTotalCharge"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblTot"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDis"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblNetCharge"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "SSTab2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDiscount"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Patient Copy"
      TabPicture(1)   =   "frmGSBSummery.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDiscountF"
      Tab(1).Control(1)=   "SSTab3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "lblNetChargeF"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(7)=   "Label26"
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(9)=   "lblTotalChargeF"
      Tab(1).Control(10)=   "lblPaymentsF"
      Tab(1).Control(11)=   "lblBalanceF"
      Tab(1).Control(12)=   "lblMedicineChargeF"
      Tab(1).Control(13)=   "lblServiceChargeF"
      Tab(1).Control(14)=   "lblProfessionalChargesF"
      Tab(1).Control(15)=   "Label22"
      Tab(1).Control(16)=   "lblBalF"
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   3435
         Width           =   1455
      End
      Begin VB.TextBox txtDiscountF 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73200
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   3555
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4695
         Left            =   -71640
         TabIndex        =   28
         Top             =   480
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Medicines"
         TabPicture(0)   =   "frmGSBSummery.frx":0038
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "gridMedicinesF"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Services"
         TabPicture(1)   =   "frmGSBSummery.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridServiceF"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Professional"
         TabPicture(2)   =   "frmGSBSummery.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "gridProfessionalF"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Payments"
         TabPicture(3)   =   "frmGSBSummery.frx":008C
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "gridPaymentsF"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid gridMedicinesF 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   55
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridServiceF 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   56
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridProfessionalF 
            Height          =   4815
            Left            =   -74880
            TabIndex        =   57
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   8493
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridPaymentsF 
            Height          =   4095
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
            WordWrap        =   -1  'True
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4695
         Left            =   3360
         TabIndex        =   15
         Top             =   480
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Medicines"
         TabPicture(0)   =   "frmGSBSummery.frx":00A8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "gridMedicines"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Services"
         TabPicture(1)   =   "frmGSBSummery.frx":00C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridService"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Professional"
         TabPicture(2)   =   "frmGSBSummery.frx":00E0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "gridProfessional"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Payments"
         TabPicture(3)   =   "frmGSBSummery.frx":00FC
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "gridPayments"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid gridMedicines 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   59
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridService 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   60
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridProfessional 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   61
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridPayments 
            Height          =   4095
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7223
            _Version        =   393216
            WordWrap        =   -1  'True
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
         End
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
         TabIndex        =   53
         Top             =   3840
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
         Left            =   1320
         TabIndex        =   54
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Left            =   -74880
         TabIndex        =   51
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblNetChargeF 
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
         Left            =   -73680
         TabIndex        =   52
         Top             =   4080
         Width           =   1935
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
         TabIndex        =   49
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   -74880
         TabIndex        =   47
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label23 
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
         Left            =   -74880
         TabIndex        =   32
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "Cost of Medicines"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Cost of Services"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Professional Charges"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblTotalChargeF 
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
         Left            =   -73680
         TabIndex        =   40
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblPaymentsF 
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
         Left            =   -73680
         TabIndex        =   39
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblBalanceF 
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
         Left            =   -73680
         TabIndex        =   38
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblMedicineChargeF 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -73680
         TabIndex        =   37
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblServiceChargeF 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -73560
         TabIndex        =   36
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblProfessionalChargesF 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -73680
         TabIndex        =   35
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label22 
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
         Left            =   -74880
         TabIndex        =   31
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblBalF 
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
         Left            =   -74880
         TabIndex        =   30
         Top             =   4800
         Width           =   1695
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
         TabIndex        =   19
         Top             =   3120
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
         Left            =   1320
         TabIndex        =   27
         Top             =   3120
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
         Left            =   1320
         TabIndex        =   26
         Top             =   4440
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
         Left            =   1320
         TabIndex        =   25
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblMedicineCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblServiceCharge 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblProfessionalCharges 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblM 
         Caption         =   "Medicines"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblS 
         Caption         =   "Services"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
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
         TabIndex        =   18
         Top             =   4440
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
         TabIndex        =   17
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label lblP 
         Caption         =   "Prof. Charges"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discharge"
      Height          =   3135
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkSamll 
         Caption         =   "&Small Print"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Print patient copy when Discharged"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   4095
      End
      Begin btButtonEx.ButtonEx btnDischarge 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "&Discharge"
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker dtpTOD 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   70647810
         CurrentDate     =   39960
      End
      Begin MSComCtl2.DTPicker dtpDOD 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   70647811
         CurrentDate     =   39960
      End
      Begin btButtonEx.ButtonEx btnPrint 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "&Print Patient copy"
         Enabled         =   0   'False
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
      Begin btButtonEx.ButtonEx btnOfficePrint 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "Paint &Office copy"
         Enabled         =   0   'False
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
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDetails 
      Height          =   2655
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   600
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Details"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Green Sheet Bill No"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmGSBSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim InwardPtCh As New clsInwardPatientCharges
    Dim MyBHT As New clsBHT
    
    Dim PtAdmissionFee As Double
    Dim PtRoomCharge As Double
    Dim PtServiceCharge As Double
    Dim PtMedicineCharge As Double
    Dim PtLinanCharge As Double
    Dim PtProfCharge As Double
    Dim PtNursingCharge As Double
    Dim PtMaintananceCharge As Double
    Dim PtAdditionalCharge As Double
    
    Dim PtTotalCharge As Double
    Dim PtTotalPayments As Double
    Dim PtBalance As Double
    
    Dim FPtAdmissionFee As Double
    Dim FPtRoomCharge As Double
    Dim FPtServiceCharge As Double
    Dim FPtMedicineCharge As Double
    Dim FPtLinanCharge As Double
    Dim FPtProfCharge As Double
    Dim FPtNursingCharge As Double
    Dim FPtMaintananceCharge As Double
    Dim FPtAdditionalCharge As Double
    
    Dim FPtTotalCharge As Double
    Dim FPtTotalPayments As Double
    Dim FPtBalance As Double
    
    
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
    
Private Sub OldDisplayDetails(): On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "GSB : " & MyBHT.BHT & vbNewLine
    'temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
    
    txtDiscount.Text = Format(MyBHT.Discount, "0.00")
    txtDiscountF.Text = Format(MyBHT.Discount, "0.00")
    
End Sub
    
    
Private Sub DisplayDetails(): On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = MyBHT.FirstName & vbNewLine
    temText = temText & "GSB : " & MyBHT.BHT & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
    
    txtDiscount.Text = Format(MyBHT.Discount, "0.00")
    txtDiscountF.Text = Format(MyBHT.Discount, "0.00")
    
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDischarge_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    
    btnProcess_Click
    
    Dim MyBHTID As Long
    
    i = MsgBox("Are you sure you want to discharge this patient", vbYesNo)
    If i = vbNo Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Discharge = True
            !DOD = dtpDOD.Value
            !TOD = Format(dtpDOD.Value, "dd MMMM yyyy") & " " & dtpTOD.Value
            !DisStaffID = UserID
            !ServicesCharge = PtServiceCharge
            !ProfessionalCharge = PtProfCharge
            !MedicineCharge = PtMedicineCharge
            !TotalCharge = PtTotalCharge
            !Payments = PtTotalPayments
            !Balance = Val(lblBalanceF.Caption)
            
            !FServicesCharge = FPtServiceCharge
            !FMedicineCharge = FPtMedicineCharge
            !FProfessionalCharge = FPtProfCharge
            !FTotalCharge = FPtTotalCharge
            !FPayments = FPtTotalPayments
            
            !Price = FPtTotalCharge
            !Discount = Val(txtDiscount.Text)
            If FPtTotalCharge <> 0 Then
                !DiscountPercent = Val(txtDiscount.Text) / FPtTotalCharge * 100
            End If
            !NetPrice = FPtTotalCharge - Val(txtDiscount.Text)
            
            .Update
        End If
        .Close
    End With
    
    MyBHTID = Val(cmbBHT.BoundText)
    
    Call ClearValues
    Call FillCombos
    cmbBHT.Text = Empty
    
    cmbBHT.BoundText = MyBHTID
    btnProcess_Click
    If chkPrint.Value = 1 Then btnPrint_Click
    
End Sub

Private Sub ActualCalculations()
    
         PtAdmissionFee = 0
         PtRoomCharge = 0
         PtServiceCharge = 0
         PtMedicineCharge = 0
         PtLinanCharge = 0
         PtProfCharge = 0
         PtNursingCharge = 0
         PtMaintananceCharge = 0
         PtTotalCharge = 0
         PtTotalPayments = 0
         PtBalance = 0
         PtAdditionalCharge = 0
    
    
    If IsNumeric(cmbBHT.BoundText) = False Then
        btnDischarge.Enabled = False
         PtAdmissionFee = 0
         PtRoomCharge = 0
         PtServiceCharge = 0
         PtMedicineCharge = 0
         PtLinanCharge = 0
         PtProfCharge = 0
         PtNursingCharge = 0
         PtMaintananceCharge = 0
         PtTotalCharge = 0
         PtTotalPayments = 0
         PtBalance = 0
         PtAdditionalCharge = 0
        Exit Sub
    End If

    
    PtAdmissionFee = 0 'InwardPtCh.AdimssionRate
    PtRoomCharge = FillRooms
    PtMedicineCharge = FillMedicines
    PtServiceCharge = FillServices
    PtLinanCharge = LinanCharges
    PtProfCharge = FillProfessionalCharges
    PtNursingCharge = NursingCharge
    PtMaintananceCharge = MaintananceCharges
    PtAdditionalCharge = AdditionalCharge
    
    PtTotalPayments = FillPayments
    
    PtTotalCharge = PtRoomCharge + PtServiceCharge + PtMedicineCharge + PtLinanCharge + PtProfCharge + PtNursingCharge + PtMaintananceCharge
    PtBalance = PtTotalCharge - PtTotalPayments - Val(txtDiscount.Text)
    
    lblMedicineCharge.Caption = Format(PtMedicineCharge, "0.00")
    lblProfessionalCharges.Caption = Format(PtProfCharge, "0.00")
    lblServiceCharge.Caption = Format(PtServiceCharge, "0.00")
    
    lblTotalCharge.Caption = Format(PtTotalCharge, "0.00")
    lblPayments.Caption = Format(PtTotalPayments, "0.00")
    
    
    lblNetCharge.Caption = Format(PtTotalCharge - MyBHT.Discount, "0.00")
    
    If PtBalance >= 0 Then
        lblBal.Caption = "Balance"
        lblBalance.Caption = Format(Abs(PtBalance), "0.00")
    Else
        lblBal.Caption = "Excess"
        lblBalance.Caption = Format(Abs(PtBalance), "0.00")
    End If

End Sub

Private Sub FakeCalculations()
         FPtAdmissionFee = 0
         FPtRoomCharge = 0
         FPtServiceCharge = 0
         FPtMedicineCharge = 0
         FPtLinanCharge = 0
         FPtProfCharge = 0
         FPtNursingCharge = 0
         FPtMaintananceCharge = 0
         FPtTotalCharge = 0
         FPtTotalPayments = 0
         FPtBalance = 0
         FPtAdditionalCharge = 0
    
    
    If IsNumeric(cmbBHT.BoundText) = False Then
        btnDischarge.Enabled = False
         FPtAdmissionFee = 0
         FPtRoomCharge = 0
         FPtServiceCharge = 0
         FPtMedicineCharge = 0
         FPtLinanCharge = 0
         FPtProfCharge = 0
         FPtNursingCharge = 0
         FPtMaintananceCharge = 0
         FPtTotalCharge = 0
         FPtTotalPayments = 0
         FPtBalance = 0
         FPtAdditionalCharge = 0
        Exit Sub
    End If
    

    
    FPtAdmissionFee = 0 ' InwardPtCh.AdimssionRate
    FPtRoomCharge = FFillRooms
    FPtMedicineCharge = FFillMedicines + AdditionalCharge
    FPtServiceCharge = FFillServices
    FPtLinanCharge = FLinanCharges
    FPtProfCharge = FFillProfessionalCharges
    FPtNursingCharge = FNursingCharge
    FPtMaintananceCharge = FMaintananceCharges
    FPtAdditionalCharge = FAdditionalCharge
    
    FPtTotalPayments = FFillPayments
    
    FPtTotalCharge = FPtRoomCharge + FPtServiceCharge + FPtMedicineCharge + FPtLinanCharge + FPtProfCharge + FPtNursingCharge + FPtMaintananceCharge
    FPtBalance = FPtTotalCharge - FPtTotalPayments - Val(txtDiscountF.Text)
    
    lblMedicineChargeF.Caption = Format(FPtMedicineCharge, "0.00")
    lblProfessionalChargesF.Caption = Format(FPtProfCharge, "0.00")
    lblServiceChargeF.Caption = Format(FPtServiceCharge, "0.00")
    
    lblTotalChargeF.Caption = Format(FPtTotalCharge, "0.00")
    lblPaymentsF.Caption = Format(FPtTotalPayments, "0.00")

    lblNetChargeF.Caption = Format(FPtTotalCharge - MyBHT.Discount, "0.00")

    If FPtBalance >= 0 Then
        lblBalF.Caption = "Balance"
        lblBalanceF.Caption = Format(Abs(FPtBalance), "0.00")
    Else
        lblBalF.Caption = "Excess"
        lblBalanceF.Caption = Format(Abs(FPtBalance), "0.00")
    End If


End Sub

Private Sub btnOfficePrint_Click()
    
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
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        CenterX = Printer.Width / 2
        FieldX = (1440) * 0.3
        SubFieldX = (1440) * 0.7
        ValueX = Printer.Width - (1440) * 1.3
        SubValueX = Printer.Width - 1440 * 2.7
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 13
        PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 12
        PrintingText 0, 0, Printer.Width, 0, HospitalDescreption, CentreAlign, MyFOnt
        PrintingText 0, 0, Printer.Width, 0, HospitalAddress, CentreAlign, MyFOnt
        PrintingText 0, temY, Printer.Width, 0, "Green Sheet BILL - " & MyBHT.BHTID, CentreAlign, MyFOnt
        
        
        MyFOnt.Size = 11
        
        Printer.Print
        
        AllLines = SeperateLines(txtDetails.Text)
        For i = 0 To UBound(AllLines) - 1
                If InStr(UCase(AllLines(i)), UCase("PAYMENT")) > 0 Then
                    DoEvents
                Else
                    PrintingText FieldX, 0, ValueX, 0, AllLines(i), leftAlign, MyFOnt
                End If
        Next
        
        Printer.Print
        
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, lblA.Caption, LeftAlign, MyFOnt
'        PrintingText FieldX, temY, ValueX, 0, lblAdmissionFee.Caption, RightAlign, MyFOnt
'
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblM.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMedicineCharge.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblS.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblServiceCharge.Caption, rightAlign, MyFOnt
        
        For i = 0 To gridService.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridService.TextMatrix(i, 0), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridService.TextMatrix(i, 1), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblP.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblProfessionalCharges.Caption, rightAlign, MyFOnt
        
        If MyBHT.PaymentMethod = "Credit" Then
            For i = 0 To gridProfessional.Rows - 1
                temY = Printer.CurrentY
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessional.TextMatrix(i, 4), leftAlign, MyFOnt
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessional.TextMatrix(i, 6), rightAlign, MyFOnt
            Next
        End If
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblTot.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotalCharge.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblPay.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblPayments.Caption, rightAlign, MyFOnt
        
        
        For i = 0 To gridPayments.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridPayments.TextMatrix(i, 1), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridPayments.TextMatrix(i, 2), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblDis.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtDiscount.Text, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblBal.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblBalance.Caption, rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, ".......................", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, ".......................", rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Patient / Guardian", rightAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
    

End Sub

Private Sub btnPrint_Click()
    If chkSamll.Value = 1 Then
        Call NewPrint
    Else
        Call OldPrint
    End If
End Sub
Private Sub NewPrint()
    
    Dim CenterX As Long
    Dim FieldX As Long
    Dim SubFieldX As Long
    Dim ValueX As Long
    Dim SubValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    
    Dim temFont As ReportFont
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim temBillPoints As MyBillPoints
    
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 9
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        
        temBillPoints = PrintThisBill(MyBHT.BHTID, MyBHT.PaymentMethod, MyBHT.FirstName, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Green Sheet No. - " & MyBHT.BHT, "")
        
'
''        CenterX = Printer.Width / 2
''        FieldX = (1440) * 0.2
'        SubFieldX = (1440) * 0.4
''        ValueX = Printer.Width - (1440) * 0.9
'        SubValueX = Printer.Width - 1440 * 1.9
'
        CenterX = temBillPoints.CenterX
        FieldX = temBillPoints.DX
        SubFieldX = (1440) * 0.6
        ValueX = temBillPoints.VX
        SubValueX = temBillPoints.VX - 1440
        
        Printer.FontSize = 6
'        Printer.Print
'        Printer.Print
'        Printer.Print
'        Printer.Print
        
        Printer.FontSize = 10
        
 '       temY = Printer.CurrentY
        
'        MyFOnt.Bold = False
'        MyFOnt.Size = 13
'        PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
'        temY = Printer.CurrentY
'
'        MyFOnt.Bold = False
'        MyFOnt.Size = 12
'        PrintingText 0, 0, Printer.Width, 0, HospitalDescreption, CentreAlign, MyFOnt
'        temY = Printer.CurrentY
'
'        PrintingText 0, 0, Printer.Width, 0, HospitalAddress, CentreAlign, MyFOnt
'        temY = Printer.CurrentY
        
        temY = temBillPoints.DY
        
'        PrintingText 0, temY, Printer.Width, 0, "           Green Sheet BILL -    " & MyBHT.BHTID, CentreAlign, MyFOnt
        
        
        MyFOnt.Size = 9
        
'        Printer.Print
        
'        AllLines = SeperateLines(txtDetails.Text)
'        For i = 0 To UBound(AllLines) - 1
'            PrintingText FieldX + 3140, 0, ValueX, 0, AllLines(i), leftAlign, MyFOnt
'        Next
        
        Printer.CurrentY = temY
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblM.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMedicineChargeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblS.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblServiceChargeF.Caption, rightAlign, MyFOnt
        
        For i = 0 To gridServiceF.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridServiceF.TextMatrix(i, 0), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridServiceF.TextMatrix(i, 1), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblP.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblProfessionalChargesF.Caption, rightAlign, MyFOnt
        
        If MyBHT.PaymentMethod = "Credit" Then
            For i = 0 To gridProfessionalF.Rows - 1
                temY = Printer.CurrentY
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessionalF.TextMatrix(i, 4), leftAlign, MyFOnt
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessionalF.TextMatrix(i, 6), rightAlign, MyFOnt
            Next
        End If
        
        Printer.Print
        temY = Printer.CurrentY
        
        temY = Printer.CurrentY
        
        
        MyFOnt.Bold = True
        MyFOnt.Underline = True
        
        PrintingText FieldX, temY, ValueX, 0, lblTot.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotalChargeF.Caption, rightAlign, MyFOnt
        
        MyFOnt.Underline = False
        MyFOnt.Bold = False
        
        
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblPay.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblPaymentsF.Caption, rightAlign, MyFOnt
        
        
        For i = 0 To gridPaymentsF.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridPaymentsF.TextMatrix(i, 1), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridPaymentsF.TextMatrix(i, 2), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblDis.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtDiscountF.Text, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblBal.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblBalanceF.Caption, rightAlign, MyFOnt
        
'        Printer.Print
'        Printer.Print
'        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "...........", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "...........", rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Patient / Guardian", rightAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
    

End Sub


Private Sub OldPrint()
    
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
        temY = Printer.CurrentY
        
        MyFOnt.Bold = False
        MyFOnt.Size = 13
        PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
        temY = Printer.CurrentY
        
        MyFOnt.Bold = False
        MyFOnt.Size = 12
        PrintingText 0, 0, Printer.Width, 0, HospitalDescreption, CentreAlign, MyFOnt
        temY = Printer.CurrentY
        
        PrintingText 0, 0, Printer.Width, 0, HospitalAddress, CentreAlign, MyFOnt
        temY = Printer.CurrentY
        
        PrintingText 0, temY, Printer.Width, 0, "Green Sheet BILL - " & MyBHT.BHTID, CentreAlign, MyFOnt
        
        
        MyFOnt.Size = 11
        
        Printer.Print
        
        AllLines = SeperateLines(txtDetails.Text)
        For i = 0 To UBound(AllLines) - 1
                If InStr(UCase(AllLines(i)), UCase("PAYMENT")) > 0 Then
                    DoEvents
                Else
                    PrintingText FieldX, 0, ValueX, 0, AllLines(i), leftAlign, MyFOnt
                End If
        Next
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblM.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblMedicineChargeF.Caption, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblS.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblServiceChargeF.Caption, rightAlign, MyFOnt
        
        For i = 0 To gridServiceF.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridServiceF.TextMatrix(i, 0), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridServiceF.TextMatrix(i, 1), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblP.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblProfessionalChargesF.Caption, rightAlign, MyFOnt
        
        If MyBHT.PaymentMethod = "Credit" Then
            For i = 0 To gridProfessionalF.Rows - 1
                temY = Printer.CurrentY
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessionalF.TextMatrix(i, 4), leftAlign, MyFOnt
                PrintingText SubFieldX, temY, SubValueX, 0, gridProfessionalF.TextMatrix(i, 6), rightAlign, MyFOnt
            Next
        End If
        
        Printer.Print
        temY = Printer.CurrentY
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblTot.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotalChargeF.Caption, rightAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblPay.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblPaymentsF.Caption, rightAlign, MyFOnt
        
        
        For i = 0 To gridPaymentsF.Rows - 1
            temY = Printer.CurrentY
            PrintingText SubFieldX, temY, SubValueX, 0, gridPaymentsF.TextMatrix(i, 1), leftAlign, MyFOnt
            PrintingText SubFieldX, temY, SubValueX, 0, gridPaymentsF.TextMatrix(i, 2), rightAlign, MyFOnt
        Next
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblDis.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtDiscountF.Text, rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, lblBal.Caption, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblBalanceF.Caption, rightAlign, MyFOnt
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, ".......................", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, ".......................", rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Patient / Guardian", rightAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
    

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


Public Sub btnProcess_Click()
    Call ClearValues
    Call ClearDisplayValues
    
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    
    
    If MyBHT.Discharge = True Then
        btnDischarge.Enabled = False
        dtpDOD.Value = MyBHT.DOD
        dtpTOD.Value = MyBHT.TOD
        dtpDOD.Enabled = False
        dtpTOD.Enabled = False
    Else
        btnDischarge.Enabled = True
        dtpDOD.Enabled = False
        dtpDOD.Value = MyBHT.DOA
        dtpTOD.Enabled = True
    End If
    
    txtDiscount.Text = Format(MyBHT.Discount, "0.00")
    txtDiscountF.Text = Format(MyBHT.Discount, "0.00")
    
    If MyBHT.Discharge = True Then
        btnDischarge.Enabled = False
    Else
        btnDischarge.Enabled = True
    End If
    
    Call ActualCalculations
    Call FakeCalculations
    Call DisplayDetails
    
    btnPrint.Enabled = True
    btnOfficePrint.Enabled = True

End Sub

Private Sub cmbBHT_Change()
    Call ClearValues
    Call ClearDisplayValues
End Sub

Private Function AdditionalCharge() As Double
'    Dim rsTem As New ADODB.Recordset
    Dim TotalValue As Double
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT tblPatientCharge.* " & _
'                    "FROM tblPatientCharge " & _
'                    "WHERE Deleted = 0   AND BHTID = " & Val(cmbBHT.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            TotalValue = TotalValue + !Charge
'            .MoveNext
'        Wend
'        .Close
'    End With
    AdditionalCharge = TotalValue
End Function

Private Function MaintananceCharges() As Double
    MaintananceCharges = 0
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'    Dim temHours As Long
'
'    FromDate = MyBHT.DOA
'    FromTime = MyBHT.TOA
'    If MyBHT.Discharge = False Then
'        ToDate = dtpDOD.Value
'        ToTime = dtpTOD.Value
'    Else
'        ToDate = MyBHT.DOD
'        ToTime = MyBHT.TOD
'    End If
'    temHours = Abs(DateDiff("h", FromDate + FromTime, ToDate + ToTime))
'    If temHours > 1 Then
'        If MyBHT.PaymentMethod <> "Cash" Then
'            If temHours Mod 6 = 0 Then
'                MaintananceCharges = (temHours \ 6) * (InwardPtCh.MaintananceRate / 4)
'            Else
'                MaintananceCharges = ((temHours \ 6) + 1) * (InwardPtCh.MaintananceRate / 4)
'            End If
'        Else
'            If temHours Mod 6 = 0 Then
'                MaintananceCharges = (temHours \ 6) * ((InwardPtCh.MaintananceRate - InwardPtCh.MaintainaceCashDiscountRate) / 4)
'            Else
'                MaintananceCharges = ((temHours \ 6) + 1) * ((InwardPtCh.MaintananceRate - InwardPtCh.MaintainaceCashDiscountRate) / 4)
'            End If
'        End If
'    End If
'    If MyBHT.Discharge = False Then MaintananceCharges = MaintananceCharges * (100 + MyBHT.PtSurcharge) / 100

End Function

Private Function NursingCharge()
    NursingCharge = 0
'    Dim rsRoom As New ADODB.Recordset
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If MyRoom.ICUNursing = True Then
'                        If DurationHours Mod 6 = 0 Then
'                            DurationFee = DurationFee + ((DurationHours \ 6)) * ((InwardPtCh.ICUNursingRate) / 4)
'                        Else
'                            DurationFee = DurationFee + ((DurationHours \ 6) + 1) * ((InwardPtCh.ICUNursingRate) / 4)
'                        End If
'                Else
'                        If DurationHours Mod 6 = 0 Then
'                            DurationFee = DurationFee + ((DurationHours \ 6)) * ((InwardPtCh.NursingRate) / 4)
'                        Else
'                            DurationFee = DurationFee + ((DurationHours \ 6) + 1) * ((InwardPtCh.NursingRate) / 4)
'                        End If
'                End If
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'    If MyBHT.Discharge = False Then
'        TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'    Else
'        TotalFee = TotalFee + DurationFee
'    End If
'    NursingCharge = TotalFee
'
End Function


Private Function FillProfessionalCharges()
    With gridProfessional
        .Rows = 1
        .Cols = 7
        .Col = 1
        .Text = "Date"
        .Col = 2
        .Text = "Time"
        .Col = 3
        .Text = "Speciality"
        .Col = 4
        .Text = "Name"
        .Col = 5
        .Text = "Comments"
        .Col = 6
        .Text = "Value"
        .ColWidth(0) = 0
        .ColWidth(2) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 1600
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1200
    End With
    Dim TotalValue As Double
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblSpeciality.Speciality, tblTitle.Title, tblStaff.Name, tblProfessionalCharges.* " & _
                    "FROM ((tblSpeciality RIGHT JOIN tblStaff ON tblSpeciality.SpecialityID = tblStaff.SpecialityID) RIGHT JOIN tblProfessionalCharges ON tblStaff.StaffID = tblProfessionalCharges.StaffID) LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID " & _
                    "WHERE (((tblProfessionalCharges.ProfessionalCharge)=1) AND ((tblProfessionalCharges.Cancelled)=0) AND ((tblProfessionalCharges.ForBHTID)=" & Val(cmbBHT.BoundText) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            gridProfessional.Rows = gridProfessional.Rows + 1
            gridProfessional.Row = gridProfessional.Rows - 1
            
            gridProfessional.Col = 0
            gridProfessional.Text = !ProfessionalChargesID
            
            gridProfessional.Col = 1
            gridProfessional.Text = Format(!Date, "dd MMM yyyy")
            
            gridProfessional.Col = 2
            gridProfessional.Text = Format(!Time, "HH MM")
            
            gridProfessional.Col = 3
            gridProfessional.Text = !Speciality
            
            gridProfessional.Col = 4
            gridProfessional.Text = !Title & " " & !Name
            
            gridProfessional.Col = 5
            gridProfessional.Text = !Comments
            
            gridProfessional.Col = 6
            gridProfessional.Text = Format(!Fee, "0.00")
            
            TotalValue = TotalValue + !Fee
            
            .MoveNext
        
        Wend
        .Close
    End With
    FillProfessionalCharges = TotalValue
End Function

Private Sub ClearDisplayValues()
    lblBalance.Caption = "0.00"
    lblMedicineCharge.Caption = "0.00"
    lblPayments.Caption = "0.00"
    lblServiceCharge.Caption = "0.00"
    lblTotalCharge.Caption = "0.00"
    lblBalanceF.Caption = "0.00"
    lblMedicineChargeF.Caption = "0.00"
    lblPaymentsF.Caption = "0.00"
    lblServiceChargeF.Caption = "0.00"
    lblTotalChargeF.Caption = "0.00"
End Sub

Private Function LinanCharges()
    LinanCharges = 0
'    Dim rsRoom As New ADODB.Recordset
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If DurationHours > 3 * 24 Then
'                    If (DurationHours - (3 * 24)) Mod 6 = 0 Then
'                        DurationFee = DurationFee + InwardPtCh.InitialLinanRate + (((DurationHours - (3 * 24)) \ 6) * (InwardPtCh.LaterLinanRate / 4))
'                    Else
'                        DurationFee = DurationFee + InwardPtCh.InitialLinanRate + (((DurationHours - ((3 * 24)) \ 6) + 1) * (InwardPtCh.LaterLinanRate / 4))
'                    End If
'                Else
'                    DurationFee = DurationFee + InwardPtCh.InitialLinanRate
'                End If
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'    If MyBHT.Discharge = False Then
'        TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'    Else
'        TotalFee = TotalFee + DurationFee
'    End If
'    LinanCharges = TotalFee
End Function

Private Function FillPayments()
    Dim TotalPayments As Double
    Dim rsTem As New ADODB.Recordset
    With gridPayments
        .Cols = 3
        .Rows = 0
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsGSBill = 1 AND Cancelled = 0  AND BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPayments.Rows = gridPayments.Rows + 1
            gridPayments.Row = gridPayments.Rows - 1
            gridPayments.Col = 0
            gridPayments.Text = Format(!Date, "dd MMMM yyyy")
            gridPayments.Col = 2
            gridPayments.Text = Format(!NetTotal, "0.00")
            gridPayments.Col = 1
            'gridPayments.Text = !IncomeBillID
            gridPayments.Text = !DisplayBillID
            
            TotalPayments = TotalPayments + !NetTotal
            
            .MoveNext
        Wend
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
'        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsInwardPaymentBill = 1 AND Cancelled = 0  AND BHTID = " & Val(cmbBHT.BoundText)
        temSql = "SELECT tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime, dbo.tblIncomeBill.PaymentComments " & _
                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (dbo.tblIncomeBill.IsHSSPaymentBill = 1) AND (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.BHTID = " & Val(cmbBHT.BoundText) & ") AND (dbo.tblIncomeBill.Cancelled = 0) " & _
                    "ORDER BY dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            gridPayments.Rows = gridPayments.Rows + 1
            gridPayments.Row = gridPayments.Rows - 1
            gridPayments.Col = 0
            gridPayments.Text = Format(!CompletedDate, "dd MMMM yyyy")
            gridPayments.Col = 2
            gridPayments.Text = Format(!NetTotal, "0.00")
            gridPayments.Col = 1
            gridPayments.Text = !DisplayBillID
            TotalPayments = TotalPayments + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeReturnBill.* FROM tblIncomeReturnBill WHERE tblIncomeReturnBill.BHTID =" & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPayments.Rows = gridPayments.Rows + 1
            gridPayments.Row = gridPayments.Rows - 1
            gridPayments.Col = 0
            gridPayments.Text = Format(!ReturnDate, "dd MMMM yyyy")
            gridPayments.Col = 1
            gridPayments.Text = !IncomeReturnBillID
            gridPayments.Col = 2
            gridPayments.Text = Format(!ReturnValue, "0.00")
            TotalPayments = TotalPayments - !ReturnValue
            .MoveNext
        Wend
        .Close
    End With
    
    
    
    FillPayments = TotalPayments
End Function

Private Function FillServices() As Double
    Dim TotalFee As Double
    Dim rsTem As New ADODB.Recordset
    
    With gridService
        .Clear
        .Cols = 2
        .Rows = 0
        .ColWidth(0) = 3600
        .ColWidth(1) = .Width - .ColWidth(0) - 150
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblServiceSubcategory.ServiceSubcategory, tblServiceCategory.ServiceCategory, Sum(tblPatientService.Charge) AS SumOfCharge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
                    "GROUP BY tblServiceCategory.ServiceCategory, dbo.tblServiceSubcategory.ServiceSubcategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !ServiceCategory & " " & !ServiceSubcategory
            gridService.Col = 1
            If MyBHT.Discharge = False Then
                gridService.Text = Format(!SumOfCharge * (100 + MyBHT.PtSurcharge) / 100, "0.00")
                TotalFee = TotalFee + (!SumOfCharge * (100 + MyBHT.PtSurcharge) / 100)
            Else
                gridService.Text = Format(!SumOfCharge, "0.00")
                TotalFee = TotalFee + !SumOfCharge
            End If
            .MoveNext
        Wend
    End With
    FillServices = TotalFee
End Function

Private Function FillMedicines() As Double
    Dim rsBill As New ADODB.Recordset
    Dim TotalFee As Double
    With gridMedicines
        .Cols = 3
        .Rows = 0
    End With
    With rsBill
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.Date, tblSaleCategory.SaleCategory, tblSaleBill.NetPrice " & _
                    "FROM tblSaleBill INNER JOIN tblSaleCategory ON tblSaleBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
                    "WHERE (((tblSaleBill.BilledBHTID)=" & Val(cmbBHT.BoundText) & ")) Order by tblSaleBill.SaleBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridMedicines.Rows = gridMedicines.Rows + 1
            gridMedicines.Row = gridMedicines.Rows - 1
            
            gridMedicines.Col = 0
            gridMedicines.Text = Format(!Date, "dd MMM yyyy")
            
            gridMedicines.Col = 1
            gridMedicines.Text = !SaleCategory
            
            gridMedicines.Col = 2
            gridMedicines.Text = Format(!NetPrice, "0.00")
            
            TotalFee = TotalFee + !NetPrice
            
            .MoveNext
        Wend
        
        If .State = 1 Then .Close
        temSql = "SELECT tblReturnBill.Date, tblReturnBill.NetPrice " & _
                    "From tblReturnBill " & _
                    "WHERE (((tblReturnBill.BilledBHTID)=" & Val(cmbBHT.BoundText) & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridMedicines.Rows = gridMedicines.Rows + 1
            gridMedicines.Row = gridMedicines.Rows - 1
            
            gridMedicines.Col = 0
            gridMedicines.Text = Format(!Date, "dd MMM yyyy")
            
            gridMedicines.Col = 1
            gridMedicines.Text = "Return"
            
            gridMedicines.Col = 2
            gridMedicines.Text = Format(!NetPrice, "0.00")
            
            TotalFee = TotalFee - !NetPrice
            
            .MoveNext
        Wend
        

    End With
    FillMedicines = TotalFee
End Function

Private Function FillRooms()
    FillRooms = 0
'    Dim rsRoom As New ADODB.Recordset
'    Dim RoomCharge As Double
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'
'    With gridRoom
'        .Cols = 5
'        .Rows = 0
'        .ColWidth(1) = 2000
'        .ColWidth(2) = 2000
'        .ColWidth(3) = 1000
'    End With
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If MyBHT.PaymentMethod = "Cash" Then
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge - MyRoom.DiscountForCash) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge - MyRoom.DiscountForCash) / 4)
'                    End If
'                ElseIf MyBHT.PaymentMethod = "Credit" Then
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge + MyRoom.SurchargeForCredit) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge + MyRoom.SurchargeForCredit) / 4)
'                    End If
'                Else
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge) / 4)
'                    End If
'                End If
'
'                gridRoom.Rows = gridRoom.Rows + 1
'                gridRoom.Row = gridRoom.Rows - 1
'
'                gridRoom.Col = 0
'                gridRoom.Text = !Room
'
'                gridRoom.Col = 1
'                gridRoom.Text = Format(FromDate, "dd MMM yy") & " - " & Format(FromTime, "HH:MM")
'
'                gridRoom.Col = 2
'                gridRoom.Text = Format(ToDate, "dd MMM yy") & " - " & Format(ToTime, "HH:MM")
'
'                gridRoom.Col = 3
'                gridRoom.Text = DurationHours & " hrs"
'
'                gridRoom.Col = 4
'
'                If MyBHT.Discharge = False Then
'                    gridRoom.Text = Format(DurationFee * ((100 + MyBHT.PtSurcharge) / 100), "0.00")
'                    TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'                Else
'                    gridRoom.Text = Format(DurationFee, "0.00")
'                    TotalFee = TotalFee + DurationFee
'                End If
'
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'    FillRooms = TotalFee
End Function

Private Function FAdditionalCharge() As Double
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT tblPatientCharge.* " & _
'                    "FROM tblPatientCharge " & _
'                    "WHERE Cancelled = 0   AND PatientChargeID = " & Val(cmbBHT.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            TotalValue = TotalValue + !Charge
'            .MoveNext
'        Wend
'        .Close
'    End With
'    FAdditionalCharge = TotalValue
    FAdditionalCharge = 0
End Function

Private Function FMaintananceCharges() As Double
    FMaintananceCharges = 0
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'    Dim temHours As Long
'
'    FromDate = MyBHT.DOA
'    FromTime = MyBHT.TOA
'    If MyBHT.Discharge = False Then
'        ToDate = dtpDOD.Value
'        ToTime = dtpTOD.Value
'    Else
'        ToDate = MyBHT.DOD
'        ToTime = MyBHT.TOD
'    End If
'    temHours = Abs(DateDiff("h", FromDate + FromTime, ToDate + ToTime))
'    If temHours > 1 Then
'        If MyBHT.PaymentMethod <> "Cash" Then
'            If temHours Mod 6 = 0 Then
'                FMaintananceCharges = (temHours \ 6) * (InwardPtCh.MaintananceRate / 4)
'            Else
'                FMaintananceCharges = ((temHours \ 6) + 1) * (InwardPtCh.MaintananceRate / 4)
'            End If
'        Else
'            If temHours Mod 6 = 0 Then
'                FMaintananceCharges = (temHours \ 6) * ((InwardPtCh.MaintananceRate - InwardPtCh.MaintainaceCashDiscountRate) / 4)
'            Else
'                FMaintananceCharges = ((temHours \ 6) + 1) * ((InwardPtCh.MaintananceRate - InwardPtCh.MaintainaceCashDiscountRate) / 4)
'            End If
'        End If
'    End If
'    If MyBHT.Discharge = False Then FMaintananceCharges = FMaintananceCharges * (100 + MyBHT.PtSurcharge) / 100
End Function

Private Function FNursingCharge()
    FNursingCharge = 0
'    Dim rsRoom As New ADODB.Recordset
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If MyRoom.ICUNursing = True Then
'                        If DurationHours Mod 6 = 0 Then
'                            DurationFee = DurationFee + ((DurationHours \ 6)) * ((InwardPtCh.ICUNursingRate) / 4)
'                        Else
'                            DurationFee = DurationFee + ((DurationHours \ 6) + 1) * ((InwardPtCh.ICUNursingRate) / 4)
'                        End If
'                Else
'                        If DurationHours Mod 6 = 0 Then
'                            DurationFee = DurationFee + ((DurationHours \ 6)) * ((InwardPtCh.NursingRate) / 4)
'                        Else
'                            DurationFee = DurationFee + ((DurationHours \ 6) + 1) * ((InwardPtCh.NursingRate) / 4)
'                        End If
'                End If
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'    If MyBHT.Discharge = False Then
'        TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'    Else
'        TotalFee = TotalFee + DurationFee
'    End If
'    FNursingCharge = TotalFee
'
End Function


Private Function FFillProfessionalCharges()
    With gridProfessionalF
        .Rows = 1
        .Cols = 7
        .Col = 1
        .Text = "Date"
        .Col = 2
        .Text = "Time"
        .Col = 3
        .Text = "Speciality"
        .Col = 4
        .Text = "Name"
        .Col = 5
        .Text = "Comments"
        .Col = 6
        .Text = "Value"
        .ColWidth(0) = 0
        .ColWidth(2) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 1600
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1200
    End With
    Dim TotalValue As Double
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblSpeciality.Speciality, tblTitle.Title, tblStaff.Name, tblProfessionalCharges.* " & _
                    "FROM ((tblSpeciality RIGHT JOIN tblStaff ON tblSpeciality.SpecialityID = tblStaff.SpecialityID) RIGHT JOIN tblProfessionalCharges ON tblStaff.StaffID = tblProfessionalCharges.StaffID) LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID " & _
                    "WHERE (((tblProfessionalCharges.ProfessionalCharge)=1) AND ((tblProfessionalCharges.Cancelled)=0) AND ((tblProfessionalCharges.ForBHTID)=" & Val(cmbBHT.BoundText) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            gridProfessionalF.Rows = gridProfessionalF.Rows + 1
            gridProfessionalF.Row = gridProfessionalF.Rows - 1
            
            gridProfessionalF.Col = 0
            gridProfessionalF.Text = !ProfessionalChargesID
            
            gridProfessionalF.Col = 1
            gridProfessionalF.Text = Format(!Date, "dd MMM yyyy")
            
            gridProfessionalF.Col = 2
            gridProfessionalF.Text = Format(!Time, "HH MM")
            
            gridProfessionalF.Col = 3
            gridProfessionalF.Text = !Speciality
            
            gridProfessionalF.Col = 4
            gridProfessionalF.Text = !Title & " " & !Name
            
            gridProfessionalF.Col = 5
            gridProfessionalF.Text = !Comments
            
            gridProfessionalF.Col = 6
            gridProfessionalF.Text = Format(!Fee, "0.00")
            
            TotalValue = TotalValue + !Fee
            
            .MoveNext
        
        Wend
        .Close
    End With
    FFillProfessionalCharges = TotalValue
End Function

Private Function FLinanCharges()
    FLinanCharges = 0
'    Dim rsRoom As New ADODB.Recordset
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If DurationHours > 3 * 24 Then
'                    If (DurationHours - (3 * 24)) Mod 6 = 0 Then
'                        DurationFee = DurationFee + InwardPtCh.InitialLinanRate + (((DurationHours - (3 * 24)) \ 6) * (InwardPtCh.LaterLinanRate / 4))
'                    Else
'                        DurationFee = DurationFee + InwardPtCh.InitialLinanRate + (((DurationHours - ((3 * 24)) \ 6) + 1) * (InwardPtCh.LaterLinanRate / 4))
'                    End If
'                Else
'                    DurationFee = DurationFee + InwardPtCh.InitialLinanRate
'                End If
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'    If MyBHT.Discharge = False Then
'        TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'    Else
'        TotalFee = TotalFee + DurationFee
'    End If
'    FLinanCharges = TotalFee
End Function

Private Function FFillPayments()
    
    Dim TotalPayments As Double
    Dim rsTem As New ADODB.Recordset
    With gridPaymentsF
        .Cols = 3
        .Rows = 0
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsGSBill = 1 AND Cancelled = 0  AND BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPaymentsF.Rows = gridPaymentsF.Rows + 1
            gridPaymentsF.Row = gridPaymentsF.Rows - 1
            gridPaymentsF.Col = 0
            gridPaymentsF.Text = Format(!Date, "dd MMMM yyyy")
            gridPaymentsF.Col = 1
            'gridPaymentsF.Text = !IncomeBillID
            gridPaymentsF.Text = !DisplayBillID
            
            gridPaymentsF.Col = 2
            gridPaymentsF.Text = Format(!NetTotal, "0.00")
            TotalPayments = TotalPayments + !NetTotal
            
            .MoveNext
        Wend
        .Close
    End With
    
    With rsTem
        If .State = 1 Then .Close
'        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsInwardPaymentBill = 1 AND Cancelled = 0  AND BHTID = " & Val(cmbBHT.BoundText)
        temSql = "SELECT tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime, dbo.tblIncomeBill.PaymentComments " & _
                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (dbo.tblIncomeBill.IsHSSPaymentBill = 1) AND (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.BHTID = " & Val(cmbBHT.BoundText) & ") AND (dbo.tblIncomeBill.Cancelled = 0) " & _
                    "ORDER BY dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            gridPayments.Rows = gridPayments.Rows + 1
            gridPayments.Row = gridPayments.Rows - 1
            gridPayments.Col = 0
            gridPayments.Text = Format(!CompletedDate, "dd MMMM yyyy")
            gridPayments.Col = 2
            gridPayments.Text = Format(!NetTotal, "0.00")
            gridPayments.Col = 1
            gridPayments.Text = !DisplayBillID
            TotalPayments = TotalPayments + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeReturnBill.* FROM tblIncomeReturnBill WHERE tblIncomeReturnBill.BHTID =" & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPaymentsF.Rows = gridPaymentsF.Rows + 1
            gridPaymentsF.Row = gridPaymentsF.Rows - 1
            gridPaymentsF.Col = 0
            gridPaymentsF.Text = Format(!ReturnDate, "dd MMMM yyyy")
            gridPaymentsF.Col = 1
            gridPaymentsF.Text = !IncomeReturnBillID
            gridPaymentsF.Col = 2
            gridPaymentsF.Text = Format(!ReturnValue, "0.00")
            TotalPayments = TotalPayments - !ReturnValue
            .MoveNext
        Wend
        .Close
    End With
    
    
    FFillPayments = TotalPayments
End Function

Private Function FFillServices() As Double
    Dim TotalFee As Double
    Dim CatFee As Double
    Dim rsTem As New ADODB.Recordset
    
    With gridServiceF
        .Clear
        .Cols = 2
        .Rows = 0
        .ColWidth(0) = 3600
        .ColWidth(1) = .Width - .ColWidth(0) - 150
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblServiceCategory.ServiceCategory, tblServiceCategory.InwardSurcharge, Sum(tblPatientService.Charge) AS SumOfCharge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.AddToMedicineCharge)=0) AND ((tblPatientService.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
                    "GROUP BY tblServiceCategory.ServiceCategory, tblServiceCategory.InwardSurcharge"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridServiceF.Rows = gridServiceF.Rows + 1
            gridServiceF.Row = gridServiceF.Rows - 1
            gridServiceF.Col = 0
            gridServiceF.Text = !ServiceCategory
            gridServiceF.Col = 1
            CatFee = !SumOfCharge
'            If IsNull(!InwardSurcharge) = False Then
'                CatFee = CatFee + (CatFee * !InwardSurcharge / 100)
'            End If
            
            If MyBHT.Discharge = False Then
                CatFee = (CatFee * (100 + MyBHT.PtSurcharge) / 100)
            End If
            TotalFee = TotalFee + CatFee
            gridServiceF.Text = Format(CatFee, "0.00")
            
'            If MyBHT.Discharge = False Then
'                gridServiceF.Text = Format(CatFee * (100 + MyBHT.PtSurcharge) / 100, "0.00")
'                TotalFee = TotalFee + (CatFee * (100 + MyBHT.PtSurcharge) / 100)
'            Else
'                gridServiceF.Text = Format(CatFee, "0.00")
'                TotalFee = TotalFee + CatFee
'            End If
            .MoveNext
        Wend
    End With
    FFillServices = TotalFee
End Function

Private Sub ClearValues()
    txtDetails.Text = Empty
    
    lblMedicineCharge.Caption = "0.00"
    lblServiceCharge.Caption = "0.00"
    lblProfessionalCharges.Caption = "0.00"
    
    lblTotalCharge.Caption = "0.00"
    lblPayments.Caption = "0.00"
    lblBalance.Caption = "0.00"
    
    gridMedicines.Clear
    gridPayments.Clear
    gridProfessional.Clear
    gridService.Clear

    lblMedicineChargeF.Caption = "0.00"
    lblServiceChargeF.Caption = "0.00"
    lblProfessionalChargesF.Caption = "0.00"
    
    lblTotalChargeF.Caption = "0.00"
    lblPaymentsF.Caption = "0.00"
    lblBalanceF.Caption = "0.00"
    
    gridMedicinesF.Clear
    gridPaymentsF.Clear
    gridProfessionalF.Clear
    gridServiceF.Clear

    
End Sub

Private Function FFillMedicines() As Double
    Dim rsBill As New ADODB.Recordset
    Dim TotalFee As Double
    With gridMedicinesF
        .Clear
        .Cols = 3
        .Rows = 0
    End With
    With rsBill
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.Date, tblSaleCategory.SaleCategory, tblSaleBill.NetPrice " & _
                    "FROM tblSaleBill INNER JOIN tblSaleCategory ON tblSaleBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
                    "WHERE (((tblSaleBill.BilledBHTID)=" & Val(cmbBHT.BoundText) & ")) Order by tblSaleBill.SaleBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridMedicinesF.Rows = gridMedicinesF.Rows + 1
            gridMedicinesF.Row = gridMedicinesF.Rows - 1
            
            gridMedicinesF.Col = 0
            gridMedicinesF.Text = Format(!Date, "dd MMM yyyy")
            
            gridMedicinesF.Col = 1
            gridMedicinesF.Text = !SaleCategory
            
            gridMedicinesF.Col = 2
            gridMedicinesF.Text = Format(!NetPrice, "0.00")
            
            TotalFee = TotalFee + !NetPrice
            
            .MoveNext
        Wend
        
        
        
        If .State = 1 Then .Close
        temSql = "SELECT tblServiceCategory.ServiceCategory, Sum(tblPatientService.Charge) AS SumOfCharge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.AddToMedicineCharge)=1) AND ((tblPatientService.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
                    "GROUP BY tblServiceCategory.ServiceCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridMedicinesF.Rows = gridMedicinesF.Rows + 1
            gridMedicinesF.Row = gridMedicinesF.Rows - 1
            gridMedicinesF.Col = 1
            gridMedicinesF.Text = !ServiceCategory
            gridMedicinesF.Col = 2
            If MyBHT.Discharge = False Then
                gridMedicinesF.Text = Format(!SumOfCharge * (100 + MyBHT.PtSurcharge) / 100, "0.00")
                TotalFee = TotalFee + (!SumOfCharge * (100 + MyBHT.PtSurcharge) / 100)
            Else
                gridMedicinesF.Text = Format(!SumOfCharge, "0.00")
                TotalFee = TotalFee + !SumOfCharge
            End If
            .MoveNext
        Wend
        
        If .State = 1 Then .Close
        temSql = "SELECT tblReturnBill.Date, tblReturnBill.NetPrice " & _
                    "From tblReturnBill " & _
                    "WHERE (((tblReturnBill.BilledBHTID)=" & Val(cmbBHT.BoundText) & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridMedicinesF.Rows = gridMedicinesF.Rows + 1
            gridMedicinesF.Row = gridMedicinesF.Rows - 1
            
            gridMedicinesF.Col = 0
            gridMedicinesF.Text = Format(!Date, "dd MMM yyyy")
            
            gridMedicinesF.Col = 1
            gridMedicinesF.Text = "Return"
            
            gridMedicinesF.Col = 2
            gridMedicinesF.Text = Format(!NetPrice, "0.00")
            
            TotalFee = TotalFee - !NetPrice
            
            .MoveNext
        Wend


    End With
    FFillMedicines = TotalFee
End Function

Private Function FFillRooms()
    FFillRooms = 0
'    Dim rsRoom As New ADODB.Recordset
'    Dim RoomCharge As Double
'
'    Dim FromDate As Date
'    Dim FromTime As Date
'    Dim ToDate As Date
'    Dim ToTime As Date
'
'    Dim TotalFee As Double
'    Dim DurationFee As Double
'    Dim DurationHours As Long
'    Dim MyRoom As New clsRoom
'    Dim MyBHT As New clsBHT
'
'    MyBHT.BHTID = Val(cmbBHT.BoundText)
'
'
'    With gridRoomF
'        .Cols = 5
'        .Rows = 0
'        .ColWidth(1) = 2000
'        .ColWidth(2) = 2000
'        .ColWidth(3) = 1000
'    End With
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRoom.Room, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash, tblRoomPatient.RoomID, tblRoomPatient.FromDate, tblRoomPatient.FromTime, tblRoomPatient.ToDate, tblRoomPatient.ToTime, tblRoomPatient.RoomPatientID " & _
'                    "FROM (tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID) RIGHT JOIN tblRoomPatient ON tblRoom.RoomID = tblRoomPatient.RoomID " & _
'                    "Where (((tblRoomPatient.BHTID) = " & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblRoomPatient.RoomPatientID"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            MyRoom.RoomID = !RoomID
'            FromDate = !FromDate
'            FromTime = !FromTime
'            If !ToDate <> Empty And !ToTime <> Empty Then
'                ToDate = !ToDate
'                ToTime = !ToTime
'            Else
'                ToDate = dtpDOD.Value
'                ToTime = dtpTOD.Value
'            End If
'            DurationHours = DateDiff("h", FromDate + FromTime, ToDate + ToTime)
'            If DurationHours >= 1 Then
'                If MyBHT.PaymentMethod = "Cash" Then
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge - MyRoom.DiscountForCash) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge - MyRoom.DiscountForCash) / 4)
'                    End If
'                ElseIf MyBHT.PaymentMethod = "Credit" Then
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge + MyRoom.SurchargeForCredit) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge + MyRoom.SurchargeForCredit) / 4)
'                    End If
'                Else
'                    If DurationHours Mod 6 = 0 Then
'                        DurationFee = ((DurationHours \ 6)) * ((MyRoom.GeneralCharge) / 4)
'                    Else
'                        DurationFee = ((DurationHours \ 6) + 1) * ((MyRoom.GeneralCharge) / 4)
'                    End If
'                End If
'
'                gridRoomF.Rows = gridRoomF.Rows + 1
'                gridRoomF.Row = gridRoomF.Rows - 1
'
'                gridRoomF.Col = 0
'                gridRoomF.Text = !Room
'
'                gridRoomF.Col = 1
'                gridRoomF.Text = Format(FromDate, "dd MMM yy") & " - " & Format(FromTime, "HH:MM")
'
'                gridRoomF.Col = 2
'                gridRoomF.Text = Format(ToDate, "dd MMM yy") & " - " & Format(ToTime, "HH:MM")
'
'                gridRoomF.Col = 3
'                gridRoomF.Text = DurationHours & " hrs"
'
'                gridRoomF.Col = 4
'
'                If MyBHT.Discharge = False Then
'                    gridRoomF.Text = Format(DurationFee * ((100 + MyBHT.PtSurcharge) / 100), "0.00")
'                    TotalFee = TotalFee + DurationFee * ((100 + MyBHT.PtSurcharge) / 100)
'                Else
'                    gridRoomF.Text = Format(DurationFee, "0.00")
'                    TotalFee = TotalFee + DurationFee
'                End If
'
'            End If
'            .MoveNext
'        Wend
'        .Close
'    End With
'    FFillRooms = TotalFee
End Function

Private Sub cmbBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnProcess_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmbBHT.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call PopulatePapers
    Call GetSettings
    
End Sub

Private Sub FillCombos()
    Dim BHT As New clsFillCombos
    BHT.FillBoolCombo cmbBHT, "BHT", "BHT", "IsGSB", False
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDOD.Value = Date
    dtpTOD.Value = Time
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1)
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    SSTab1.Tab = GetSetting(App.EXEName, Me.Name, SSTab1.Name, "0")
    SSTab2.Tab = GetSetting(App.EXEName, Me.Name, SSTab2.Name, "0")
    SSTab3.Tab = GetSetting(App.EXEName, Me.Name, SSTab3.Name, "0")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
    SaveSetting App.EXEName, Me.Name, SSTab2.Name, SSTab2.Tab
    SaveSetting App.EXEName, Me.Name, SSTab3.Name, SSTab3.Tab
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

Private Sub txtDiscount_Change()
    txtDiscountF.Text = Format(txtDiscount.Text, "0.00")
End Sub

Private Sub txtDiscount_LostFocus()
    Dim rsTem As New ADODB.Recordset
    If Val(txtDiscount.Text) <> MyBHT.Discount Then
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblBHT where BHTID = " & MyBHT.BHTID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Price = FPtTotalCharge
                !Discount = Val(txtDiscount.Text)
                !DiscountPercent = Val(txtDiscount.Text) / FPtTotalCharge * 100
                !NetPrice = FPtTotalCharge - Val(txtDiscount.Text)
                .Update
            End If
            .Close
        End With
    End If
    MyBHT.BHTID = MyBHT.BHTID
    lblBalance.Caption = Format(PtBalance, "0.00")
End Sub

Private Sub txtDiscountF_LostFocus()
    Dim rsTem As New ADODB.Recordset
    If Val(txtDiscountF.Text) <> MyBHT.Discount Then
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblBHT where BHTID = " & MyBHT.BHTID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Price = FPtTotalCharge
                !Discount = Val(txtDiscount.Text)
                !DiscountPercent = Val(txtDiscount.Text) / FPtTotalCharge * 100
                !NetPrice = FPtTotalCharge - Val(txtDiscount.Text)
                .Update
            End If
            .Close
        End With
    End If
    MyBHT.BHTID = MyBHT.BHTID
    lblBalanceF.Caption = Format(FPtBalance, "0.00")
    
    
    Call ActualCalculations
    Call FakeCalculations

End Sub

Private Sub txtDiscountF_Change()
    txtDiscount.Text = Format(txtDiscountF.Text, "0.00")
End Sub

