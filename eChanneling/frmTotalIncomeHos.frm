VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTotalIncomeHos 
   Caption         =   "Total Income"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Bindings        =   "frmTotalIncomeHos.frx":0000
      Height          =   5175
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      DataMember      =   "ccmmdPatientFacilities_Grouping"
      _NumberOfBands  =   3
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   1
      _Band(0)._MapCol(0)._Name=   "HospitalFacility_ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   7
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   7
      _Band(1)._MapCol(0)._Name=   "BookingDate"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(1)._Name=   "InstitutionFee"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(1)._Alignment=   7
      _Band(1)._MapCol(2)._Name=   "PersonalFee"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(2)._Alignment=   7
      _Band(1)._MapCol(3)._Name=   "PersonalRefund"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(1)._MapCol(3)._Alignment=   7
      _Band(1)._MapCol(4)._Name=   "TotalFee"
      _Band(1)._MapCol(4)._RSIndex=   4
      _Band(1)._MapCol(4)._Alignment=   7
      _Band(1)._MapCol(5)._Name=   "TotalRefund"
      _Band(1)._MapCol(5)._RSIndex=   5
      _Band(1)._MapCol(5)._Alignment=   7
      _Band(1)._MapCol(6)._Name=   "HospitalFacility_ID"
      _Band(1)._MapCol(6)._RSIndex=   6
      _Band(1)._MapCol(6)._Alignment=   7
      _Band(2).BandIndent=   2
      _Band(2).Cols   =   71
      _Band(2).GridLinesBand=   1
      _Band(2).TextStyleBand=   0
      _Band(2).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmTotalIncomeHos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'With Grid1
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'.ColWidth(1, 0) = 800
'
'End With
End Sub
