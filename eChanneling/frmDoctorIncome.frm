VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDoctorIncome 
   Caption         =   "Total Doctor Payment"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6975
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12303
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPrintView 
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print &View"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   14208
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Toda&y"
      TabPicture(0)   =   "frmDoctorIncome.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblToday"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Selected Day"
      TabPicture(1)   =   "frmDoctorIncome.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DTPicker1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "S&elected Period"
      TabPicture(2)   =   "frmDoctorIncome.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label18"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPicker3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -70800
         TabIndex        =   4
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161021955
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161021955
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161021955
         CurrentDate     =   39442
      End
      Begin VB.Label Label18 
         Caption         =   "&To"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "&From"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Selec&ted Date"
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
         Left            =   -74040
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         Caption         =   "Today :"
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
         Left            =   -73920
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmDoctorIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub PrintView()
End Sub
Private Sub Setcolours()


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

'bttnCashIncome.BackColor = BttnBackColour
'bttnCashIncome.ForeColor = BttnForeColour
'
'bttnPrintAgentCash.BackColor = BttnBackColour
'bttnPrintAgentCash.ForeColor = BttnForeColour
'
'bttnCashFromCreditChanneling.BackColor = BttnBackColour
'bttnCashFromCreditChanneling.ForeColor = BttnForeColour
'
'bttnCashRefunds.BackColor = BttnBackColour
'bttnCashRefunds.ForeColor = BttnForeColour
'
'bttnDoctorPayments.BackColor = BttnBackColour
'bttnDoctorPayments.ForeColor = BttnForeColour

bttnPrintView.BackColor = BttnBackColour
bttnPrintView.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour
'
'bttnDelete.BackColor = BttnBackColour
'bttnDelete.ForeColor = BttnForeColour


'FrameShiftSummary.BackColor = FrameBackColour
'FrameShiftSummary.ForeColor = FrameForeColour




frmDoctorIncome.BackColor = FrameBackColour
frmDoctorIncome.ForeColor = FrameForeColour

'FrameOfficial.BackColor = FrameBackColour
'FrameOfficial.ForeColor = FrameForeColour
'
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'chkBypassOrder.BackColor = LblBackColour
'chkBypassOrder.ForeColor = LblForeColour
'
'ChkCalculateTime.BackColor = LblBackColour
'ChkCalculateTime.ForeColor = LblForeColour
'
'chkFullDayLeave.BackColor = LblBackColour
'chkFullDayLeave.ForeColor = LblForeColour
'
'
'DataComboDoctor.BackColor = TxtBackColour
'DataComboDoctor.ForeColor = TxtForeColour

'DataComboPaymenyMethod.BackColor = TxtBackColour
'DataComboPaymenyMethod.ForeColor = TxtForeColour
'
'DataComboSex.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour
'
'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




'Grid1.BackColor = GridBackColor
'Grid1.ForeColor = GridForeColor
'
'Grid1.BackColorBkg = GridBackColorBkg
'Grid1.BackColorFixed = GridBackColorFixed
'Grid1.BackColorSel = GridBackColorSel
'
'Grid1.ForeColor = GridForeColor
'Grid1.ForeColorFixed = GridForeColorFixed
'Grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid



'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour
'
'Label10.BackColor = LblBackColour
'Label10.ForeColor = LblForeColour
'Label11.BackColor = LblBackColour
'Label11.ForeColor = LblForeColour
'Label12.BackColor = LblBackColour
'Label12.ForeColor = LblForeColour
'Label13.BackColor = LblBackColour
'Label13.ForeColor = LblForeColour
'Label14.BackColor = LblBackColour
'Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
'Label16.BackColor = LblBackColour
'Label16.ForeColor = LblForeColour
'Label2.BackColor = LblBackColour
'Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
'Label3.BackColor = LblBackColour
'Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
'Label26.BackColor = LblBackColour
'Label26.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label5.BackColor = LblBackColour
'Label5.ForeColor = LblForeColour
'Label6.BackColor = LblBackColour
'Label6.ForeColor = LblForeColour
'Label7.BackColor = LblBackColour
'Label7.ForeColor = LblForeColour

'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour

'lblOfficialEmail.BackColor = LblBackColour
'lblOfficialEmail.ForeColor = LblForeColour

'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


'txtAccount.BackColor = TxtBackColour
'txtAccount.ForeColor = TxtForeColour
'
'txtBankBranch.BackColor = TxtBackColour
'txtBankBranch.ForeColor = TxtForeColour
'
'txtComments.BackColor = TxtBackColour
'txtComments.ForeColor = TxtForeColour
'txtCredit.BackColor = TxtBackColour
'txtCredit.ForeColor = TxtForeColour
'txtDesignation.BackColor = TxtBackColour
'txtDesignation.ForeColor = TxtForeColour
'txtListedName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour
'txtName.BackColor = TxtBackColour
'txtName.ForeColor = TxtForeColour
'txtOfficialAddress.BackColor = TxtBackColour
'txtOfficialAddress.ForeColor = TxtForeColour
'txtOfficialEMail.BackColor = TxtBackColour
'txtOfficialEMail.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour
'
'txtPrivateAddress.BackColor = TxtBackColour
'txtPrivateAddress.ForeColor = TxtForeColour
'txtPrivateEmail.BackColor = TxtBackColour
'txtPrivateEmail.ForeColor = TxtForeColour
'txtPrivateFax.BackColor = TxtBackColour
'txtPrivateFax.ForeColor = TxtForeColour
'txtPrivateMobile.BackColor = TxtBackColour
'txtPrivateMobile.ForeColor = TxtForeColour
'txtPrivateTel.BackColor = TxtBackColour
'txtPrivateTel.ForeColor = TxtForeColour
'
'
'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
'txtSearch.BackColor = TxtBackColour
'txtSearch.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour







End Sub
Private Sub bttnPrintView_Click()

With DataReportDoctorIncome

If HospitalDetails = True Then
    .Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    .Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    .Sections("repotrFooter").Controls.Item("lblAds").Caption = LongAd
Else
    .Sections("ReportHeader").Controls.Item("RptName").Caption = Empty
    .Sections("ReportHeader").Controls.Item("RptAddress").Caption = Empty
    .Sections("repotrFooter").Controls.Item("lblAds").Caption = LongAd
End If

 Select Case SSTab1.Tab
 
 Case 0
    .Sections("PageHeader").Controls.Item("rptlFrom").Caption = Format(Date, DefaultLongDate)
    .Sections("PageHeader").Controls.Item("rptlTo").Caption = Format(Date, DefaultLongDate)

 Case 1
    .Sections("PageHeader").Controls.Item("rptlFrom").Caption = DTPicker1.Value
    .Sections("PageHeader").Controls.Item("rptlTo").Caption = DTPicker1.Value
 Case 2
    .Sections("PageHeader").Controls.Item("rptlFrom").Caption = DTPicker2.Value
    .Sections("PageHeader").Controls.Item("rptlTo").Caption = DTPicker3.Value
 
 End Select
 

    DataReportDoctorIncome.Show
    
End With

End Sub


Private Sub FormatGrid()
    With Grid1
        .Cols = 5
        
        .ColWidth(0) = 600
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + 100)
        
        .Row = 0
        
        .col = 0
        .CellAlignment = 4
        .Text = "No."
        
        .col = 1
        .CellAlignment = 4
        .Text = "Doctor Name"
        
        .col = 2
        .CellAlignment = 4
        .Text = "Doctor Fee"
        
        .col = 3
        .CellAlignment = 4
        .Text = "Repayments"
        
        .col = 4
        .CellAlignment = 4
        .Text = "Net Total"
        
        
    
    End With
End Sub

Private Sub CalculateDoctorIncome()
        Dim TemDocId As Long
        Dim TemDocFee As Double
        Dim TemDocRefund As Double
        Dim NowROw As Long
        Dim GrandDocFee As Double
        Dim GrandDocRefund As Double
        
        Grid1.Visible = False
        Me.MousePointer = vbHourglass

With DataEnvironment1.rssqlTem17
    If .State = 1 Then .Close
    .Source = "select * from tbltemdoctorincome"
    .Open
        On Error Resume Next
        While .EOF = False
            .Delete adAffectCurrent
            .MoveNext
        Wend
    End With

With DataEnvironment1.rssqlTem16
    If .State = 1 Then .Close
    .Source = "Select * from tbldoctor order by doctorname"
    .Open
    If .RecordCount = 0 Then Exit Sub
    While .EOF = False
        TemDocId = !Doctor_ID
        NowROw = NowROw + 1
        Grid1.Rows = NowROw + 1
        Grid1.Row = NowROw
        Grid1.col = 0
        Grid1.Text = NowROw
        Grid1.col = 1
        Grid1.Text = !doctorname
        DataEnvironment1.rssqlTem17.AddNew
        DataEnvironment1.rssqlTem17!doctorname = !doctorname
        DataEnvironment1.rssqlTem17!Id = NowROw
        With DataEnvironment1.rssqlTem15
                If .State = 1 Then .Close
                Select Case SSTab1.Tab
                Case 0:
                    .Source = "select * from tblpatientfacility where staff_ID = " & TemDocId & " and BookingDate = '" & Date & "' and PersonalFee > 0 "
                Case 1:
                    .Source = "select * from tblpatientfacility where staff_ID = " & TemDocId & " and BookingDate = '" & DTPicker1.Value & "' and PersonalFee > 0 "
                Case 2:
                    .Source = "select * from tblpatientfacility where staff_ID = " & TemDocId & " and (BookingDate between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and PersonalFee > 0 "
                End Select
                .Open
                TemDocFee = 0
                TemDocRefund = 0
                If .RecordCount <> 0 Then
                    While .EOF = False
                        TemDocFee = TemDocFee + !PersonalFee
                        If Not IsNull(!Personalrefund) Then TemDocRefund = TemDocRefund + !Personalrefund
                        .MoveNext
                    Wend
                End If
                DataEnvironment1.rssqlTem17!DoctorFee = TemDocFee
                DataEnvironment1.rssqlTem17!doctorrefund = TemDocRefund
                DataEnvironment1.rssqlTem17!doctornetincome = TemDocFee - TemDocRefund
                DataEnvironment1.rssqlTem17.Update
                GrandDocFee = GrandDocFee + TemDocFee
                GrandDocRefund = GrandDocRefund + TemDocRefund
                Grid1.col = 2
                Grid1.CellAlignment = 7
                Grid1.Text = Format(TemDocFee, "0.00")
                Grid1.col = 3
                Grid1.CellAlignment = 7
                Grid1.Text = Format(TemDocRefund, "0.00")
                Grid1.col = 4
                Grid1.CellAlignment = 7
                Grid1.Text = Format(TemDocFee - TemDocRefund, "0.00")
        End With
    .MoveNext
    Wend
    .Close
End With

DataEnvironment1.rssqlTem17.Close
With Grid1
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .col = 1
    .Text = "Total"
    .col = 2
    .Text = Format(GrandDocFee, "0.00")
    .col = 3
    .Text = Format(GrandDocRefund, "0.00")
    .col = 4
    .Text = Format(GrandDocFee - GrandDocRefund, "0.00")
    .Visible = True
    
    Me.MousePointer = vbDefault
End With
End Sub

Private Sub DTPicker1_Change()
    Call FormatGrid
    Call CalculateDoctorIncome
End Sub

Private Sub DTPicker2_Change()
    Call FormatGrid
    Call CalculateDoctorIncome
End Sub

Private Sub DTPicker3_Change()
    Call FormatGrid
    Call CalculateDoctorIncome
End Sub


Private Sub Form_Load()
    Call FormatGrid
    Call CalculateDoctorIncome
    Call Setcolours
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker3 = Date
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call FormatGrid
    Call CalculateDoctorIncome
End Sub

