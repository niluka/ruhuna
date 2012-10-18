VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDoctorsIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment For Doctors"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDoctorsIncome.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   9945
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   8520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   9255
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optPaid 
         Caption         =   "Paid"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optTopay 
         Caption         =   "To Pay"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   5295
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9340
         _Version        =   393216
         BackColorFixed  =   -2147483637
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   14420
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "frmDoctorsIncome.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDate"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmDoctorsIncome.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Selected Period"
      TabPicture(2)   =   "frmDoctorsIncome.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPicker2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPicker3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   39472
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   39472
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -71640
         TabIndex        =   9
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   39472
      End
      Begin VB.Label lblDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71880
         TabIndex        =   13
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   -71640
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Date From"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Date To"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDoctorsIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblDoctor.Doctor_ID, tblDoctor.DoctorName FROM tblDoctor RIGHT JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID"
Const PostSHape = "} AS cmmdDuctorsPayment COMPUTE cmmdDuctorsPayment, SUM(cmmdDuctorsPayment.'PersonalFee') AS TotalPersonalFee, SUM(cmmdDuctorsPayment.'PersonalRefund') AS PerRefund, SUM(cmmdDuctorsPayment.'PersonalDue') AS TotalDueAmount BY 'Doctor_ID','DoctorName' "

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Call FormatGrid1
Call FindDoctor
End Sub

Private Sub DTPicker2_Change()
Call FormatGrid1
Call FindDoctor
End Sub

Private Sub DTPicker3_Change()
Call FormatGrid1
Call FindDoctor
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date

lblDate.Caption = Date
SSTab1.Tab = 0
Call FormatGrid1
Call FindDoctor
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If

End Sub

Private Sub FindDoctor()
Dim i As Long
Dim r As Long
Dim TemTotal As Long
Dim TemRefund As Long
TemTotal = 0

    With DataEnvironment1
        If .rscmmdDuctorsPayment_Grouping.State = 1 Then .rscmmdDuctorsPayment_Grouping.Close
       
       Select Case SSTab1.Tab
       
        Case 0
        .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & Date & "') and (fullypaid = 1) and (paidtostaff = 0)" & PostSHape 'and (paidtostaff = 0)
         .cmmdDuctorsPayment_Grouping
        Case 1
        .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & DTPicker1.Value & "') and (fullypaid = 1) and (paidtostaff = 0)" & PostSHape
         .cmmdDuctorsPayment_Grouping
        Case 2
        .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate Between  '" & DTPicker2.Value & "' and  '" & DTPicker3.Value & "' ) and (fullypaid = 1) and (paidtostaff = 0) " & PostSHape
         .cmmdDuctorsPayment_Grouping
        End Select
        i = 1
        r = 1
        
        If .rscmmdDuctorsPayment_Grouping.RecordCount = 0 Then Exit Sub
        Do While .rscmmdDuctorsPayment_Grouping.EOF = False
        r = r + 1
        Grid1.Rows = r
        Grid1.Row = i
        Grid1.Col = 0
        Grid1.Text = .rscmmdDuctorsPayment_Grouping!doctorname
        Grid1.Col = 1
        Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!TotalDueAmount, "0.00")
        Grid1.CellAlignment = 7
        TemTotal = TemTotal + Val(.rscmmdDuctorsPayment_Grouping!TotalDueAmount)
        Grid1.Col = 2
        Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!TotalPersonalFee, "0.00")
        Grid1.CellAlignment = 7
        Grid1.Col = 3
        If Not .rscmmdDuctorsPayment_Grouping!PerRefund = "" Then Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!PerRefund, "0.00")
        Grid1.CellAlignment = 7
        i = i + 1
        
        .rscmmdDuctorsPayment_Grouping.MoveNext
        Loop
        Grid1.Rows = r + 1
        Grid1.Row = i
        Grid1.Col = 0
        Grid1.Text = "Total"
        Grid1.Col = 1
        Grid1.Text = Format(TemTotal, "0.00")
    End With
End Sub

Private Sub FormatGrid1()
With Grid1
    .Cols = 4
    .ColWidth(0) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + 350)
    .Row = 0
    .Col = 0
    .Text = "Doctor Name"
    .Col = 1
    .Text = "Due Amount"
    .CellAlignment = 7
    .Col = 2
    .Text = "Total Income"
    .CellAlignment = 7
    .Col = 3
    .Text = "Total Refund"
    .CellAlignment = 7
    .Rows = 1
End With

End Sub

Private Sub FindDoctorbyOption()
Dim i As Long
Dim r As Long
Dim TemTotal As Long
Dim TemRefund As Long
TemTotal = 0

    With DataEnvironment1
        If .rscmmdDuctorsPayment_Grouping.State = 1 Then .rscmmdDuctorsPayment_Grouping.Close
       
       Select Case SSTab1.Tab
       
        Case 0
            If optTopay.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & Date & "') and (fullypaid = 1) and (paidtostaff = 0)" & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If
            If optPaid.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & Date & "') and (fullypaid = 1) and (PaidToSTaff = True)" & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If
            If optAll.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & Date & "') and (fullypaid = 1)" & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If
        
        Case 1
           If optTopay.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & DTPicker1.Value & "') and (fullypaid = 1) and (paidtostaff = 0)" & PostSHape
                .cmmdDuctorsPayment_Grouping
           End If
           If optPaid.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & DTPicker1.Value & "') and (fullypaid = 1) and (PaidToSTaff = True)" & PostSHape
                .cmmdDuctorsPayment_Grouping
           End If
           If optAll.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & DTPicker1.Value & "') and (fullypaid = 1) " & PostSHape
                .cmmdDuctorsPayment_Grouping
           End If

        Case 2
            If optTopay.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate Between  '" & DTPicker2.Value & "' and  '" & DTPicker3.Value & "' ) and (fullypaid = 1) and (paidtostaff = 0) " & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If
            If optPaid.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate Between  '" & DTPicker2.Value & "' and  '" & DTPicker3.Value & "' ) and (fullypaid = 1) and (PaidToSTaff = True) " & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If
            If optAll.Value = True Then
                .Commands!cmmdDuctorsPayment_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate Between  '" & DTPicker2.Value & "' and  '" & DTPicker3.Value & "' ) and (fullypaid = 1)" & PostSHape
                .cmmdDuctorsPayment_Grouping
            End If

        End Select
        i = 1
        r = 1
        FormatGrid1
        If .rscmmdDuctorsPayment_Grouping.RecordCount = 0 Then Exit Sub
        Do While .rscmmdDuctorsPayment_Grouping.EOF = False
        r = r + 1
        Grid1.Rows = r
        Grid1.Row = i
        Grid1.Col = 0
        Grid1.Text = .rscmmdDuctorsPayment_Grouping!doctorname
        Grid1.Col = 1
        Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!TotalDueAmount, "0.00")
        Grid1.CellAlignment = 7
        TemTotal = TemTotal + Val(.rscmmdDuctorsPayment_Grouping!TotalDueAmount)
        Grid1.Col = 2
        Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!TotalPersonalFee, "0.00")
        Grid1.CellAlignment = 7
        Grid1.Col = 3
        If Not .rscmmdDuctorsPayment_Grouping!PerRefund = "" Then Grid1.Text = Format(.rscmmdDuctorsPayment_Grouping!PerRefund, "0.00")
        Grid1.CellAlignment = 7
        i = i + 1
        
        .rscmmdDuctorsPayment_Grouping.MoveNext
        Loop
        Grid1.Rows = r + 1
        Grid1.Row = i
        Grid1.Col = 0
        Grid1.Text = "Total"
        Grid1.Col = 1
        Grid1.Text = Format(TemTotal, "0.00")
    End With
End Sub


Private Sub optAll_Click()
Call FindDoctorbyOption
End Sub

Private Sub optPaid_Click()
Call FindDoctorbyOption
End Sub

Private Sub optTopay_Click()
Call FindDoctorbyOption
End Sub

Private Sub SSTab1_DblClick()
Call FormatGrid1
Call FindDoctor
End Sub
