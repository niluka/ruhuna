VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCancellationAgentPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation Agent Payment"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancellationAgentPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8475
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cancel Bill"
      TabPicture(0)   =   "frmCancellationAgentPayments.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search Bill By Amount"
      TabPicture(1)   =   "frmCancellationAgentPayments.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dtpTo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dtpFrom"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtSearchAmount"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btnSearchByAmount"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "gridBills"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin MSFlexGridLib.MSFlexGrid gridBills 
         Height          =   4695
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx btnSearchByAmount 
         Height          =   375
         Left            =   5760
         TabIndex        =   23
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Search"
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
      Begin VB.TextBox txtSearchAmount 
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   1560
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58851331
         CurrentDate     =   40669
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   2
         Top             =   1080
         Width           =   7695
         Begin VB.TextBox txtReceiptNo 
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   600
            Width           =   2655
         End
         Begin VB.Frame Frame2 
            Height          =   2175
            Left            =   480
            TabIndex        =   4
            Top             =   1680
            Width           =   6855
            Begin VB.Label lblUserName 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   3360
               TabIndex        =   5
               Top             =   1320
               Width           =   2895
            End
            Begin VB.Label lblAmount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   1680
               TabIndex        =   6
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label lblDate 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   120
               TabIndex        =   7
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblAgentName 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   4215
            End
            Begin VB.Label Label1 
               Caption         =   "Agent Name"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label2 
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Paid Amount"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   12
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   "Cash Received User Name"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3360
               TabIndex        =   11
               Top             =   1080
               Width           =   3135
            End
            Begin VB.Label Label5 
               Caption         =   "Agent Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4560
               TabIndex        =   10
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblAgentCode 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   375
               Left            =   4560
               TabIndex        =   8
               Top             =   480
               Width           =   1695
            End
         End
         Begin btButtonEx.ButtonEx bttnSerch 
            Height          =   375
            Left            =   3240
            TabIndex        =   3
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Serch"
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
         Begin btButtonEx.ButtonEx bttnCancel 
            Height          =   495
            Left            =   360
            TabIndex        =   15
            Top             =   4200
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            Appearance      =   3
            Enabled         =   0   'False
            Caption         =   "Cancel"
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
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58851331
         CurrentDate     =   40669
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   7200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
End
Attribute VB_Name = "frmCancellationAgentPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemAgentCashsettleId As Long
Dim TemAgentId As Long

Private Sub btnSearchByAmount_Click()
    With gridBills
        .Clear
        .Rows = 1
        .Cols = 3
        .Row = 0
        .col = 0
        .Text = "Bill No"
        .col = 1
        .Text = "Agent"
        .col = 2
        .Text = "Date"
        
    End With
    Dim temSQL As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
    
        If .State = 1 Then .Close
        temSQL = "SELECT tblAgentCashSettle.*, tblInstitutions.InstitutionName,tblInstitutions.InstitutionCode, tblStaff.StaffListedName FROM tblStaff RIGHT JOIN (tblAgentCashSettle LEFT JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID) ON tblStaff.Staff_ID = tblAgentCashSettle.User_ID Where Cash = " & Val(txtSearchAmount.Text) & " AND SettledDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        
        
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            gridBills.Rows = gridBills.Rows + 1
            gridBills.Row = gridBills.Rows - 1
            gridBills.col = 0
            gridBills.Text = !ReceiptNo
            gridBills.col = 1
            gridBills.Text = !InstitutionName
            gridBills.col = 2
            gridBills.Text = Format(!SettledDate, "dd MMMM yyyy")
            
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub bttnCancel_Click()
Dim A
A = MsgBox("Are you Sure, Cancel This Payment?", vbInformation + vbYesNo, "Verify")
If A = vbNo Then: Exit Sub

With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    .Open "SELECT* From tblAgentCashSettle Where AgentCashSettle_ID = " & TemAgentCashsettleId & ""
    If .RecordCount = 0 Then Exit Sub
    If !Cash = 0 Then A = MsgBox("This Recipt no already cancel Or Error please check", vbInformation + vbOKOnly, "No Payment"): Exit Sub
    
    !Cash = 0
    .Update
    
    If .State = 1 Then .Close
    
    .Open "Select* From tblInstitutions Where Institution_ID = " & TemAgentId & ""
    If .RecordCount = 0 Then Exit Sub
    
    !InstitutionCredit = Val(!InstitutionCredit) - Val(lblAmount.Caption)
    .Update
    
    If .State = 1 Then .Close
    .Open "Select * From tblAgentPaymentCancellation"
    .AddNew
    !AgentID = TemAgentId
    !UserID = UserID
    !Date = Date
    !RefNo = Val(txtReceiptNo.Text)
    !Amount = Val(lblAmount.Caption)
    .Update
    .Close
End With
Call ClearValues
bttnCancel.Enabled = False
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnSerch_Click()
Call FindReceiptNo
End Sub

Private Sub FindReceiptNo()

With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    .Open "SELECT tblAgentCashSettle.*, tblInstitutions.InstitutionName,tblInstitutions.InstitutionCode, tblStaff.StaffListedName FROM tblStaff RIGHT JOIN (tblAgentCashSettle LEFT JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID) ON tblStaff.Staff_ID = tblAgentCashSettle.User_ID Where ReceiptNo = '" & txtReceiptNo.Text & "'"
    If .RecordCount = 0 Then Call ClearValues: Exit Sub
    
    lblAgentName.Caption = !InstitutionName
    lblAgentCode.Caption = !InstitutionCode
    lblDate.Caption = !SettledDate
    lblAmount.Caption = Format(!Cash, "0.00")
    lblUserName.Caption = !staffListedName
    TemAgentCashsettleId = !AgentCashSettle_ID
    TemAgentId = !Institution_Id
    bttnCancel.Enabled = True

    If .State = 1 Then .Close
End With

End Sub

Private Sub ClearValues()
    lblAgentName.Caption = Empty
    lblAgentCode.Caption = Empty
    lblDate.Caption = Empty
    lblAmount.Caption = Empty
    lblUserName.Caption = Empty
    TemAgentCashsettleId = Empty
    TemAgentId = Empty

End Sub

Private Sub Form_Load()
    GetCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call FindReceiptNo
End Sub
