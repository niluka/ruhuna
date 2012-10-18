VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAgentRefranceNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Refrance No"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
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
   ScaleHeight     =   5310
   ScaleWidth      =   7635
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Issue New Book"
      TabPicture(0)   =   "frmAgentRefranceNo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "bttnUpdate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPages"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtBookno"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtEndRefNo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtBeginRefno"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Book Issue Details"
      TabPicture(1)   =   "frmAgentRefranceNo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTab2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   -74280
         TabIndex        =   15
         Top             =   1440
         Width           =   5535
         Begin btButtonEx.ButtonEx ButtonEx1 
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "View All Details"
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
      Begin VB.TextBox txtBeginRefno 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtEndRefNo 
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtBookno 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtPages 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   3000
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Issue"
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   2775
         Left            =   -74640
         TabIndex        =   16
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Today"
         TabPicture(0)   =   "frmAgentRefranceNo.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblDate"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Selected Day"
         TabPicture(1)   =   "frmAgentRefranceNo.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DTPicker1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Period"
         TabPicture(2)   =   "frmAgentRefranceNo.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label7"
         Tab(2).Control(1)=   "Label8"
         Tab(2).Control(2)=   "DTPicker2"
         Tab(2).Control(3)=   "DTPicker3"
         Tab(2).ControlCount=   4
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   -71400
            TabIndex        =   22
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   62259201
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   -73920
            TabIndex        =   21
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   62259201
            CurrentDate     =   39536
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   -72840
            TabIndex        =   18
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   62259201
            CurrentDate     =   39536
         End
         Begin VB.Label Label8 
            Caption         =   "To"
            Height          =   255
            Left            =   -71760
            TabIndex        =   20
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "From"
            Height          =   255
            Left            =   -74520
            TabIndex        =   19
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblDate 
            Height          =   375
            Left            =   2160
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Begin Refrance No"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "End Refrance No"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Boook Number"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Pages"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
   End
   Begin MSDataListLib.DataCombo dtcAgentName 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataListLib.DataCombo dtcAgentCode 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Agent Code"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Agent Name"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAgentRefranceNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CheckValue As Boolean
Dim TemBookID As Long

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnUpdate_Click()
If CanIssue = False Then Exit Sub
Call UpdateBookNumber
End Sub

Private Sub ButtonEx1_Click()
    With DataEnvironment1.rscmmdTemText
    If .State = 1 Then .Close
    
    Select Case SSTab2.Tab
    Case 0
        .Open "SELECT tblAgentRefBook.*, tblInstitutions.InstitutionName, tblStaff.StaffName FROM tblStaff RIGHT JOIN (tblInstitutions RIGHT JOIN tblAgentRefBook ON tblInstitutions.Institution_ID = tblAgentRefBook.Agent_ID) ON tblStaff.Staff_ID = tblAgentRefBook.IssuedStaff_ID Where (IssuedDate = '" & Date & "') Order By Start"
    Case 1
        .Open "SELECT tblAgentRefBook.*, tblInstitutions.InstitutionName, tblStaff.StaffName FROM tblStaff RIGHT JOIN (tblInstitutions RIGHT JOIN tblAgentRefBook ON tblInstitutions.Institution_ID = tblAgentRefBook.Agent_ID) ON tblStaff.Staff_ID = tblAgentRefBook.IssuedStaff_ID Where (IssuedDate = '" & DTPicker1.Value & "') Order By Start"
    Case 2
        .Open "SELECT tblAgentRefBook.*, tblInstitutions.InstitutionName, tblStaff.StaffName FROM tblStaff RIGHT JOIN (tblInstitutions RIGHT JOIN tblAgentRefBook ON tblInstitutions.Institution_ID = tblAgentRefBook.Agent_ID) ON tblStaff.Staff_ID = tblAgentRefBook.IssuedStaff_ID Where (IssuedDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') Order By Start"
    
    End Select
    
        With dtrBookRefIssueNo
         Set .DataSource = DataEnvironment1.rscmmdTemText
        
            Select Case SSTab2.Tab
            Case 0
            .Sections("Section4").Controls("lblDate").Caption = "Date   :  " & Date
            Case 1
            .Sections("Section4").Controls("lblDate").Caption = "Date   :  " & DTPicker1.Value
            Case 2
            .Sections("Section4").Controls("lblDate").Caption = "Date From   :  " & DTPicker2.Value & vbTab & " To    :" & DTPicker3.Value
            End Select
        
         .Show
        End With
    End With
End Sub

Private Sub dtcAgentCode_Click(Area As Integer)
If Not IsNumeric(dtcAgentCode.BoundText) Then Exit Sub
dtcAgentName.BoundText = dtcAgentCode.BoundText
End Sub

Private Sub dtcAgentName_Click(Area As Integer)
If Not IsNumeric(dtcAgentName.BoundText) Then Exit Sub
dtcAgentCode.BoundText = dtcAgentName.BoundText
End Sub


Private Sub Form_Load()
lblDate.Caption = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date

Call FillAgentconbos
End Sub

Private Sub FillAgentconbos()

    With DataEnvironment1.rscmmdViewAgent
        If .State = 1 Then .Close
        .Open "Select* From tblInstitutions Order By InstitutionName"
        Set dtcAgentName.RowSource = DataEnvironment1.rscmmdViewAgent
        dtcAgentName.BoundColumn = "Institution_ID"
        dtcAgentName.ListField = "InstitutionName"
        Set dtcAgentCode.RowSource = DataEnvironment1.rscmmdViewAgent
        dtcAgentCode.BoundColumn = "Institution_ID"
        dtcAgentCode.ListField = "InstitutionCode"
    End With
    
End Sub

Private Function CanIssue() As Boolean
    Dim A As Integer
    CanIssue = False
    If Not IsNumeric(dtcAgentName.BoundText) Then
        A = MsgBox("Select Agent Name", vbCritical + vbOKOnly, "Empty Name"): dtcAgentName.SetFocus
        dtcAgentName.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcAgentCode.BoundText) Then
        A = MsgBox("Select Agent Name", vbCritical + vbOKOnly, "Empty Name"): dtcAgentCode.SetFocus
        dtcAgentCode.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtBeginRefno.Text) Then
        A = MsgBox("Enter valid Beging Ref No", vbCritical + vbOKOnly, "Empty Name"): txtBeginRefno.SetFocus
        txtBeginRefno.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtEndRefNo.Text) Then
        A = MsgBox("Enter valied End Ref No", vbCritical + vbOKOnly, "Empty Name"): txtEndRefNo.SetFocus
        txtEndRefNo.SetFocus
        Exit Function
    End If
    If Val(txtBeginRefno.Text) > Val(txtEndRefNo.Text) Then
        A = MsgBox("Enter Correct Begind and  End Ref No", vbCritical + vbOKOnly, "Empty Name"): txtEndRefNo.SetFocus
        txtEndRefNo.SetFocus
        Exit Function
    End If
    If txtBookno.Text = Empty Then
        A = MsgBox("Enter Book No", vbCritical + vbOKOnly, "Empty Name"): txtBookno.SetFocus
        txtBookno.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtPages.Text) Then
        A = MsgBox("Enter Total Pages", vbCritical + vbOKOnly, "Empty Name"): txtPages.SetFocus
        txtPages.SetFocus
        Exit Function
    End If
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
'        .Source = "SELECT tblAgentRef.AgentRefNo FROM tblAgentRef WHERE (((tblAgentRef.AgentRefNo)>=10)) OR (((tblAgentRef.AgentRefNo)<=100)) ORDER BY tblAgentRef.AgentRefNo"
       .Source = "SELECT tblAgentRef.AgentRefNo FROM tblAgentRef WHERE (tblAgentRef.AgentRefNo Between " & Val(txtBeginRefno.Text) & " and " & Val(txtEndRefNo.Text) & " ) ORDER BY tblAgentRef.AgentRefNo"

        .Open
        If .RecordCount <> 0 Then
            A = MsgBox("Some numbers in this Book are already issued. Please check", vbInformation, "Already Issued")
            txtBeginRefno.SetFocus
            SendKeys "{home}+{end}"
            If .State = 1 Then .Close
            Exit Function
        End If
        If .State = 1 Then .Close
    End With
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Open "Select* From tblAgentRefBook Where BookName = '" & txtBookno.Text & "'"
        If .RecordCount <> 0 Then
            A = MsgBox("This Book No " & txtBookno.Text & " is already issued", vbCritical + vbOKOnly, "Error")
            txtBookno.SetFocus
            If .State = 1 Then .Close
            Exit Function
        If .State = 1 Then .Close
        End If
    End With
    CanIssue = True
End Function

Private Sub ClearValues()
dtcAgentName.Text = Empty
dtcAgentCode.Text = Empty
txtBeginRefno.Text = Empty
txtEndRefNo.Text = Empty
txtBookno.Text = Empty
txtPages.Text = Empty
End Sub

Private Sub UpdateBookNumber()
Dim i As Long
Dim Number As Long
Dim A
Number = Val(txtBeginRefno.Text)

With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    .Open "Select* From tblAgentRefBook Where BookName = '" & txtBookno.Text & "'"
'
'    If .RecordCount = 1 Then A = MsgBox("This Book No  & vbTab & & txtBookno.text &  Already Ented", vbCritical + vbOKOnly, "Error"): Exit Sub
'
    .AddNew
    !BookName = txtBookno.Text
    !Start = Val(txtBeginRefno.Text)
    !End = Val(txtEndRefNo.Text)
    !Leaves = Val(txtPages.Text)
    !IssuedStaff_ID = UserID
    !IssuedDate = Date
    !IssuedTime = Time
    !Agent_ID = Val(dtcAgentName.BoundText)
    .Update
     TemBookID = !AgentRefBook_ID
  
    If .State = 1 Then .Close
    .Open "Select* From tblAgentRef"
    
    
    For i = Number To Val(txtEndRefNo.Text)
        .AddNew
        !AgentRefNo = i
        !Agent_ID = Val(dtcAgentName.BoundText)
        !AgenRefBook_ID = TemBookID
        !IssuedDate = Date
        .Update
    Next i
    
    If .State = 1 Then .Close
    Call ClearValues
    MsgBox "Issued"
End With

End Sub

Private Sub CalculatePages()
If Not IsNumeric(txtPages.Text) Then Exit Sub
If IsNumeric(txtEndRefNo.Text) = False And IsNumeric(txtBeginRefno.Text) = False Then Exit Sub
txtEndRefNo.Text = Val(txtPages.Text) + Val(txtBeginRefno.Text) - 1
End Sub

Private Sub txtBeginRefno_Change()
    Call CalculatePages
End Sub

Private Sub txtPages_Change()
    Call CalculatePages
End Sub
