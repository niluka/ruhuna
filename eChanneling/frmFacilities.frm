VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmFacilities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facilities"
   ClientHeight    =   6105
   ClientLeft      =   2385
   ClientTop       =   3390
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacilities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11130
   Begin VB.Frame framFacility 
      Caption         =   "Facility Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton OptionDoctor 
         Caption         =   "&By a Doctor"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.OptionButton OptionStaff 
         Caption         =   "By An&other Staff Member"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   1560
         Width           =   3975
      End
      Begin VB.OptionButton OptionOther 
         Caption         =   "O&ther"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   2040
         Width           =   2655
      End
      Begin VB.OptionButton OptionInvestigation 
         Caption         =   "&Investigation"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtComments 
         Height          =   1695
         Left            =   1920
         TabIndex        =   9
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
         BackStyle       =   1
         Caption         =   "&Save"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   495
         Left            =   2040
         TabIndex        =   11
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "C&hange"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Ca&ncel"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Catogery"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Facility &Comments"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Facility Name"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      Top             =   5520
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4455
      Left            =   480
      TabIndex        =   14
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7858
      _Version        =   393216
      ScrollTrack     =   -1  'True
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Delete"
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
Attribute VB_Name = "frmFacilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemFacilityID As Long
    Dim FromGrid As Boolean
    Dim CatogeryID  As Integer


Private Sub bttnAdd_Click()
Call ClearValues
Call AfterAdd

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




bttnAdd.BackColor = BttnBackColour
bttnAdd.ForeColor = BttnForeColour

bttnCancel.BackColor = BttnBackColour
bttnCancel.ForeColor = BttnForeColour

bttnChange.BackColor = BttnBackColour
bttnChange.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnEdit.BackColor = BttnBackColour
bttnEdit.ForeColor = BttnForeColour

bttnSave.BackColor = BttnBackColour
bttnSave.ForeColor = BttnForeColour

frmFacilities.BackColor = FrmBackColour
frmFacilities.ForeColor = FrmForeColour

framFacility.BackColor = FrameBackColour
framFacility.ForeColor = FrameForeColour

Label1.BackColor = FrameBackColour
Label1.ForeColor = FrameForeColour

Label2.BackColor = FrameBackColour
Label2.ForeColor = FrameForeColour

'FramePrivate.BackColor = FrameBackColour
'FramePrivate.ForeColor = FrameForeColour

'FrameOfficial.BackColor = FrameBackColour
'FrameOfficial.ForeColor = FrameForeColour

'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'chkCurrentlyChanneling.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour

'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour

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
OptionDoctor.BackColor = FrameBackColour
OptionDoctor.ForeColor = FrameForeColour

OptionInvestigation.BackColor = FrameBackColour
OptionInvestigation.ForeColor = FrameForeColour

OptionOther.BackColor = FrameBackColour
OptionOther.ForeColor = FrameForeColour

OptionStaff.BackColor = FrameBackColour
OptionStaff.ForeColor = FrameForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




grid1.BackColor = GridBackColor
grid1.ForeColor = GridForeColor

grid1.BackColorBkg = GridBackColorBkg
grid1.BackColorFixed = GridBackColorFixed
grid1.BackColorSel = GridBackColorSel

grid1.ForeColor = GridForeColor
grid1.ForeColorFixed = GridForeColorFixed
grid1.ForeColorSel = GridForeColorSel



End Sub







Private Sub bttnCancel_Click()
 Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnChange_Click()
Dim TemResponce  As Integer
If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You have not entered a name of a facility to add", vbCritical, "Facility?")
    txtName.SetFocus
    Exit Sub
End If

Call EditData

End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnDelete_Click()
Dim TemResponce  As Integer
grid1.Col = 2
If Not IsNumeric(grid1.Text) Then Exit Sub
grid1.Col = 1
TemResponce = MsgBox("Are you sure you want to remove " & grid1.Text & " from the Facilities list that the hospital provide", vbCritical + vbYesNo, "?Remove")
If TemResponce = vbNo Then Exit Sub

grid1.Col = 2
With DataEnvironment1.rssqlHospitalFacility
    If .State = 1 Then .Close
    .Source = "Select tblhospitalfacility.* from tblhospitalfacility where hospitalfacility_ID = " & grid1.Text
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    .Delete adAffectCurrent
    .Close
    
End With
Call FormatGrid
Call FillGrid
Call BeforeAddEdit
End Sub

Private Sub bttnEdit_Click()
grid1.Col = 2
TemFacilityID = grid1.Text
Call AfterEdit
End Sub

Private Sub bttnSave_Click()
Dim TemResponce  As Integer
If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You have not entered a name of a facility to add", vbCritical, "Facility?")
    txtName.SetFocus
    Exit Sub
End If
Call SaveData
End Sub

Private Sub Form_Load()
Call ClearValues
Call BeforeAddEdit
Call FormatGrid
Call FillGrid
Call Setcolours


End Sub


Private Sub BeforeAddEdit()

Call ClearValues

bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

grid1.Enabled = True

framFacility.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = False
bttnCancel.Visible = False

OptionDoctor.Value = True

txtSearch.Text = Empty
On Error Resume Next
txtSearch.SetFocus

End Sub

Private Sub AfterAdd()

Call ClearValues

txtName.Text = txtSearch.Text

bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnDelete.Enabled = False

grid1.Enabled = False

framFacility.Enabled = True

bttnSave.Visible = True
bttnChange.Visible = False
bttnCancel.Visible = True

txtName.SetFocus

End Sub

Private Sub AfterEdit()

bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnDelete.Enabled = False

grid1.Enabled = False

framFacility.Enabled = True

bttnSave.Visible = False
bttnChange.Visible = True
bttnCancel.Visible = True

txtName.SetFocus
SendKeys "{Home}+{end}"

End Sub

Private Sub GetData()

With DataEnvironment1.rssqlHospitalFacility
    If .State = 1 Then .Close
    .Source = "SELECT tblHospitalFacility.* from tblHospitalFacility where (HospitalFacility_ID = " & TemFacilityID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    
    Call ClearValues
    
    txtName.Text = !hospitalfacility
    If Not IsNull(!HospitalFacilityComment) Then txtComments.Text = !HospitalFacilityComment
    
    Select Case !PersonCatogery
        Case Doctor:        OptionDoctor.Value = True
        
        Case Staff:         OptionStaff.Value = True
        
        Case Investigation: OptionInvestigation.Value = True
        
        Case Other:         OptionOther.Value = True
    End Select
    
    .Close

End With

End Sub

Private Sub SaveData()

With DataEnvironment1.rssqlHospitalFacility
    If .State = 1 Then .Close
    .Source = "SELECT tblHospitalFacility.* from tblHospitalFacility"
    If .State = 0 Then .Open
    
    .AddNew
    
    !hospitalfacility = txtName.Text
    !HospitalFacilityComment = txtComments.Text
    
    If OptionDoctor.Value = True Then
        !PersonCatogery = Doctor
    ElseIf OptionStaff.Value = True Then
        !PersonCatogery = Staff
    ElseIf OptionInvestigation.Value = True Then
        !PersonCatogery = Investigation
    ElseIf OptionOther.Value = True Then
        !PersonCatogery = Other
    End If
    
    .Update
    .Close

End With

Call ClearValues
Call BeforeAddEdit
Call FormatGrid
Call FillGrid


End Sub

Private Sub EditData()

With DataEnvironment1.rssqlHospitalFacility
    If .State = 1 Then .Close
    .Source = "SELECT tblHospitalFacility.* from tblHospitalFacility where (HospitalFacility_ID = " & TemFacilityID & ")"
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    !hospitalfacility = txtName.Text
    !HospitalFacilityComment = txtComments.Text
    
    If OptionDoctor.Value = True Then
        !PersonCatogery = Doctor
    ElseIf OptionStaff.Value = True Then
        !PersonCatogery = Staff
    ElseIf OptionInvestigation.Value = True Then
        !PersonCatogery = Investigation
    ElseIf OptionOther.Value = True Then
        !PersonCatogery = Other
    End If
    
    .Update
    .Close

End With

Call ClearValues
Call BeforeAddEdit
Call FormatGrid
Call FillGrid

End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtComments.Text = Empty
End Sub


Private Sub FormatGrid()
Dim BorderMargin As Integer
BorderMargin = 100

With grid1
    .Clear
    .Cols = 3
    .Rows = 1
    
    .Row = 0
    
    .ColWidth(0) = 600
    .ColWidth(2) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
    
    .Col = 0
    .CellAlignment = 4
    .Text = "No."
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Facility"

End With

End Sub

Private Sub FillGrid()
Dim NowRow As Long

With DataEnvironment1.rssqlHospitalFacility
    If .State = 1 Then .Close
    .Source = "SELECT tblHospitalFacility.* from tblHospitalFacility order by HospitalFacility"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    
    NowRow = 0
    
    While .EOF = False
        
        NowRow = NowRow + 1
        
        grid1.Rows = NowRow + 1
        grid1.Row = NowRow
        
        grid1.Col = 0
        grid1.CellAlignment = 1
        grid1.Text = NowRow
        
        grid1.Col = 1
        grid1.CellAlignment = 1
        grid1.Text = !hospitalfacility
        
        grid1.Col = 2
        grid1.Text = !HospitalFacility_id
        
        .MoveNext
                
    Wend
    
End With
End Sub


Private Sub Grid1_Click()
FromGrid = True
With grid1
    If .Row < 1 Then Exit Sub
    .Col = 2
    TemFacilityID = Val(.Text)
    If Not IsNumeric(.Text) Then Exit Sub
    .Col = 1
    txtSearch.Text = .Text
    Call GetData
    
    .Col = 0
    .ColSel = .Cols - 1
    
    txtSearch.SetFocus
    SendKeys "{home}+{end}"

FromGrid = False

bttnAdd.Enabled = False
bttnEdit.Enabled = True
bttnDelete.Enabled = True

End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then Grid1_Click
End Sub

Private Sub txtSearch_Change()
If FromGrid = True Then Exit Sub
Dim TemFRows As Long
Dim TemNowRow As Long
Dim TemArray As Long
Dim SearchSuccess As Boolean
Dim TemLength As Single

TemFRows = grid1.Rows
grid1.Col = 1

If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess

SearchSuccess = False


For TemArray = 1 To (TemFRows - 1)
    grid1.Row = TemArray
    If Len(txtSearch.Text) > Len(grid1.Text) Then
        GoTo FinishLoop
    Else
        TemLength = Len(txtSearch.Text)
    End If
    
    If UCase(Left((grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
        SearchSuccess = True
        Exit For
    Else
        SearchSuccess = False
    End If
FinishLoop:
Next

MeasureSuccess:

If SearchSuccess = True Then
    grid1.TopRow = TemArray
    grid1.Row = TemArray
    grid1.Col = 0
    grid1.ColSel = (grid1.Cols - 1)
    bttnEdit.Enabled = True
    bttnDelete.Enabled = True
    bttnAdd.Enabled = False
    'grid1_Click
    grid1.Col = 2
    TemFacilityID = grid1.Text
    Call GetData
    grid1.Col = 0
    grid1.ColSel = grid1.Cols - 1
    
    
Else
    grid1.Row = 0
    grid1.Col = 0
    grid1.ColSel = 0
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
End If

End Sub
