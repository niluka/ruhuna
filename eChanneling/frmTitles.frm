VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTitles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titles"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11280
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4335
   End
   Begin VB.Frame FrameIx 
      Height          =   3375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Sa&ve"
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
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
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
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "A&dd"
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
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9975
      _Version        =   393216
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Ed&it"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
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
Attribute VB_Name = "frmTitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FromGrid As Boolean
    Dim TemTitleID As Long
    Dim BorderMargin As Long

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
    txtName.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
Dim TemResponce As Byte

If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You must enter a Title", vbCritical, "No Title")
    txtName.SetFocus
    Exit Sub
End If

Call EditData
Call ClearValues
Call BeforeAddEdit

End Sub

Private Sub EditData()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblTitle where Title_ID = " & TemTitleID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    !Title = Trim(txtName.Text)
    .Update
    MsgBox Grid1.Row
    Grid1.Col = 1
    Grid1.Text = Trim(txtName.Text)
    Grid1.Col = 2
    Grid1.Text = !Title_ID
    TemTitleID = !Title_ID
End With
End Sub

Private Sub FormatGrid()

Dim BorderMargin As Long
BorderMargin = 100


With Grid1
    .Clear
    
    .Rows = 1
    .Cols = 3
    
    .ColWidth(0) = 600
    .ColWidth(2) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
    
    .Col = 0
    .CellAlignment = 4
    .Text = "No."
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Title"
    
End With
End Sub

Private Sub FillGrid()
Dim NowRow As Long
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tbltitle order by title"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    .MoveFirst
    NowRow = 0
    While .EOF = False
        NowRow = NowRow + 1
        Grid1.Rows = NowRow + 1
        Grid1.Row = NowRow
        Grid1.Col = 0
        Grid1.Text = NowRow
        Grid1.Col = 1
        Grid1.Text = !Title
        Grid1.Col = 2
        Grid1.Text = !Title_ID
        .MoveNext
    Wend

End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * from tbltitle where Title_ID = " & TemTitleID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close:   Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    Dim TemNum As Long
    With Grid1
        .RemoveItem (Grid1.Row)
        .Col = 0
        
        .Sort = 1
        
        For TemNum = 1 To .Rows - 1
            .Row = TemNum
            .Text = TemNum
        Next
    End With
    
    Call ClearValues
    Call BeforeAddEdit
    
    
End Sub

Private Sub bttnEdit_Click()
    FromGrid = True
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
Dim TemResponce As Byte

If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You must enter Title", vbCritical, "No Title")
    txtName.SetFocus
    Exit Sub
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Select * from tbltitle where ( title = '" & Trim(txtName.Text) & "')"
    If .State = 0 Then .Open
    If .RecordCount <> 0 Then
        TemResponce = MsgBox("The Title you entered already exist.", vbCritical, "Name Exists")
        txtName.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
End With

Call SaveDetails
Call ClearValues
Call BeforeAddEdit

End Sub


Private Sub SaveDetails()
Dim TemNum As Long

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tbltitle "
    If .State = 0 Then .Open
    .AddNew
    !Title = Trim(txtName.Text)
    .Update
    TemTitleID = !Title_ID
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Row = Grid1.Rows - 1
    Grid1.Col = 1
    Grid1.Text = Trim(txtName.Text)
    Grid1.Col = 2
    Grid1.Text = TemTitleID
    Grid1.Col = 1
    Grid1.Sort = 1
    Grid1.Col = 0
    For TemNum = 1 To Grid1.Rows - 1
        Grid1.Row = TemNum
        Grid1.Text = TemNum
    Next

End With
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call BeforeAddEdit
    
End Sub

Private Sub Grid1_Click()
    FromGrid = True
            
            bttnAdd.Enabled = True
            
            With Grid1
                If .Row < 1 Then FromGrid = False: Exit Sub
                .Col = 2
                If Not IsNumeric(.Text) Then FromGrid = False: Exit Sub
                TemTitleID = Val(.Text)
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

Private Sub ClearValues()
    txtName.Text = Empty
End Sub

Private Sub GetData()
With DataEnvironment1.rssqlTem2
    If .State = 1 Then .Close
    .Source = "SELECT * from tbltitle where Title_ID = " & TemTitleID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    If Not IsNull(!Title) Then txtName.Text = !Title
    If .State = 1 Then .Close
End With
End Sub


Private Sub BeforeAddEdit()

txtSearch.Text = Empty

Grid1.Enabled = True
FrameIx.Enabled = False
bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = False
bttnCancel.Visible = False

FromGrid = False
Call ClearValues

End Sub


Private Sub AfterAdd()

txtSearch.Text = Empty

Call ClearValues

Grid1.Enabled = False
FrameIx.Enabled = True
bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = True
bttnChange.Visible = False
bttnCancel.Visible = True

End Sub

Private Sub AfterEdit()

txtSearch.Text = Empty

Grid1.Enabled = True
FrameIx.Enabled = True
bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = True
bttnCancel.Visible = True

End Sub


Private Sub txtSearch_Change()

    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = Grid1.Rows
    Grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        Grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(Grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((Grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        Grid1.TopRow = TemArray
        Grid1.Row = TemArray
        Grid1.Col = 0
        Grid1.ColSel = (Grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        bttnAdd.Enabled = False
        Grid1.Col = 2
        TemTitleID = Grid1.Text
        Call GetData
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
    Else
        Grid1.TopRow = 1
        Grid1.Row = 0
        Grid1.Col = 0
        Grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
    End If
'**************************************



End Sub
