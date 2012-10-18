VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPatientMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Main Details"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
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
   ScaleHeight     =   9000
   ScaleWidth      =   11190
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePatientMainDetails 
      Caption         =   "Patient Details"
      Height          =   8055
      Left            =   4560
      TabIndex        =   34
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtNotes 
         Height          =   825
         Left            =   2040
         TabIndex        =   23
         Top             =   7080
         Width           =   4215
      End
      Begin VB.TextBox txtEmail 
         Height          =   345
         Left            =   2040
         TabIndex        =   22
         Top             =   6600
         Width           =   4215
      End
      Begin VB.TextBox txtFax 
         Height          =   345
         Left            =   2040
         TabIndex        =   21
         Top             =   6120
         Width           =   4215
      End
      Begin VB.TextBox txtTelephone 
         Height          =   345
         Left            =   2040
         TabIndex        =   20
         Top             =   5640
         Width           =   4215
      End
      Begin VB.TextBox txtAddress 
         Height          =   825
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4680
         Width           =   4215
      End
      Begin VB.TextBox txtNIC 
         Height          =   345
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtSurname 
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtOtherName 
         Height          =   345
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtFirstName 
         Height          =   345
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo DataComboTitle 
         Bindings        =   "frmPatientMain.frx":0000
         Height          =   360
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Title"
         BoundColumn     =   "Title_ID"
         Text            =   ""
         Object.DataMember      =   "sqlTitle"
      End
      Begin MSDataListLib.DataCombo DataComboSex 
         Bindings        =   "frmPatientMain.frx":001F
         Height          =   360
         Left            =   2040
         TabIndex        =   12
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Sex"
         BoundColumn     =   "Sex_ID"
         Text            =   ""
         Object.DataMember      =   "sqlSex"
      End
      Begin MSDataListLib.DataCombo DataComboMarietal 
         Bindings        =   "frmPatientMain.frx":003E
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Marietal"
         BoundColumn     =   "Marietal_ID"
         Text            =   ""
         Object.DataMember      =   "sqlMarietal"
      End
      Begin MSDataListLib.DataCombo DataComboRace 
         Bindings        =   "frmPatientMain.frx":005D
         Height          =   360
         Left            =   2040
         TabIndex        =   13
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Race"
         BoundColumn     =   "Race_ID"
         Text            =   ""
         Object.DataMember      =   "sqlRace"
      End
      Begin MSComCtl2.DTPicker DTPickerDOB 
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   4200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39413
      End
      Begin btButtonEx.ButtonEx bttnLoad 
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   3720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "&Load"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnRemove 
         Height          =   255
         Left            =   5280
         TabIndex        =   16
         Top             =   3720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "Re&move"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         Caption         =   "(Age)"
         Height          =   255
         Left            =   3600
         TabIndex        =   49
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date &of Birth"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Image ImagePatient 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   3840
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "No&tes"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   7200
         Width           =   3615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "E-&Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   6120
         Width           =   3615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "&Telephone"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "NIC &No."
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Race"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marietal"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Se&x"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Title"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Surname"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Othe&r Names"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "First &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   3615
      End
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8160
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
      Height          =   4815
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8493
      _Version        =   393216
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin btButtonEx.ButtonEx bttnSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Se&arch"
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
   Begin VB.Frame FrameSearchNames 
      Caption         =   "Search By Names"
      Height          =   1575
      Left            =   120
      TabIndex        =   30
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox txtSearchSurname 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtSearchFirstName 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameSearchID 
      Caption         =   "Search By ID"
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtSearchID 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   8160
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
      Left            =   9480
      TabIndex        =   27
      Top             =   8520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   8160
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   8160
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
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   255
      Left            =   7800
      TabIndex        =   26
      Top             =   8160
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
   Begin VB.TextBox txtPhoto 
      Height          =   360
      Left            =   240
      TabIndex        =   50
      Top             =   7680
      Visible         =   0   'False
      Width           =   9255
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   8160
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
Attribute VB_Name = "frmPatientMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemPatientID As Long
Dim TemRow As Long
Dim TemRowSel As Long
Dim TemArrey1 As Long
Dim FromAge As Boolean

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce  As Integer
    If Trim(txtFirstName.Text) = "" Then
        TemResponce = MsgBox("Please enter the Firstname", vbCritical, "First Name")
        On Error Resume Next
        txtFirstName.SetFocus
        Exit Sub
    End If
    If Trim(txtSurname.Text) = "" Then
        TemResponce = MsgBox("Please enter the Surname", vbCritical, "First Name")
        txtSurname.SetFocus
        Exit Sub
    End If
    If DTPickerDOB.Value = Date Then
        TemResponce = MsgBox("Please enter the Birthday or Age", vbCritical, "First Name")
        DTPickerDOB.SetFocus
        Exit Sub
    End If
    Call EditDetails
    Call BeforeAddEdit

End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
Dim TemResponce  As Integer
If Grid1.Row < 1 Or Grid1.RowSel < 1 Then Exit Sub

If Grid1.Row < Grid1.RowSel Then
    TemRow = Grid1.Row
    TemRowSel = Grid1.RowSel
Else
    TemRow = Grid1.RowSel
    TemRowSel = Grid1.Row
End If




    For TemArrey1 = TemRow To TemRowSel
        Grid1.Row = TemArrey1
        TemResponce = MsgBox("Are you sure you want to delete the record of " & Grid1.TextMatrix(TemArrey1, 1) & " " & Grid1.TextMatrix(TemArrey1, 2) & "?", vbCritical + vbYesNo, "Delete?")
        If TemResponce = vbYes Then
            With DataEnvironment1.rssqlTem
            If .State = 1 Then .Close
            .Source = "select * from tblpatientmaindetails where patient_ID = " & Grid1.TextMatrix(TemArrey1, 0)
            .Open
            If .RecordCount = 0 Then Exit Sub
            .Delete adAffectCurrent
            .Close
            Grid1.RemoveItem TemArrey1
            End With
        End If
    Next

    bttnDelete.Enabled = False

End Sub

Private Sub bttnEdit_Click()

    AfterEdit
End Sub

Private Sub bttnLoad_Click()
    Dim TemResponce  As Integer
    Dim TemPhotoName As String
    ImagePatient.Stretch = True
    CommonDialog1.Filter = "BMP|*.BMP|JPG|*.JPG;JPE;JPEG|GIF|*.GIF|All Images|*.BMP;*.JPG;*.JPE;*.JPGE;*.GIF|All Files|*.*"
    CommonDialog1.ShowOpen
    On Error GoTo PhotoError:
    TemPhotoName = CommonDialog1.FileName
    ImagePatient.Picture = LoadPicture(TemPhotoName)
    txtPhoto.Text = App.Path & "\graphics\" & "Photo of " & txtFirstName.Text & " " & txtOtherName.Text & " " & txtSurname.Text & ".BMP"
    Exit Sub
PhotoError:
    If Err.Number = 481 Then
        TemResponce = MsgBox("The Photo you choose is not suitable, try using a medium size BMP, JPG or GIF file", vbOKOnly, "Photo Error")
    ElseIf Err.Number = 53 Then
        TemResponce = MsgBox("No photo exist to selected, try to select again correctly.", vbOKOnly, "Photo Error")
    Else
        TemResponce = MsgBox("An unknown error has occured, try again," & Chr(13) & Err.Description, vbOKOnly, "Photo Error")
    End If
End Sub


Private Sub bttnRemove_Click()
    ImagePatient.Picture = LoadPicture()
    txtPhoto.Text = Empty
    
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce  As Integer
    If Trim(txtFirstName.Text) = "" Then
        TemResponce = MsgBox("Please enter the Firstname", vbCritical, "First Name")
        txtFirstName.SetFocus
        Exit Sub
    End If
    If Trim(txtSurname.Text) = "" Then
        TemResponce = MsgBox("Please enter the Surname", vbCritical, "First Name")
        txtSurname.SetFocus
        Exit Sub
    End If
    If DTPickerDOB.Value = Date Then
        TemResponce = MsgBox("Please enter the Birthday or Age", vbCritical, "First Name")
        DTPickerDOB.SetFocus
        Exit Sub
    End If
    Call SaveDetails
    Call BeforeAddEdit
End Sub

Private Sub bttnSearch_Click()
    If Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) = "" Then
        ListAllPatients
    ElseIf Trim(txtSearchID.Text) <> "" Then
        SearchFromID
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) = "" Then
        ListFirstNames
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) <> "" Then
        ListSurname
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then
        ListBothNames
    End If
    ClearSearchValues
End Sub


Private Sub DTPickerDOB_Change()
If FromAge = True Then Exit Sub
    txtAge.Text = CalculateAgeInWords(DTPickerDOB.Value)
    FromAge = False
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call Setcolours
    Call BeforeAddEdit
    DTPickerDOB.Format = dtpCustom
    DTPickerDOB.CustomFormat = DefaultShortDate
End Sub

Private Sub BeforeAddEdit()
    TemPatientID = Empty
    Call ClearSearchValues
    Call ClearPatientValues
    Call FormatGrid
    frameSearchID.Enabled = True
    FrameSearchNames.Enabled = True
    Grid1.Enabled = True
    FramePatientMainDetails.Enabled = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False

End Sub


Private Sub AfterAdd()
    Call ClearSearchValues
    Call ClearPatientValues
    Call FormatGrid
    frameSearchID.Enabled = False
    FrameSearchNames.Enabled = False
    FramePatientMainDetails.Enabled = True
    Grid1.Enabled = False
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Visible = True
    bttnCancel.Visible = True
    bttnChange.Visible = False
End Sub

Private Sub AfterEdit()
    Call ClearSearchValues
    Call FormatGrid
    frameSearchID.Enabled = False
    FrameSearchNames.Enabled = False
    FramePatientMainDetails.Enabled = True
    Grid1.Enabled = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnSave.Visible = False
    bttnCancel.Visible = True
    bttnChange.Visible = True
End Sub

Private Sub ClearPatientValues()
    txtFirstName.Text = Empty
    txtOtherName.Text = Empty
    txtSurname.Text = Empty
    txtAddress.Text = Empty
    DataComboTitle.Text = Empty
    DataComboSex.Text = Empty
    DataComboMarietal.Text = Empty
    DataComboRace.Text = Empty
    txtPhoto.Text = Empty
    ImagePatient.Picture = LoadPicture()
End Sub

Private Sub ClearSearchValues()
    txtSearchFirstName.Text = Empty
    txtSearchID.Text = Empty
    txtSearchSurname.Text = Empty
End Sub

Private Sub FormatGrid()
    Dim BorderMargin As Long
    BorderMargin = 100
    With Grid1
        .Clear
        .Rows = 1
        .Cols = 3
        .ColWidth(0) = 600
        .ColWidth(1) = ((.Width) - (.ColWidth(0)) - BorderMargin) * 2 / 5
        .ColWidth(2) = ((.Width) - (.ColWidth(0)) - BorderMargin) * 3 / 5
        .Col = 0
        .CellAlignment = 4
        .Text = "ID"
        .Col = 1
        .CellAlignment = 4
        .Text = "Firstname"
        .Col = 2
        .CellAlignment = 4
        .Text = "Surname"
    End With
End Sub

Private Sub FillGrid()
    Dim NowROw As Long
    With DataEnvironment1.rssqlPatientMain
    If .State = 1 Then .Close
        .Open
        If .RecordCount = 0 Then
            bttnEdit.Enabled = False
            Exit Sub
        Else
            bttnEdit.Enabled = True
        End If
        .MoveFirst
        NowROw = 0
        While .EOF = False
            NowROw = NowROw + 1
            Grid1.Rows = NowROw + 1
            Grid1.Row = NowROw
            Grid1.Col = 0
            Grid1.CellAlignment = 7
            Grid1.Text = !patient_ID
            Grid1.Col = 1
            Grid1.CellAlignment = 7
            If Not IsNull(!firstname) Then Grid1.Text = !firstname
            Grid1.Col = 2
            Grid1.CellAlignment = 7
            If Not IsNull(!surname) Then Grid1.Text = !surname
            .MoveNext
        Wend
    End With
End Sub


Private Sub ListAllPatients()
    Dim NowROw As Long
    Call FormatGrid
    With DataEnvironment1.rssqlTem6
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails order by Patient_ID"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub SearchFromID()
    Dim NowROw As Long
    Dim TemResponce  As Integer
    Call FormatGrid
    If Not IsNumeric(txtSearchID.Text) Then Exit Sub
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where Patient_ID = " & txtSearchID.Text & " order by Patient_ID"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no record with the patient ID of " & txtSearchID.Text, vbCritical, "Wrong ID")
            txtSearchID.SetFocus
            ClearPatientValues
            SendKeys "{Home}+{end}"
        End If
        Call FillGrid
        .Close
    End With
End Sub

Private Sub ListFirstNames()
    Dim NowROw As Long
    Call FormatGrid
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where firstname like '" & txtSearchFirstName.Text & "%' order by FIrstname"
        If .State = 0 Then .Open
        Call FillGrid
        If .State = 1 Then .Close
    End With
End Sub

Private Sub ListSurname()
    Dim NowROw As Long
    Call FormatGrid
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where surname like '" & txtSearchSurname.Text & "%' order by surname"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub ListBothNames()
    Dim NowROw As Long
    Call FormatGrid
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where (surname like '" & txtSearchSurname.Text & "%') and ( firstname like '" & txtSearchFirstName.Text & "%') order by FIrstname, surname"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub bttnAdd_Click()
    Call ClearSearchValues
    Call ClearPatientValues
    Call AfterAdd
End Sub

Private Sub GetDetails()
    Call ClearPatientValues
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where (patient_ID =" & TemPatientID & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!firstname) Then txtFirstName.Text = !firstname
        If Not IsNull(!surname) Then txtSurname.Text = !surname
        If Not IsNull(!othernames) Then txtOtherName.Text = !othernames
        If Not IsNull(!Title_ID) Then DataComboTitle.BoundText = !Title_ID
        If Not IsNull(!sex_ID) Then DataComboSex.BoundText = !sex_ID
        If Not IsNull(!Marital_ID) Then DataComboMarietal.BoundText = !Marital_ID
        If Not IsNull(!Race_ID) Then DataComboRace.BoundText = !Race_ID
        If Not IsNull(!NICNo) Then txtNIC.Text = !NICNo
        If Not IsNull(!Address) Then txtAddress.Text = !Address
        If Not IsNull(!phone) Then txtTelephone.Text = !phone
        If Not IsNull(!fax) Then txtFax.Text = !fax
        If Not IsNull(!email) Then txtEmail.Text = !email
        If Not IsNull(!DateOfBirth) Then DTPickerDOB.Value = !DateOfBirth
        If Not IsNull(!notes) Then txtNotes.Text = !notes
        If Not IsNull(!photo) Then txtPhoto.Text = !photo
        Call DisplayPhoto
        .Close
    End With
End Sub

Private Sub SaveDetails()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select tblpatientmaindetails.* from tblPatientMainDetails order by Patient_ID"
        If .State = 0 Then .Open
        .AddNew
        !firstname = txtFirstName.Text
        !surname = txtSurname.Text
        !othernames = txtOtherName.Text
        If IsNumeric(DataComboTitle.BoundText) Then !Title_ID = DataComboTitle.BoundText
        If IsNumeric(DataComboSex.BoundText) Then !sex_ID = DataComboSex.BoundText
        If IsNumeric(DataComboMarietal.BoundText) Then !Marital_ID = DataComboMarietal.BoundText
        If IsNumeric(DataComboRace.BoundText) Then !Race_ID = DataComboRace.BoundText
        !NICNo = txtNIC.Text
        !Address = txtAddress.Text
        !phone = txtTelephone.Text
        !fax = txtFax.Text
        !email = txtEmail.Text
        !DateOfBirth = DTPickerDOB.Value
        !notes = txtNotes.Text
        SavePhoto
        !photo = txtPhoto.Text
        !registeredDate = Date
        .Update
        .Close
    End With
End Sub

Private Sub SavePhoto()
    On Error Resume Next
    SavePicture ImagePatient.Picture, txtPhoto.Text
End Sub


Private Sub EditDetails()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select tblpatientmaindetails.* from tblPatientMainDetails where Patient_ID =" & TemPatientID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        !firstname = txtFirstName.Text
        !surname = txtSurname.Text
        !othernames = txtOtherName.Text
        If IsNumeric(DataComboTitle.BoundText) Then !Title_ID = DataComboTitle.BoundText
        If IsNumeric(DataComboSex.BoundText) Then !sex_ID = DataComboSex.BoundText
        If IsNumeric(DataComboMarietal.BoundText) Then !Marital_ID = DataComboMarietal.BoundText
        If IsNumeric(DataComboRace.BoundText) Then !Race_ID = DataComboRace.BoundText
        !NICNo = txtNIC.Text
        !Address = txtAddress.Text
        !phone = txtTelephone.Text
        !fax = txtFax.Text
        !email = txtEmail.Text
        !DateOfBirth = DTPickerDOB.Value
        !notes = txtNotes.Text
        SavePhoto
        !photo = txtPhoto.Text
        .Update
        .Close
    End With

End Sub

Private Sub Grid1_Click()
    If Grid1.Row < 1 Then Exit Sub
    Grid1.Col = 0
    If Not IsNumeric(Grid1.Text) Then Exit Sub
    bttnEdit.Enabled = True
    bttnDelete.Enabled = True
    TemPatientID = Val(Grid1.Text)
    Call GetDetails
    Grid1.Col = 0
    Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPickerDOB.Value = DateAdd("yyyy", (-Val(txtAge.Text)), Now)
    End If
   If Len(txtAge.Text) = 0 And KeyAscii = 45 Then
      KeyAscii = 0
   End If
   If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtAge_LostFocus()
    If DTPickerDOB.Value = Date And IsNumeric(txtAge.Text) Then
        FromAge = True
        DTPickerDOB.Value = DateAdd("yyyy", (-Val(txtAge.Text)), Now)
    End If
End Sub


Private Sub DisplayPhoto()
    ImagePatient.Picture = LoadPicture()
    ImagePatient.Stretch = True
    On Error Resume Next
    ImagePatient.Picture = LoadPicture(txtPhoto.Text)
End Sub


Private Sub txtFax_KeyPress(KeyAscii As Integer)
   If Len(txtFax.Text) = 0 And KeyAscii = 45 Then
      KeyAscii = 0
   End If
   If KeyAscii = 32 Or KeyAscii = 43 Then Exit Sub
   If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtNIC_KeyPress(KeyAscii As Integer)
   If Len(txtNIC.Text) = 0 And KeyAscii = 45 Then
      KeyAscii = 0
   End If
   If KeyAscii = 118 Or KeyAscii = 86 Or KeyAscii = 120 Or KeyAscii = 88 Then Exit Sub
    
   If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
      KeyAscii = 0
   End If

End Sub

Private Sub txtSearchFirstName_Change()
    ClearPatientValues
    If Trim(txtSearchFirstName.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) = "" Then ListFirstNames
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub

Private Sub txtSearchID_KeyPress(KeyAscii As Integer)
    ClearPatientValues
    If KeyAscii = 13 Then
        SearchFromID
    End If
    If Len(txtSearchID.Text) = 0 And KeyAscii = 45 Then
        KeyAscii = 0
    End If
    If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSearchSurname_Change()
    ClearPatientValues
    If Trim(txtSearchSurname.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) = "" Then ListSurname
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
   If Len(txtTelephone.Text) = 0 And KeyAscii = 45 Then
      KeyAscii = 0
   End If
   If KeyAscii = 32 Or KeyAscii = 43 Then Exit Sub
   If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
      KeyAscii = 0
   End If
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

bttnRemove.BackColor = FrmBackColour
bttnRemove.ForeColor = FrmForeColour

bttnLoad.BackColor = FrmBackColour
bttnLoad.ForeColor = FrmForeColour

bttnSearch.BackColor = BttnBackColour
bttnSearch.ForeColor = BttnForeColour

bttnDelete.BackColor = BttnBackColour
bttnDelete.ForeColor = BttnForeColour


frmPatientMain.BackColor = FrmBackColour
frmPatientMain.ForeColor = FrmForeColour

FramePatientMainDetails.BackColor = FrameBackColour
FramePatientMainDetails.ForeColor = FrameForeColour

frameSearchID.BackColor = FrameBackColour
frameSearchID.ForeColor = FrameForeColour

FrameSearchNames.BackColor = FrameBackColour
FrameSearchNames.ForeColor = FrameForeColour

'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'chkCurrentlyChanneling.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour

DataComboMarietal.BackColor = TxtBackColour
DataComboMarietal.ForeColor = TxtForeColour

DataComboRace.BackColor = TxtBackColour
DataComboRace.ForeColor = TxtForeColour

DataComboSex.BackColor = TxtBackColour
DataComboSex.ForeColor = TxtForeColour

'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

DataComboTitle.BackColor = TxtBackColour
DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




Grid1.BackColor = GridBackColor
Grid1.ForeColor = GridForeColor

Grid1.BackColorBkg = GridBackColorBkg
Grid1.BackColorFixed = GridBackColorFixed
Grid1.BackColorSel = GridBackColorSel

Grid1.ForeColor = GridForeColor
Grid1.ForeColorFixed = GridForeColorFixed
Grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid



Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

Label10.BackColor = LblBackColour
Label10.ForeColor = LblForeColour
Label11.BackColor = LblBackColour
Label11.ForeColor = LblForeColour
Label12.BackColor = LblBackColour
Label12.ForeColor = LblForeColour
Label13.BackColor = LblBackColour
Label13.ForeColor = LblForeColour
Label14.BackColor = LblBackColour
Label14.ForeColor = LblForeColour
Label15.BackColor = LblBackColour
Label15.ForeColor = LblForeColour
Label16.BackColor = LblBackColour
Label16.ForeColor = LblForeColour
Label2.BackColor = LblBackColour
Label2.ForeColor = LblForeColour
Label18.BackColor = LblBackColour
Label18.ForeColor = LblForeColour
Label3.BackColor = LblBackColour
Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
Label17.BackColor = LblBackColour
Label17.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
Label5.BackColor = LblBackColour
Label5.ForeColor = LblForeColour
Label6.BackColor = LblBackColour
Label6.ForeColor = LblForeColour
Label7.BackColor = LblBackColour
Label7.ForeColor = LblForeColour

Label8.BackColor = LblBackColour
Label8.ForeColor = LblForeColour
Label9.BackColor = LblBackColour
Label9.ForeColor = LblForeColour

Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


txtAddress.BackColor = TxtBackColour
txtAddress.ForeColor = TxtForeColour

txtAge.BackColor = TxtBackColour
txtAge.ForeColor = TxtForeColour

txtEmail.BackColor = TxtBackColour
txtEmail.ForeColor = TxtForeColour
txtFax.BackColor = TxtBackColour
txtFax.ForeColor = TxtForeColour
txtFirstName.BackColor = TxtBackColour
txtFirstName.ForeColor = TxtForeColour
txtNIC.BackColor = TxtBackColour
txtNIC.ForeColor = TxtForeColour
txtNotes.BackColor = TxtBackColour
txtNotes.ForeColor = TxtForeColour
txtOtherName.BackColor = TxtBackColour
txtOtherName.ForeColor = TxtForeColour
txtPhoto.BackColor = TxtBackColour
txtPhoto.ForeColor = TxtForeColour
txtSearchFirstName.BackColor = TxtBackColour
txtSearchFirstName.ForeColor = TxtForeColour
txtSearchID.BackColor = TxtBackColour
txtSearchID.ForeColor = TxtForeColour
txtSearchSurname.BackColor = TxtBackColour
txtSearchSurname.ForeColor = TxtForeColour

txtSurname.BackColor = TxtBackColour
txtSurname.ForeColor = TxtForeColour
txtTelephone.BackColor = TxtBackColour
txtTelephone.ForeColor = TxtForeColour
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

