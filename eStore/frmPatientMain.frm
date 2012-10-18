VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatientMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Main Details"
   ClientHeight    =   9420
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
   ScaleHeight     =   9420
   ScaleWidth      =   11190
   Begin VB.Frame framPatient 
      Height          =   4815
      Left            =   120
      TabIndex        =   52
      Top             =   3240
      Width           =   4215
      Begin MSDataListLib.DataCombo dtcPatient 
         Height          =   4500
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7938
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   4560
      TabIndex        =   48
      Top             =   8040
      Width           =   6495
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Height          =   375
         Left            =   4800
         TabIndex        =   51
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   44
      Top             =   8040
      Width           =   4215
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2880
         TabIndex        =   46
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnDelete 
         Height          =   375
         Left            =   1560
         TabIndex        =   47
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      TabIndex        =   27
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtNotes 
         Height          =   825
         Left            =   2040
         TabIndex        =   20
         Top             =   7080
         Width           =   4215
      End
      Begin VB.TextBox txtEmail 
         Height          =   345
         Left            =   2040
         TabIndex        =   19
         Top             =   6600
         Width           =   4215
      End
      Begin VB.TextBox txtFax 
         Height          =   345
         Left            =   2040
         TabIndex        =   18
         Top             =   6120
         Width           =   4215
      End
      Begin VB.TextBox txtTelephone 
         Height          =   345
         Left            =   2040
         TabIndex        =   17
         Top             =   5640
         Width           =   4215
      End
      Begin VB.TextBox txtAddress 
         Height          =   825
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   4680
         Width           =   4215
      End
      Begin VB.TextBox txtNIC 
         Height          =   345
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   11
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtSurname 
         Height          =   345
         Left            =   2040
         TabIndex        =   6
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtOtherName 
         Height          =   345
         Left            =   2040
         TabIndex        =   5
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtFirstName 
         Height          =   345
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo DataComboTitle 
         Bindings        =   "frmPatientMain.frx":0000
         Height          =   360
         Left            =   2040
         TabIndex        =   7
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   10
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
         TabIndex        =   14
         Top             =   4200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   39413
      End
      Begin btButtonEx.ButtonEx bttnLoad 
         Height          =   255
         Left            =   3960
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   42
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date &of Birth"
         Height          =   255
         Left            =   240
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   7200
         Width           =   3615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "E-&Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   6120
         Width           =   3615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "&Telephone"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "NIC &No."
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Race"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marietal"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Se&x"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Title"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Surname"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Othe&r Names"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "First &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   3615
      End
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
      TabIndex        =   23
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
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameSearchID 
      Caption         =   "Search By ID"
      Height          =   975
      Left            =   120
      TabIndex        =   22
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
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9360
      TabIndex        =   21
      Top             =   8880
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
   Begin VB.TextBox txtPhoto 
      Height          =   4680
      Left            =   120
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   9255
   End
End
Attribute VB_Name = "frmPatientMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsviewPatient As New ADODB.Recordset
Dim rsTem As New ADODB.Recordset
Dim TemSql As String

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
        FillPatientCombo
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


Private Sub dtcPatient_Click(Area As Integer)
If IsNumeric(dtcPatient.BoundText) = False Then Exit Sub
Call GetDetails
End Sub

Private Sub DTPickerDOB_Change()
If FromAge = True Then Exit Sub
'    txtAge.Text = CalculateAgeInWords(DTPickerDOB.Value)
'    FromAge = False
End Sub

Private Sub Form_Load()
Call FillPatientCombo
    Call BeforeAddEdit
    DTPickerDOB.Value = Date
End Sub

Private Sub BeforeAddEdit()
    TemPatientID = Empty
    Call ClearSearchValues
    Call ClearPatientValues
    frameSearchID.Enabled = True
    FrameSearchNames.Enabled = True
    FramePatientMainDetails.Enabled = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnDelete.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    framPatient.Enabled = True

End Sub


Private Sub AfterAdd()
    Call ClearSearchValues
    Call ClearPatientValues
    frameSearchID.Enabled = False
    FrameSearchNames.Enabled = False
    FramePatientMainDetails.Enabled = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    framPatient.Enabled = False
End Sub

Private Sub AfterEdit()
    Call ClearSearchValues
    frameSearchID.Enabled = False
    FrameSearchNames.Enabled = False
    FramePatientMainDetails.Enabled = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    framPatient.Enabled = False

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

Private Sub FillPatientCombo()
    Dim NowROw As Long
    
    With rsviewPatient
    If .State = 1 Then .Close
    TemSql = "Select* From tblPatientMainDetails Order By FirstName"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcPatient.RowSource = rsviewPatient
    dtcPatient.BoundColumn = "patient_ID"
    dtcPatient.ListField = "firstname"

    End With
    
End Sub


Private Sub SearchFromID()
Dim TemResponce  As Integer

If Not IsNumeric(txtSearchID.Text) Then Exit Sub

With rsviewPatient
    If .State = 1 Then .Close
    TemSql = "SELECT * FROM tblPatientmainDetails where Patient_ID = " & Val(txtSearchID.Text) & " order by Patient_ID"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    
    If .RecordCount = 0 Then
    

        TemResponce = MsgBox("There is no record with the patient ID of " & txtSearchID.Text, vbCritical, "Wrong ID")
        txtSearchID.SetFocus
        ClearPatientValues
        SendKeys "{Home}+{end}"
        Call FillPatientCombo
        
    Else
        Set dtcPatient.RowSource = rsviewPatient
        dtcPatient.BoundColumn = "patient_ID"
        dtcPatient.ListField = "firstname"
  
    End If
    
End With
    
End Sub

Private Sub ListFirstNames()
    Dim NowROw As Long
    
     With rsviewPatient
        If .State = 1 Then .Close
        TemSql = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where firstname like '" & txtSearchFirstName.Text & "%' order by FIrstname"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcPatient.RowSource = rsviewPatient
        dtcPatient.BoundColumn = "patient_ID"
        dtcPatient.ListField = "firstname"
        
    End With
End Sub

Private Sub ListSurname()
    With rsviewPatient
        If .State = 1 Then .Close
        TemSql = "SELECT * FROM tblPatientmainDetails where surname like '" & txtSearchSurname.Text & "%' order by surname"
        .Open , TemSql, adOpenStatic, adLockReadOnly
        
        Set dtcPatient.RowSource = rsviewPatient
        dtcPatient.BoundColumn = "patient_ID"
        dtcPatient.ListField = "Surname"
        
    End With
End Sub

Private Sub ListBothNames()

    With rsviewPatient
        If .State = 1 Then .Close
        TemSql = "SELECT  patient_ID, surname, firstname, (surname + firstname) as Temname FROM tblPatientmainDetails where (surname like '" & txtSearchSurname.Text & "%') and ( firstname like '" & txtSearchFirstName.Text & "%')order by FIrstname,surname"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcPatient.RowSource = rsviewPatient
        dtcPatient.BoundColumn = "patient_ID"
        dtcPatient.ListField = "Temname"
        
    End With
End Sub

Private Sub bttnAdd_Click()
    Call ClearSearchValues
    Call ClearPatientValues
    Call AfterAdd
End Sub

Private Sub GetDetails()
    Call ClearPatientValues
    With rsTem
        If .State = 1 Then .Close
        TemSql = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where (patient_ID =" & Val(dtcPatient.BoundText) & ")"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
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
        If .State = 1 Then .Close
    End With
End Sub

Private Sub SaveDetails()
    With rsTem
        If .State = 1 Then .Close
        TemSql = "Select* from tblPatientMainDetails"
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        
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
       
        If .State = 1 Then .Close
        Call FillPatientCombo
        
    End With
End Sub

Private Sub SavePhoto()
    On Error Resume Next
    SavePicture ImagePatient.Picture, txtPhoto.Text
End Sub


Private Sub EditDetails()

    With rsTem
        If .State = 1 Then .Close
        TemSql = "Select* from tblPatientMainDetails where Patient_ID =" & Val(dtcPatient.BoundText) & ""
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        
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
        
        If .State = 1 Then .Close
        
        Call FillPatientCombo
    End With

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


