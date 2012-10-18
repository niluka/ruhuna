VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProgramPreferance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Preferance"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
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
   ScaleHeight     =   5745
   ScaleWidth      =   5265
   Begin VB.Frame Frame2 
      Caption         =   "Short Date Format"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5055
      Begin VB.ComboBox cmbShortDateFormat 
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Text            =   "cmbShortDateFormat"
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblShortDate 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Long Date Format"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbLongDateFormat 
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Text            =   "cmbLongDateFormat"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblLongDate 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame21 
      Caption         =   "Database"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtDatabase 
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin btButtonEx.ButtonEx bttnSelectDatabasePath 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   12648384
         Caption         =   "Select Database"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   12648384
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
End
Attribute VB_Name = "frmProgramPreferance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SaveToMemory()
    LongDateFormat = cmbLongDateFormat.Text
    ShortDateFormat = cmbShortDateFormat.Text
    Database = txtDatabase.Text
End Sub

Private Sub SaveToReg()
    SaveSetting App.EXEName, "Options", "LongDateFormat", cmbLongDateFormat.Text
    SaveSetting App.EXEName, "Options", "ShortDateFormat", cmbShortDateFormat.Text
    SaveSetting App.EXEName, "Options", "Database", txtDatabase.Text
End Sub

Private Sub GetFromReg()
    txtDatabase.Text = Database
    cmbLongDateFormat.Text = GetSetting(App.EXEName, "Options", "LongDateFormat", "dd MMMM yyyy")
    cmbShortDateFormat.Text = GetSetting(App.EXEName, "Options", "ShortDateFormat", "dd MM yy")
End Sub

Private Sub cmbLongDateFormat_Change()
    lblLongDate.Caption = Format(Date, cmbLongDateFormat.Text)
End Sub

Private Sub cmbShortDateFormat_Change()
    lblShortDate.Caption = Format(Date, cmbShortDateFormat.Text)
End Sub

Private Sub Form_Load()
    Call FillData
    Call GetFromReg
End Sub

Private Sub bttnSelectDatabasePath_Click()
    On Error Resume Next
    CommonDialog1.FileName = Database
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "Lakmedipro Database|eStore.mdb"
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.Text = CommonDialog1.FileName
        SaveSetting App.EXEName, "Options", "DatabaseLocation", txtDatabase.Text
        Database = txtDatabase.Text
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
        bttnSelectDatabasePath.SetFocus
    End If
End Sub


Private Sub FillData()
With cmbLongDateFormat
    .AddItem "dd MMMM yyyy"
    .AddItem "dddd, dd MMMM yyyy"
    .AddItem "dd MMM yyyy"
    .AddItem "yyyy MMMM dd"
    .AddItem "yyyy MMMM dd , dddd"
    .AddItem "yyyy MMM dd"
    .AddItem "MMMM dd yyyy"
    .AddItem "dddd, MMMM dddd yyyy"
    .AddItem "MMM dd yyyy"
End With
With cmbShortDateFormat
    .AddItem "dd MM yy"
    .AddItem "d M yy"
    .AddItem "yy MM dd"
    .AddItem "yy M d"
    .AddItem "MM dd yy"
    .AddItem "M d yy"
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim tr As Integer
'    If FSys.FileExists(txtDatabase.Text) = False Then
'        tr = MsgBox("You have not Selected a valid database", vbCritical, "Database?")
'        bttnSelectDatabasePath.SetFocus
'        Cancel = True
'        Exit Sub
'    End If
    Call SaveToMemory
    Call SaveToReg
'    If Not IsNumeric(txtHighRate.Text) Or Val(txtHighRate.Text) > 100 Then
'        tr = MsgBox("You have not Selected a valid high rate for Insurance Patients", vbCritical, "High Rate?")
'        txtHighRate.SetFocus
'        Cancel = True
'        Exit Sub
'    End If
End Sub

