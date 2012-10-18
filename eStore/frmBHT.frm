VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9465
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtPatientID 
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtBHT 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin MSComCtl2.MonthView mvDOA 
         Height          =   2370
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   72089601
         CurrentDate     =   39589
      End
      Begin MSComCtl2.MonthView mvDOD 
         Height          =   2370
         Left            =   6360
         TabIndex        =   14
         Top             =   840
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   72089601
         CurrentDate     =   39589
      End
      Begin MSComCtl2.DTPicker dtpTOA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hour MIN sec"
         Format          =   72089602
         CurrentDate     =   39589
      End
      Begin MSComCtl2.DTPicker dtpTOD 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   16
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hour MIN sec"
         Format          =   72089602
         CurrentDate     =   39589
      End
      Begin VB.Label Label6 
         Caption         =   "Time of Discharge"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Date of Discharge"
         Height          =   495
         Left            =   4800
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Time of Admission"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date of Admission"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Patient ID"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BHT"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBHT As New ADODB.Recordset
    Dim rsViewBHT As New ADODB.Recordset
    Dim temSql As String
    Dim temDOA As Date
    Dim temDOD As Date
    Dim temTOA As Date
    Dim temTOD As Date
    Dim PatientID As Long

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnSave.Enabled = False
    bttnCancel.Enabled = False
    bttnClose.Enabled = True
    
    txtBHT.Enabled = False
    txtPatientID.Enabled = False
    mvDOA.Enabled = False
    mvDOD.Enabled = False
    dtpTOA.Enabled = False
    dtpTOD.Enabled = False
    
    bttnSave.Visible = False
End Sub

Private Sub AfterAdd()
    bttnAdd.Enabled = True
    bttnSave.Enabled = True
    bttnCancel.Enabled = True
    bttnClose.Enabled = True
    
    txtBHT.Enabled = True
    txtPatientID.Enabled = True
    mvDOA.Enabled = True
    mvDOD.Enabled = True
    dtpTOA.Enabled = True
    dtpTOD.Enabled = True
    
    bttnSave.Visible = True
End Sub

Private Sub ClearValues()
    txtBHT.Text = Empty
    txtPatientID = Empty
End Sub

Private Sub SaveDetails()
    On Error Resume Next

    With rsViewBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !BHT = txtBHT.Text
        !PatientID = txtPatientID.Text
        !DOA = mvDOA.Value
        !DOD = mvDOD.Value
        !TOA = dtpTOA.Value
        !TOD = dtpTOD.Value
        .Update
        .Close
    End With
End Sub

Private Sub bttnAdd_Click()
    Call AfterAdd
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub Form_Load()
    Call BeforeAddEdit
End Sub
