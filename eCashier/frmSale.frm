VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
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
   ScaleHeight     =   6240
   ScaleWidth      =   11625
   Begin VB.TextBox txtDeleteID 
      Height          =   360
      Left            =   9480
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtEditID 
      Height          =   360
      Left            =   8520
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtSaleBillID 
      Height          =   360
      Left            =   10440
      TabIndex        =   27
      Top             =   5655
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5280
      Width           =   3255
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   9120
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   19
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   17
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   15
      Top             =   2760
      Width           =   2535
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7011
      _Version        =   393216
   End
   Begin VB.TextBox txtTotalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtHospitalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtProfessionalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8880
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   360
      Left            =   8880
      TabIndex        =   3
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   8880
      TabIndex        =   5
      Top             =   4200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Settle"
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Discount"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Total Charge"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Hospital Charge"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Professioanl Charge"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Staff"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Service"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    Dim FSys As New Scripting.FileSystemObject
    Private CSetPrinter As New cSetDfltPrinter
    Dim rsIxList As New ADODB.Recordset
    Dim i As Integer
    Dim rsPastVisits As New ADODB.Recordset

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings

End Sub

Private Sub FillCombos()
    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True

End Sub

Private Sub GetSettings()
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    'cmbPaper.Clear
    CSetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = 1440 * 8
            .cy = 1400 * 11
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
