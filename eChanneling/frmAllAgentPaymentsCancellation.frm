VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAllAgentPaymentsCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Agents Payment Cancellation"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
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
   ScaleHeight     =   5880
   ScaleWidth      =   11400
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   6960
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin btButtonEx.ButtonEx bttnFill 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Fill"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnPrint 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnClose 
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71237635
         CurrentDate     =   39515
      End
      Begin MSComCtl2.DTPicker dtpTO 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   71237635
         CurrentDate     =   39515
      End
      Begin VB.Label lblFromDate 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblToDate 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBills 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAllAgentPaymentsCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsRefill As New ADODB.Recordset
    Dim CSetPrinter As New cSetDfltPrinter

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnFill_Click()
    Call FillGrid
End Sub

Private Sub bttnPrint_Click()
    On Error Resume Next
    
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)

    
If SelectForm(BillPaperName, Me.hwnd) = 1 Then
    Dim i As Integer
    Dim Tab1 As Integer
    Dim Tab2 As Integer
    Dim Tab3 As Integer
    Dim Tab4 As Integer
    Dim TabName As Integer
    Dim TabAddress As Integer
    Dim TabPhone As Integer
    Dim Tab11 As Integer
    Dim Tab12 As Integer
    Dim Tab13 As Integer
    Dim Tab14 As Integer
       
    Tab1 = 5
    Tab2 = 50
    Tab3 = 70
    Tab4 = 90
    
    TabName = 40
    TabAddress = 35
    TabPhone = 45
    
    Tab11 = 5
    Tab12 = 50
    Tab13 = 70
    Tab14 = 90
    
    Printer.FontName = "Tahoma"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print Tab(TabName); InstitutionName
    Printer.FontBold = False
    Printer.FontName = "Tahoma"
    Printer.FontSize = 12
    Printer.Print Tab(TabAddress); InstitutionAddress
    Printer.Print Tab(TabPhone); InstitutionTelephone
    Printer.Print
    
    Printer.Print Tab(Tab11); "Agent Name";
    Printer.Print Tab(Tab12); "Ref. No";
    Printer.Print Tab(Tab13); "Cancel Date";
    Printer.Print Tab(Tab14); "Cancel Amount"
    
    
    With gridBills
        For i = 1 To .Rows - 1
            Printer.Print Tab(Tab1); Left(.TextMatrix(i, 0), 30);
            Printer.Print Tab(Tab2); .TextMatrix(i, 1);
            Printer.Print Tab(Tab3); .TextMatrix(i, 2);
            Printer.Print Tab(Tab4); Right(Space(12) & .TextMatrix(i, 3), 12)
        Next i
    End With
    Printer.EndDoc
End If
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FormatGrid
End Sub

Private Sub FormatGrid()
    With gridBills
        .Clear
        .Rows = 1
        .Cols = 4
        
        .Col = 0
        .Text = "Institution Name"
        
        .Col = 1
        .Text = "Referance No"
        
        .Col = 2
        .Text = "Date"
        
        .Col = 3
        .Text = "Amount"
        
        .ColWidth(0) = 4000
        .ColWidth(1) = 3000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        
    End With
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim i As Integer
    With rsRefill
        If .State = 1 Then .Close
        temSql = "SELECT tblInstitutions.InstitutionName, tblStaff.StaffName, tblAgentPaymentCancellation.Date, tblAgentPaymentCancellation.RefNo , tblAgentPaymentCancellation.Amount " & _
                    "FROM (tblAgentPaymentCancellation LEFT JOIN tblInstitutions ON tblAgentPaymentCancellation.AgentID = tblInstitutions.Institution_ID) LEFT JOIN tblStaff ON tblAgentPaymentCancellation.UserID = tblStaff.Staff_ID " & _
                    "WHERE (((tblAgentPaymentCancellation.Date) Between '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "'))"
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            gridBills.Rows = gridBills.Rows + 1
            gridBills.Row = gridBills.Rows - 1
            gridBills.Col = 0
            gridBills.CellAlignment = 1
            gridBills.Text = !InstitutionName
            gridBills.Col = 1
            gridBills.CellAlignment = 7
            gridBills.Text = Format(!RefNo, "")
            gridBills.Col = 2
            gridBills.CellAlignment = 1
            gridBills.Text = Format(!Date, "dd MMMM yyyy")
            gridBills.Col = 3
            gridBills.CellAlignment = 1
            gridBills.Text = Format(!Amount, "0.00")

            .MoveNext
        Wend
    End If
        gridBills.Row = 0
        .Close
    End With
End Sub

Private Sub gridBills_Click()
    With gridBills
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub


