VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportStockAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Adjustment Report"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6630
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
   ScaleHeight     =   4920
   ScaleWidth      =   6630
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin MSDataListLib.DataCombo dtcDepts 
         Height          =   360
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.OptionButton optAllDepts 
         Caption         =   "All Departments"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optSelectedDepts 
         Caption         =   "Selected Department"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Duration"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6255
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   67239939
         CurrentDate     =   29224
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   67239939
         CurrentDate     =   29224
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Staff Members"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   6255
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.OptionButton optAllStaff 
         Caption         =   "All Staff Members"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optSelectedStaff 
         Caption         =   "Selected Staff Member"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   4200
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
Attribute VB_Name = "frmReportStockAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStore As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsAdjustment As New ADODB.Recordset
    Dim temSql As String
    Dim CSetPrinter As New cSetDfltPrinter

    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim tr As Integer
    Dim RetVal As Integer
    If optSelectedStaff.Value = True And IsNumeric(dtcStaff.BoundText) = False Then
        MsgBox "Staff Member?"
        dtcStaff.SetFocus
        Exit Sub
    End If
    If optSelectedDepts.Value = True And IsNumeric(dtcDepts.BoundText) = False Then
        MsgBox "Department?"
        dtcDepts.SetFocus
        Exit Sub
    End If

    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            tr = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
        temSql = "SELECT tblAdjustment.Date, tblAdjustment.Time, tblStaff.Name, tblAdjustmentCategory.AdjustmentCategory, tblItem.Display, tblBatch.Batch, tblAdjustment.Amount, tblIssueUnit.IssueUnit " & _
                    "FROM (tblAdjustmentCategory RIGHT JOIN (((tblAdjustment LEFT JOIN tblItem ON tblAdjustment.ItemID = tblItem.ItemID) LEFT JOIN tblBatch ON tblAdjustment.BatchID = tblBatch.BatchID) LEFT JOIN tblStaff ON tblAdjustment.StaffID = tblStaff.StaffID) ON tblAdjustmentCategory.AdjustmentCategoryID = tblAdjustment.CategoryID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID " & _
                    "WHERE tblAdjustment.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If optSelectedStaff.Value = True Then temSql = temSql & " AND tblAdjustment.StaffID = " & Val(dtcStaff.BoundText) & " "
        If optSelectedDepts.Value = True Then temSql = temSql & " AND tblAdjustment.StoreID = " & Val(dtcStaff.BoundText) & " "
        With rsAdjustment
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With dtrReport1
            Set .DataSource = rsAdjustment
            .Sections("Section1").Controls("txtDate").DataField = "Date"
            .Sections("Section1").Controls("txtTime").DataField = "Time"
            .Sections("Section1").Controls("txtStaff").DataField = "Name"
            .Sections("Section1").Controls("txtItem").DataField = "Display"
            .Sections("Section1").Controls("txtBatch").DataField = "Batch"
            .Sections("Section1").Controls("txtQty").DataField = "Amount"
            .Sections("Section1").Controls("txtUnit").DataField = "IssueUnit"
            
            .Show





        End With


        Case FORM_ADDED   ' 2
            tr = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
    
End Sub

Private Sub optAllDepts_Click()
    If optSelectedDepts.Value = True Then
        dtcDepts.Enabled = True
    Else
        dtcDepts.Enabled = False
        dtcDepts.Text = Empty
    End If
End Sub

Private Sub optAllStaff_Click()
    If optAllStaff.Value = True Then
        dtcStaff.Enabled = False
        dtcStaff.Text = Empty
    Else
        dtcStaff.Enabled = True
    End If
End Sub

Private Sub optSelectedDepts_Click()
    If optSelectedDepts.Value = True Then
        dtcDepts.Enabled = True
    Else
        dtcDepts.Enabled = False
        dtcDepts.Text = Empty
    End If
End Sub

Private Sub optSelectedStaff_Click()
    If optAllStaff.Value = True Then
        dtcStaff.Enabled = False
        dtcStaff.Text = Empty
    Else
        dtcStaff.Enabled = True
    End If

End Sub

Private Sub FillCombos()
    With rsStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDepts
        Set .RowSource = rsStore
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
End Sub

