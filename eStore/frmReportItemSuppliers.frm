VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReportItemSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Suppliers"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
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
   ScaleHeight     =   8130
   ScaleWidth      =   12930
   Begin VB.OptionButton bySupplier 
      Caption         =   "Order by Supplier"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   7560
      Width           =   2535
   End
   Begin VB.OptionButton optByItem 
      Caption         =   "Order By Item"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7560
      Value           =   -1  'True
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx btnCLose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   11520
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Default         =   -1  'True
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   12726
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmReportItemSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItemSuppliers As New ADODB.Recordset
    Dim temSQL As String
    Dim P As New cSetDfltPrinter
    Dim i As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    P.SetPrinterAsDefault (ReportPrinterName)
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            
            With dtrItemSuppliers
                Set .DataSource = rsItemSuppliers
                
                .Sections.Item("Section4").Controls("lblName").Caption = HospitalName
                .Sections.Item("Section4").Controls("lblCOntact").Caption = HospitalDescreption
                .Sections.Item("Section4").Controls("lbltopic").Caption = "Item Suppliers"
                
                .Sections("Section1").Controls("txtItem").DataField = "DIsplay"
                .Sections("Section1").Controls("txtSupplier").DataField = "DistributorName"
                
                .Show
                
                
            
            
            End With
            
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

End Sub

Private Sub FillGrid()
    Grid.Clear
    Grid.Rows = 2
    Grid.Cols = 2
    
    With rsItemSuppliers
        If .State = 1 Then .Close
        temSQL = "SELECT tblItem.Display, tblDistrubutor.DistributorName " & _
                    "FROM (tblItem LEFT JOIN tblItemDistributor ON tblItem.ItemID = tblItemDistributor.ItemID) LEFT JOIN tblDistrubutor ON tblItemDistributor.DistributorID = tblDistrubutor.DistributorID "
        If optByItem.Value = True Then
              temSQL = temSQL & " ORDER BY tblItem.Display"
        Else
                temSQL = temSQL & " ORDER BY tblDistrubutor.DistributorName"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            
            Grid.Cols = 2
            Grid.Rows = .RecordCount + 1
            Grid.ColWidth(0) = 5000
            Grid.ColWidth(1) = Grid.Width - Grid.ColWidth(0) - 150
            
            i = 0
            While .EOF = False
                i = i + 1
                Grid.TextMatrix(i, 0) = !Display
                If Not IsNull(!DistributorName) Then Grid.TextMatrix(i, 1) = !DistributorName
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub bySupplier_Click()
    FillGrid
End Sub

Private Sub Form_Load()
    Call FillGrid
End Sub

Private Sub optByItem_Click()
    FillGrid
End Sub
