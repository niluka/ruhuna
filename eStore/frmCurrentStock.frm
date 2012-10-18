VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCurrentStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
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
   ScaleHeight     =   8505
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   8
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid GridStock 
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order by"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton optValue 
         Caption         =   "Value"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optQty 
         Caption         =   "Quentity"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   7920
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
      Left            =   7080
      TabIndex        =   0
      Top             =   7920
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
   Begin VB.Label Label1 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmCurrentStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim iLen As Long
    Dim qLen As Long
    Dim vLen As Long
    Dim rsItemStock As New ADODB.Recordset
    Dim temSql As String
    Dim TemStr1 As String
    Dim TemStr2 As String
    Dim TemStr3 As String
    Dim CsetPrinter As New cSetDfltPrinter
    Dim rsTem As New ADODB.Recordset
    Dim rsViewCategory As New ADODB.Recordset

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim TemResponce As Long
    Dim RetVal As Integer
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            Call SelectPrint
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select
End Sub
    
Private Sub SelectPrint()
       Dim i As Integer
       With rsTem
            If .State = 1 Then .Close
            temSql = "Delete from tblTemReport1"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .State = 1 Then .Close
            temSql = "SELECT * from tblTemReport1"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            For i = 1 To GridStock.Rows - 1
                 .AddNew
                 !txt1 = GridStock.TextMatrix(i, 0)
                 !txt2 = GridStock.TextMatrix(i, 1)
                 !Double1 = Val(GridStock.TextMatrix(i, 2))
                 !txt4 = GridStock.TextMatrix(i, 3)
                 .Update
            Next
        End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblTemReport1 where txt1 <> '' and double1 <> 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrTemReport1
        Set .DataSource = rsTem
        .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
        .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Total Stock"
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = "Date " & Format(Date, "dd MMMM yyyy")
        .Show
    End With
End Sub

Private Sub cmbCategory_Change()
    Call FormatGrid
    Call FillItems
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillItems
    Call FormatGrid
    Call FillItems
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        temSql = "Select * form tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsViewCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
End Sub

Private Sub FormatGrid()
    With GridStock
        .Clear
        
        .Cols = 4
        .Rows = 1
    
        .FixedCols = 0
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Item"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Stock (In Issue Units)"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Purchase Value"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Sale Value"
        
        .ColWidth(0) = 2700
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
    
    End With
End Sub

Private Sub FillItems()
    With rsItemStock
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "SELECT tblItem.Display, tblItem.ItemID, tblBatchStock.Stock " & _
                        "FROM (tblItem LEFT JOIN tblBatch ON tblItem.ItemID = tblBatch.ItemID) LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "Where tblItem.ItemCategoryID = " & Val(cmbCategory.BoundText) & " " & _
                        "GROUP BY tblItem.Display, tblBatchStock.Stock, tblItem.ItemID"
        Else
            temSql = "SELECT tblItem.Display, tblItem.ItemID, tblBatchStock.Stock " & _
                        "FROM (tblItem LEFT JOIN tblBatch ON tblItem.ItemID = tblBatch.ItemID) LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "GROUP BY tblItem.Display, tblBatchStock.Stock, tblItem.ItemID"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            GridStock.Rows = .RecordCount + 1
            .MoveFirst
            Dim i As Integer
            Dim NewItem As New Item
            While .EOF = False
                If IsNull(!Stock) = False Then
                    If !Stock > 0 Then
                        i = i + 1
                        If IsNull(!Display) = False Then
                            GridStock.TextMatrix(i, 0) = !Display
                        End If
                        NewItem.ID = !ItemID
                        GridStock.TextMatrix(i, 1) = !Stock
                        GridStock.TextMatrix(i, 2) = Format(!Stock * NewItem.PPrice, "0.00")
                        GridStock.TextMatrix(i, 3) = Format(!Stock * NewItem.SPrice, "0.00")
                    End If
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub
