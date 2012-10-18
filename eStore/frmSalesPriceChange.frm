VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSalesPriceChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Price Change"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12150
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
   ScaleHeight     =   7845
   ScaleWidth      =   12150
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   4920
      TabIndex        =   14
      Top             =   120
      Width           =   7095
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   6480
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Cancel"
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
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   6480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Update"
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
      Begin VB.TextBox txtNewSalesPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox txtCSalesPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   5520
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   4695
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpPTo 
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72941569
         CurrentDate     =   39542
      End
      Begin MSComCtl2.DTPicker dtpPFrom 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72941569
         CurrentDate     =   39542
      End
      Begin VB.Label Label3 
         Caption         =   "&From"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "&To"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblP1 
         Caption         =   "Per Unit"
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label lblP2 
         Caption         =   "Per Unit"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   6000
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Enter &New Price"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Current sales Price"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   5520
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin MSDataListLib.DataCombo dtcItem 
         Height          =   5220
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9208
         _Version        =   393216
         Style           =   1
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCategory 
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "&Item"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Category"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
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
End
Attribute VB_Name = "frmSalesPriceChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewCatogery As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset
    Dim rsViewCatItem As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim NewItem As New Item
    Dim temSql As String
    Dim A As Long
    
Private Sub bttnCancel_Click()
    bttnUpdate.Enabled = False
    Frame1.Enabled = True
    bttnCancel.Visible = False
    Call ClearVales
    dtcCategory.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnUpdate_Click()
    If txtNewSalesPrice.Text = Empty Then A = MsgBox("Enter New Sales Price", vbCritical + vbOKOnly, "Error"): Exit Sub
    Call SaveCurrentSalePrice
    bttnUpdate.Enabled = False
    Frame1.Enabled = True
    bttnCancel.Visible = False
    Call ClearVales
    Call FillSalesPrice(Val(dtcItem.BoundText))
    dtcCategory.SetFocus
End Sub

Private Sub ClearVales()
    txtCSalesPrice.Text = Empty
    txtNewSalesPrice.Text = Empty
    GridSPrice.Enabled = True
End Sub

Private Sub SaveCurrentSalePrice()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select* From tblSalePrice"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ItemID = Val(dtcItem.BoundText)
            !SPrice = Val(txtNewSalesPrice.Text)
            !setdate = Date
            !SetTime = Now
            !StaffID = UserID
            .Update
        If .State = 1 Then .Close
        temSql = "Select* From tblCurrentSalePrice Where ItemID = " & Val(dtcItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount < 1 Then
                .AddNew
                !ItemID = Val(dtcItem.BoundText)
                !SPrice = Val(txtNewSalesPrice.Text)
                !setdate = Date
                !SetTime = Now
                !StaffID = UserID
                .Update
            ElseIf .RecordCount = 1 Then
                !SPrice = Val(txtNewSalesPrice.Text)
                !setdate = Date
                !SetTime = Now
                !StaffID = UserID
                .Update
            Else
                If .State = 1 Then .Close
                temSql = "Delete From tblCurrentSalePrice Where ItemID = " & Val(dtcItem.BoundText)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .State = 1 Then .Close
                temSql = "Select * From tblCurrentSalePrice Where ItemID = " & Val(dtcItem.BoundText)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                .AddNew
                !ItemID = Val(dtcItem.BoundText)
                !SPrice = Val(txtNewSalesPrice.Text)
                !setdate = Date
                !SetTime = Now
                !StaffID = UserID
                .Update
            End If
        If .State = 1 Then .Close
    End With
End Sub


Private Sub dtcCategory_Click(Area As Integer)
    If IsNumeric(dtcCategory.BoundText) = False Then Exit Sub
    Call FillCategoryItem
End Sub

Private Sub dtcItem_Click(Area As Integer)
    If IsNumeric(dtcItem.BoundText) = False Then bttnUpdate.Enabled = False: Exit Sub
    Call FillSalesPrice(Val(dtcItem.BoundText))
    NewItem.ID = Val(dtcItem.BoundText)
    lblP1.Caption = "per " & NewItem.IUnit
    lblP2.Caption = lblP1.Caption
    bttnUpdate.Enabled = True
End Sub

Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtNewSalesPrice.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCategory
    Call FillItem
    dtpPFrom.Value = Date
    dtpPTo.Value = Date
    bttnUpdate.Enabled = False
End Sub

Private Sub FillCategory()
    With rsViewCatogery
        If .State = 1 Then .Close
        temSql = "SELECT tblItemCategory.ItemCategoryID, tblItemCategory.ItemCategory FROM tblItemCategory ORDER BY tblItemCategory.ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        Set dtcCategory.RowSource = rsViewCatogery
        dtcCategory.ListField = "ItemCategory"
        dtcCategory.BoundColumn = "ItemCategoryID"
    End With

End Sub

Private Sub FillItem()
    With rsViewItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        Set dtcItem.RowSource = rsViewItem
        dtcItem.ListField = "display"
        dtcItem.BoundColumn = "ItemID"
    End With

End Sub

Private Sub FillCategoryItem()
    With rsViewItem
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblItem Where (ItemCategoryID = " & dtcCategory.BoundText & ") ORDER BY display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        dtcItem.Text = Empty
        Set dtcItem.RowSource = rsViewItem
        dtcItem.ListField = "display"
        dtcItem.BoundColumn = "ItemID"
    End With

End Sub

Private Sub FillSalesPrice(ByVal ItemID As Long)
Dim i As Long
With GridSPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Sale Price"
    
    .ColWidth(0) = (.Width - 150) / 2
    .ColWidth(1) = (.Width - 150) / 2
    
End With

With rsTem
    If .State = 1 Then .Close
    temSql = "SELECT tblSalePrice.SetDate, tblSalePrice.SPrice FROM tblSalePrice WHERE tblSalePrice.ItemID=" & ItemID & " ORDER BY tblSalePrice.SetDate DESC , tblSalePrice.SetTime DESC "
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        Do While .EOF = False
            With GridSPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTem!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTem!SPrice, "#,#00.00")
            End With
        .MoveNext
        Loop
    End If
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsViewCatogery.State = 1 Then rsViewCatogery.Close
    If rsViewItem.State = 1 Then rsViewItem.Close
    If rsViewCatItem.State = 1 Then rsViewCatItem.Close
    If rsTem.State = 1 Then rsTem.Close
End Sub

Private Sub GridSPrice_Click()
    Frame1.Enabled = False
    bttnCancel.Visible = True
    With GridSPrice
    .Col = 1
    txtCSalesPrice.Text = Format(.Text, "0.00")
    
    .Col = 0
    .ColSel = .Cols - 1
    .Enabled = False
    End With
End Sub

Private Sub txtNewSalesPrice_Change()
    If IsNumeric(txtNewSalesPrice.Text) = True Then
    bttnUpdate.Enabled = True
    Else
    bttnUpdate.Enabled = False
    End If
End Sub

Private Sub txtNewSalesPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnUpdate_Click
    End If
End Sub

Private Sub txtNewSalesPrice_LostFocus()
    With GridSPrice
        .Col = 1
        .Text = txtNewSalesPrice.Text
        .Text = Format(.Text, "0.00")
    End With
End Sub
