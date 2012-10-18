VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchasePriceChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Price Change"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12135
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4695
      Begin MSDataListLib.DataCombo dtcItem 
         Height          =   5220
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9208
         _Version        =   393216
         Style           =   1
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dtcCategory 
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
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
         Caption         =   "Categ&ory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "&Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   4920
      TabIndex        =   15
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtCPurchasePrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtNewPurchasePrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   5760
         Width           =   1695
      End
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
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   4215
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker dtpPTo 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   72941569
         CurrentDate     =   39542
      End
      Begin MSComCtl2.DTPicker dtpPFrom 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   72941569
         CurrentDate     =   39542
      End
      Begin VB.Label lblP2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label lblP1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "&To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Current Purchase Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   5280
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Enter &New Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   5760
         Width           =   2775
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
Attribute VB_Name = "frmPurchasePriceChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewCatogery As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset
    Dim rsViewCatItem As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    Dim A As Long
    Dim NewItem As New Item
    
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
    If txtNewPurchasePrice.Text = Empty Then A = MsgBox("Enter New Purchase Price", vbCritical + vbOKOnly, "Error"): Exit Sub
    Call SavePurchasePrice
    bttnUpdate.Enabled = False
    Frame1.Enabled = True
    bttnCancel.Visible = False
    Call ClearVales
    Call FillPurchasePrice(Val(dtcItem.BoundText))
    dtcCategory.SetFocus
End Sub

Private Sub ClearVales()
    txtNewPurchasePrice.Text = Empty
    txtNewPurchasePrice.Text = Empty
    GridPPrice.Enabled = True
End Sub

Private Sub SavePurchasePrice()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select* From tblPurchasePrice"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ItemID = Val(dtcItem.BoundText)
            !PPrice = Val(txtNewPurchasePrice.Text)
            !setdate = Date
            !SetTime = Now
            !StaffID = UserID
            .Update
        If .State = 1 Then .Close
        temSql = "Select* From tblCurrentPurchasePrice Where ItemID = " & Val(dtcItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount < 1 Then
                .AddNew
                !ItemID = Val(dtcItem.BoundText)
                !PPrice = Val(txtNewPurchasePrice.Text)
                !setdate = Date
                !SetTime = Now
                !StaffID = UserID
                .Update
            ElseIf .RecordCount = 1 Then
                !PPrice = Val(txtNewPurchasePrice.Text)
                !setdate = Date
                !SetTime = Now
                !StaffID = UserID
                .Update
            Else
                If .State = 1 Then .Close
                temSql = "Delete From tblCurrentPurchasePrice Where ItemID = " & Val(dtcItem.BoundText)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .State = 1 Then .Close
                temSql = "Select* From tblCurrentPurchasePrice Where ItemID = " & Val(dtcItem.BoundText)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                .AddNew
                !ItemID = Val(dtcItem.BoundText)
                !PPrice = Val(txtNewPurchasePrice.Text)
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
    NewItem.ID = dtcItem.BoundText
    lblP1.Caption = "per " & NewItem.IUnit
    lblP2.Caption = lblP1.Caption
    Call FillPurchasePrice(Val(dtcItem.BoundText))
    bttnUpdate.Enabled = True
End Sub

Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtNewPurchasePrice.SetFocus
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

Private Sub FillPurchasePrice(ByVal ItemID As Long)
Dim i As Long
With GridPPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Purchase Price"
    
    .ColWidth(0) = (.Width - 150) / 2
    .ColWidth(1) = (.Width - 150) / 2
    
End With

With rsTem
    If .State = 1 Then .Close
    temSql = "SELECT tblPurchasePrice.SetDate, tblPurchasePrice.PPrice FROM tblPurchasePrice WHERE tblPurchasePrice.ItemID=" & ItemID & " ORDER BY tblPurchasePrice.SetDate DESC , tblPurchasePrice.SetTime DESC "
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        Do While .EOF = False
            With GridPPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTem!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTem!PPrice)
            End With
        .MoveNext
        Loop
    End If
End With

With rsTem
    If .State = 1 Then .Close
    temSql = "SELECT tblCurrentPurchasePrice.PPrice FROM tblCurrentPurchasePrice WHERE ItemID=" & ItemID & " Order by SetDate desc, SetTime Desc"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        txtCPurchasePrice.Text = Format(!PPrice)
    Else
        txtCPurchasePrice.Text = Empty
    End If
End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsViewCatogery.State = 1 Then rsViewCatogery.Close
    If rsViewItem.State = 1 Then rsViewItem.Close
    If rsViewCatItem.State = 1 Then rsViewCatItem.Close
    If rsTem.State = 1 Then rsTem.Close
End Sub

Private Sub GridPPrice_Click()
    Frame1.Enabled = False
    bttnCancel.Visible = True
    With GridPPrice
    .Col = 1
    txtCPurchasePrice.Text = Format(.Text, "0.00")
    
    .Col = 0
    .ColSel = .Cols - 1
    .Enabled = False
    End With
End Sub

Private Sub txtNewPurchasePrice_Change()
    If IsNumeric(txtNewPurchasePrice.Text) = True Then
    bttnUpdate.Enabled = True
    Else
    bttnUpdate.Enabled = False
    End If
End Sub

Private Sub txtNewPurchasePrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnUpdate_Click
    End If
End Sub

Private Sub txtNewPurchasePrice_LostFocus()
    With GridPPrice
        .Col = 1
        .Text = txtNewPurchasePrice.Text
        .Text = Format(.Text, "0.00")
    End With
End Sub

