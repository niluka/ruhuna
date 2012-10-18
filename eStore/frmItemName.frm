VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmItemName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Name"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
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
   ScaleHeight     =   8910
   ScaleWidth      =   11235
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   7440
      Width           =   4455
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Add"
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
         TabIndex        =   40
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Edit"
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
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   7440
      Width           =   6495
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Save"
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
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Cancel"
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
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin TabDlg.SSTab SSTab1 
         Height          =   6735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   11880
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Product Details"
         TabPicture(0)   =   "frmItemName.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Supply Details"
         TabPicture(1)   =   "frmItemName.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame6 
            Height          =   6015
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   6015
            Begin VB.TextBox txtReorderqty 
               Height          =   375
               Left            =   2160
               TabIndex        =   42
               Top             =   840
               Width           =   3615
            End
            Begin VB.TextBox txtReorderLeval 
               Height          =   375
               Left            =   2160
               TabIndex        =   41
               Top             =   360
               Width           =   3615
            End
            Begin btButtonEx.ButtonEx bttnDelete 
               Height          =   255
               Left            =   4200
               TabIndex        =   37
               Top             =   5640
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               Appearance      =   3
               Enabled         =   0   'False
               Caption         =   "Delete"
               Enabled         =   0   'False
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
            Begin MSDataListLib.DataList dtlDistributors 
               Height          =   1980
               Left            =   2160
               TabIndex        =   36
               Top             =   3120
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   3493
               _Version        =   393216
            End
            Begin btButtonEx.ButtonEx bttnAddDistributor 
               Height          =   255
               Left            =   2160
               TabIndex        =   32
               Top             =   2760
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   450
               Appearance      =   3
               Enabled         =   0   'False
               Caption         =   "Add Distributor"
               Enabled         =   0   'False
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
            Begin MSDataListLib.DataCombo dtcMnufacture 
               Height          =   360
               Left            =   2160
               TabIndex        =   28
               Top             =   1320
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcImporter 
               Height          =   360
               Left            =   2160
               TabIndex        =   29
               Top             =   1800
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcDistributor 
               Height          =   360
               Left            =   2160
               TabIndex        =   30
               Top             =   2280
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label Label7 
               Caption         =   "Re -Order Qty"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label Label6 
               Caption         =   "Re- Order Leval"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label14 
               Caption         =   "Selected Distributors"
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   3120
               Width           =   1935
            End
            Begin VB.Label Label13 
               Caption         =   "Select Distrubutor "
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2280
               Width           =   2055
            End
            Begin VB.Label Label12 
               Caption         =   "Importor Name"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1800
               Width           =   2055
            End
            Begin VB.Label Label11 
               Caption         =   "Mnufacture Name"
               Height          =   375
               Left            =   120
               TabIndex        =   26
               Top             =   1320
               Width           =   2055
            End
         End
         Begin VB.Frame Frame5 
            Height          =   6015
            Left            =   -74760
            TabIndex        =   9
            Top             =   480
            Width           =   5895
            Begin VB.TextBox txtComments 
               Height          =   615
               Left            =   2040
               MultiLine       =   -1  'True
               TabIndex        =   52
               Top             =   5160
               Width           =   3615
            End
            Begin VB.TextBox txtPurchaseUnit 
               Height          =   360
               Left            =   2040
               TabIndex        =   50
               Top             =   3720
               Width           =   1455
            End
            Begin VB.TextBox txtIsueUnit 
               Height          =   375
               Left            =   2040
               TabIndex        =   48
               Top             =   4200
               Width           =   1455
            End
            Begin VB.TextBox txtStrength 
               Height          =   375
               Left            =   2040
               TabIndex        =   46
               Top             =   3240
               Width           =   1455
            End
            Begin VB.TextBox txtItemCode 
               Height          =   375
               Left            =   2040
               TabIndex        =   34
               Top             =   840
               Width           =   3615
            End
            Begin VB.TextBox txtItemname 
               Height          =   375
               Left            =   2040
               TabIndex        =   16
               Top             =   1320
               Width           =   3615
            End
            Begin VB.TextBox txtDisplayName 
               Height          =   375
               Left            =   2040
               TabIndex        =   11
               Top             =   360
               Width           =   3615
            End
            Begin VB.TextBox txtUnitRatio 
               Height          =   375
               Left            =   2040
               TabIndex        =   10
               Top             =   4680
               Width           =   3615
            End
            Begin MSDataListLib.DataCombo dtcPurchaseUnit 
               Height          =   360
               Left            =   3600
               TabIndex        =   12
               Top             =   4200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcTradeName 
               Height          =   360
               Left            =   2040
               TabIndex        =   13
               Top             =   2280
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcGenericName 
               Height          =   360
               Left            =   2040
               TabIndex        =   14
               Top             =   2760
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcCatogeryName 
               Height          =   360
               Left            =   2040
               TabIndex        =   15
               Top             =   1800
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcStrengthUnit 
               Height          =   360
               Left            =   3600
               TabIndex        =   47
               Top             =   3240
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcIssueUnit 
               Height          =   360
               Left            =   3600
               TabIndex        =   49
               Top             =   3720
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label Label17 
               Caption         =   "Comments"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   5160
               Width           =   1815
            End
            Begin VB.Label Label16 
               Caption         =   "Strength"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   3240
               Width           =   1935
            End
            Begin VB.Label Label15 
               Caption         =   "Item Code"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "Pack Name"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label2 
               Caption         =   "Item Name"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label3 
               Caption         =   "Catogery Name"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label4 
               Caption         =   "Generic Name"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Label Label5 
               Caption         =   "Trade Name"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   2280
               Width           =   1695
            End
            Begin VB.Label Label8 
               Caption         =   "Purchase Unit"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   4200
               Width           =   1935
            End
            Begin VB.Label Label9 
               Caption         =   "Issue Unit"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   3720
               Width           =   2175
            End
            Begin VB.Label Label10 
               Caption         =   "Purchase/Issue Ratio"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   4680
               Width           =   1935
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSDataListLib.DataCombo dtcItemName 
         Height          =   6900
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   12171
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
Attribute VB_Name = "frmItemName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsviewItem As New ADODB.Recordset
Dim rsViewCatogery As New ADODB.Recordset
Dim rsViewTrade As New ADODB.Recordset
Dim rsViewGenaric As New ADODB.Recordset
Dim rsViewUnitName As New ADODB.Recordset
Dim rsViewDistributor As New ADODB.Recordset
Dim rsViewManufcture As New ADODB.Recordset
Dim rsViewImporter As New ADODB.Recordset
Dim rsViewItemDistributor As New ADODB.Recordset
Dim rsItem As New ADODB.Recordset
Dim rsDistributor As New ADODB.Recordset
Dim TemItemID As Long
Dim CheckEmptys As Boolean
Dim rsTem As New ADODB.Recordset


Private Sub bttnAdd_Click()
Call AfterAdd
End Sub

Private Sub bttnAddDistributor_Click()
If txtItemname.Text = Empty Then A = MsgBox("Enter Item Name", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 0: txtItemname.SetFocus: Exit Sub
If dtcDistributor.BoundText = Empty Then A = MsgBox("Select Distributor Name", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: dtcDistributor.SetFocus: Exit Sub

With rsDistributor
    If .State = 1 Then .Close
    .Open "Select* From tblItemDistributor", cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !ItemID = dtcItemName.BoundText
    !DistributorID = dtcDistributor.BoundText
    .Update
    bttnAddDistributor.Enabled = False
    If .State = 1 Then .Close
    Call FillItemDistributors

End With

End Sub

Private Sub FillItemDistributors()
With rsViewItemDistributor
    If .State = 1 Then .Close
    .Open "SELECT tblItemDistributor.*, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN tblItemDistributor ON tblDistrubutor.DistributorID = tblItemDistributor.DistributorID Where ItemID = " & dtcItemName.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtlDistributors.RowSource = rsViewItemDistributor
    dtlDistributors.BoundColumn = "ItemDistributorID"
    dtlDistributors.ListField = "DistributorName"
End With
End Sub

Private Sub bttnCancel_Click()
Call AfterAddEdit
End Sub

Private Sub bttnChange_Click()
Call CheckEmptyValues
If CheckEmptys = False Then Exit Sub
Call EditItem
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnDelete_Click()
If IsNumeric(dtlDistributors.BoundText) = False Then: MsgBox "Select Distributor Name": Exit Sub
With rsDistributor
    If .State = 1 Then .Close
    .Open "Delete* From tblItemDistributor Where ItemDistributorID = " & dtlDistributors.BoundText & "", cnnStores, adOpenStatic, adLockOptimistic
    If .State = 1 Then .Close
    MsgBox "Distributor Deleted"
    bttnDelete.Enabled = False
    Call FillItemDistributors
    
End With
End Sub

Private Sub bttnEdit_Click()
Call AfterEdit
End Sub

Private Sub bttnSave_Click()
Call CheckEmptyValues
If CheckEmptys = False Then Exit Sub
Call SaveItem
End Sub

Private Sub SaveItem()
With rsItem
    If .State = 1 Then .Close
    .Open "Select* From tblItem", cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !ItemDisplayName = txtDisplayName.Text
    !Item = txtItemname.Text
    !ItemCode = txtItemCode.Text
    !ItemCatogeryID = Val(dtcCatogeryName.BoundText)
    !GenericNameID = Val(dtcGenericName.BoundText)
    !TradeNameID = Val(dtcTradeName.BoundText)
    !ROL = Val(txtReorderLeval.Text)
    !ROQ = Val(txtReorderqty.Text)
    !Strength = Val(txtStrength.Text)
    !StrengthUnitID = Val(dtcStrengthUnit.BoundText)
    !PurchaseUnit = Val(txtPurchaseUnit.Text)
    !PurchaseUnitID = Val(dtcPurchaseUnit.BoundText)
    !IssueUnit = Val(txtIsueUnit.Text)
    !IssueUnitID = Val(dtcIssueUnit.BoundText)
    !PurchaseIssueUnitRatio = Val(txtUnitRatio.Text)
    !ManufacturerID = Val(dtcMnufacture.BoundText)
    !ImporterID = Val(dtcImporter.BoundText)
    !Comments = txtComments.Text
    .Update
    
    Call ClearVales
    Call AfterAddEdit
    Call FillItemCombo
    If .State = 1 Then .Close
    
End With

End Sub

Private Sub EditItem()

With rsItem
    If .State = 1 Then .Close
    .Open "Select* From tblItem Where ItemID = " & TemItemID & "", cnnStores, adOpenStatic, adLockOptimistic
    
    !ItemDisplayName = txtDisplayName.Text
    !Item = txtItemname.Text
    !ItemCode = txtItemCode.Text
    !ItemCatogeryID = Val(dtcCatogeryName.BoundText)
    !GenericNameID = Val(dtcGenericName.BoundText)
    !TradeNameID = Val(dtcTradeName.BoundText)
    !ROL = Val(txtReorderLeval.Text)
    !ROQ = Val(txtReorderqty.Text)
    !Strength = Val(txtStrength.Text)
    !StrengthUnitID = Val(dtcStrengthUnit.BoundText)
    !PurchaseUnit = Val(txtPurchaseUnit.Text)
    !PurchaseUnitID = Val(dtcPurchaseUnit.BoundText)
    !IssueUnit = Val(txtIsueUnit.Text)
    !IssueUnitID = Val(dtcIssueUnit.BoundText)
    !PurchaseIssueUnitRatio = Val(txtUnitRatio.Text)
    !ManufacturerID = Val(dtcMnufacture.BoundText)
    !ImporterID = Val(dtcImporter.BoundText)
    !Comments = txtComments.Text
    .Update

    Call ClearVales
    Call AfterAddEdit
    Call FillItemCombo
    If .State = 1 Then .Close
    
End With

End Sub
Private Sub DisplaySelctedItem()
With rsItem
    If .State = 1 Then .Close
    .Open "Select* From TblItem Where ItemId = " & dtcItemName.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
    
    If .RecordCount = 0 Then Exit Sub
    Call ClearVales
    
    If Not (!ItemDisplayName) = "" Then txtDisplayName.Text = !ItemDisplayName
    If Not (!Item) = "" Then txtItemname.Text = !Item
    If Not (!ItemCode) = "" Then txtItemCode.Text = !ItemCode
    If Not (!ItemCatogeryID) = "" Then dtcCatogeryName.BoundText = !ItemCatogeryID
    If Not (!GenericNameID) = "" Then dtcGenericName.BoundText = !GenericNameID
    If Not (!TradeNameID) = "" Then dtcTradeName.BoundText = !TradeNameID
    If Not (!ROL) = "" Then txtReorderLeval.Text = !ROL
    If Not (!ROQ) = "" Then txtReorderqty = !ROQ
    If Not (!Strength) = "" Then txtStrength.Text = !Strength
    If Not (!StrengthUnitID) = "" Then dtcStrengthUnit.BoundText = !StrengthUnitID
    If Not (!PurchaseUnit) = "" Then txtPurchaseUnit.Text = !PurchaseUnit
    If Not (!PurchaseUnitID) = "" Then dtcPurchaseUnit.BoundText = !PurchaseUnitID
    If Not (!IssueUnit) = "" Then txtIsueUnit.Text = !IssueUnit
    If Not (!IssueUnitID) = "" Then dtcIssueUnit.BoundText = !IssueUnitID
    If Not (!PurchaseIssueUnitRatio) = "" Then txtUnitRatio.Text = !PurchaseIssueUnitRatio
    If Not (!ManufacturerID) = "" Then dtcMnufacture.BoundText = !ManufacturerID
    If Not (!ImporterID) = "" Then dtcImporter.BoundText = !ImporterID
    If Not (!Comments) = "" Then txtComments.Text = !Comments
    
    TemItemID = !ItemID
    
    If .State = 1 Then .Close
End With

End Sub

Private Sub FillItemCombo()
With rsviewItem
    If .State = 1 Then .Close
    .Open "Select ItemID,Item From TblItem", cnnStores, adOpenStatic, adLockReadOnly
    Set dtcItemName.RowSource = rsviewItem
    dtcItemName.BoundColumn = "ItemID"
    dtcItemName.ListField = "Item"

End With
End Sub

Private Sub FillCatogeryCombo()
With rsViewCatogery
    If .State = 1 Then .Close
    .Open "Select tblItemCatogery.* From tblItemCatogery Order By ItemCatogery", cnnStores, adOpenStatic, adLockReadOnly

    If .RecordCount = 0 Then Exit Sub
    Set dtcCatogeryName.RowSource = rsViewCatogery
    dtcCatogeryName.ListField = "ItemCatogery"
    dtcCatogeryName.BoundColumn = "ItemCategoryID"


End With
End Sub

Private Sub FillGrnaricCombo()
Dim TemId As Long
Dim TemDql As String

With rsViewGenaric
    If .State = 1 Then .Close
    TemSql = "SELECT * from tblgenericname order by genericname"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount < 1 Then Exit Sub
End With
With rsTem
    If .State = 1 Then .Close
    TemSql = "SELECT * from tbltradename where tradenameid = " & dtcTradeName.BoundText
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount < 1 Then Exit Sub
    TemDql = !GenericNameID
End With
With dtcGenericName
    Set .RowSource = rsViewGenaric
    .ListField = "genericname"
    .BoundColumn = "GenericnameID"
    .BoundText = TemDql
    .Locked = True
End With

'dtcGenericName.BoundColumn = Empty
'dtcGenericName.ListField = Empty
'
'With rsViewGenaric
'    If .State = 1 Then .Close
'    .Open "SELECT tblTradeName.*, tblGenericName.GenericName FROM tblGenericName RIGHT JOIN tblTradeName ON tblGenericName.GenericNameID = tblTradeName.GenericNameID Where (TradeNameID = " & Val(dtcTradeName.BoundText) & ") ORDER BY tblGenericName.GenericName", cnnStores, adOpenStatic, adLockReadOnly
'    If .RecordCount = 0 Then Exit Sub
'    TemId = !GenericNameID
'    Set dtcGenericName.RowSource = rsViewGenaric
'    dtcGenericName.BoundColumn = "GenericNameID"
'    dtcGenericName.ListField = "GenericName"
'    dtcGenericName.BoundText = TemId
 
'End With

End Sub

Private Sub FillTradeName()
With rsViewTrade
    If .State = 1 Then .Close
    .Open "Select* From tblTradeName Order By TradeName", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcTradeName.RowSource = rsViewTrade
    dtcTradeName.BoundColumn = "TradeNameID"
    dtcTradeName.ListField = "TradeName"
End With
End Sub

Private Sub FillUnitName()
With rsViewUnitName
    If .State = 1 Then .Close
    .Open "Select* From tblUnit Order By Unit", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    
    Set dtcIssueUnit.RowSource = rsViewUnitName
    dtcIssueUnit.BoundColumn = "UnitID"
    dtcIssueUnit.ListField = "Unit"
    
    Set dtcPurchaseUnit.RowSource = rsViewUnitName
    dtcPurchaseUnit.BoundColumn = "UnitID"
    dtcPurchaseUnit.ListField = "Unit"
    
    Set dtcStrengthUnit.RowSource = rsViewUnitName
    dtcStrengthUnit.BoundColumn = "UnitID"
    dtcStrengthUnit.ListField = "Unit"
        
End With
End Sub

Private Sub FillDistributors()
With rsViewDistributor
    If .State = 1 Then .Close
    .Open "Select * From tblDistrubutor Order By DistributorName", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcDistributor.RowSource = rsViewDistributor
    dtcDistributor.BoundColumn = "DistributorID"
    dtcDistributor.ListField = "DistributorName"
End With
End Sub

Private Sub FillMnufacture()
With rsViewManufcture
    If .State = 1 Then .Close
    .Open "Select * From tblManufacturer Order By ManufacturerName", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcMnufacture.RowSource = rsViewManufcture
    dtcMnufacture.BoundColumn = "ManufacturerID"
    dtcMnufacture.ListField = "ManufacturerName"
End With
End Sub

Private Sub FillImporter()
With rsViewImporter
    If .State = 1 Then .Close
    .Open "Select * From tblImporter Order By ImporterName", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcImporter.RowSource = rsViewImporter
    dtcImporter.BoundColumn = "ImporterID"
    dtcImporter.ListField = "ImporterName"
End With
End Sub

Private Sub dtcDistributor_Click(Area As Integer)
If IsNumeric(dtcDistributor.BoundText) = False Then Exit Sub
bttnAddDistributor.Enabled = True
End Sub

Private Sub dtcIssueUnit_LostFocus()
Call CreateName
End Sub

Private Sub dtcItemName_Click(Area As Integer)
If IsNumeric(dtcItemName.BoundText) = False Then Exit Sub
Call DisplaySelctedItem
Call FillItemDistributors
End Sub

Private Sub dtcUnitName_LostFocus()
If IsNumeric(dtcUnitName.BoundText) = False Then Exit Sub
Call CreateName
End Sub

Private Sub CreateName()
txtDisplayName.Text = ""
txtItemname.Text = ""
If Not dtcTradeName.Text = dtcGenericName.Text Then
    txtItemname = dtcTradeName.Text & " " & txtStrength.Text & "" & dtcStrengthUnit.Text & " " & dtcIssueUnit.Text
    txtDisplayName.Text = dtcTradeName.Text & " " & txtStrength.Text & "" & dtcStrengthUnit.Text & " " & dtcIssueUnit.Text & " " & txtUnitRatio.Text & " x " & dtcIssueUnit.Text
Else
    txtItemname = dtcTradeName.Text & " " & txtStrength.Text & "" & dtcStrengthUnit.Text & " " & dtcIssueUnit.Text & " (" & dtcGenericName.Text & " )"
    txtDisplayName.Text = dtcTradeName.Text & " " & txtStrength.Text & "" & dtcStrengthUnit.Text & " (" & dtcGenericName.Text & ")" & " " & dtcIssueUnit.Text & " " & txtUnitRatio.Text & " x " & dtcIssueUnit.Text
End If
End Sub



Private Sub dtcStrengthUnit_LostFocus()
Call CreateName
End Sub

Private Sub dtcTradeName_Change()
If IsNumeric(dtcTradeName.BoundText) = False Then Exit Sub
Call FillGrnaricCombo
End Sub

Private Sub dtcTradeName_Click(Area As Integer)
If IsNumeric(dtcTradeName.BoundText) = False Then Exit Sub
Call FillGrnaricCombo
End Sub

Private Sub dtcTradeName_LostFocus()
Call CreateName
End Sub

Private Sub dtlDistributors_Click()
If IsNumeric(dtlDistributors.BoundText) = False Then Exit Sub
bttnDelete.Enabled = True
End Sub

Private Sub Form_Load()
Call FillItemCombo
Call FillCatogeryCombo
'Call FillGrnaricCombo
Call FillTradeName
Call FillUnitName
Call FillMnufacture
Call FillImporter
Call AfterAddEdit
Call FillDistributors
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsviewItem.State = 1 Then rsviewItem.Close: Set rsviewItem = Nothing
If rsViewCatogery.State = 1 Then rsViewCatogery.Close: Set rsViewCatogery = Nothing
If rsViewGenaric.State = 1 Then rsViewGenaric.Close: Set rsViewGenaric = Nothing
If rsViewTrade.State = 1 Then rsViewTrade.Close: Set rsViewTrade = Nothing
If rsItem.State = 1 Then rsItem.Close: Set rsItem = Nothing
If rsViewUnitName.State = 1 Then rsViewUnitName.Close: Set rsViewUnitName = Nothing
If rsViewDistributor.State = 1 Then rsViewDistributor.Close: Set rsViewDistributor = Nothing
If rsViewManufcture.State = 1 Then rsViewManufcture.Close: Set rsViewManufcture = Nothing
If rsViewImporter.State = 1 Then rsViewImporter.Close: Set rsViewImporter = Nothing
End Sub

Private Sub AfterAdd()
bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnSave.Visible = True
bttnChange.Visible = False
bttnCancel.Visible = True
Frame2.Enabled = True
Frame1.Enabled = False
Call ClearVales
End Sub

Private Sub AfterEdit()
bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnSave.Visible = False
bttnChange.Visible = True
bttnCancel.Visible = True
Frame2.Enabled = True
Frame1.Enabled = False
End Sub

Private Sub AfterAddEdit()
bttnAdd.Enabled = True
bttnEdit.Enabled = True
bttnSave.Visible = False
bttnChange.Visible = False
bttnCancel.Visible = False
Frame2.Enabled = False
Frame1.Enabled = True
bttnAddDistributor.Enabled = False
bttnDelete.Enabled = False
SSTab1.Tab = 0
End Sub

Private Sub ClearVales()
txtDisplayName.Text = Empty
txtItemname.Text = Empty
txtItemCode.Text = Empty
dtcCatogeryName.BoundText = Empty
dtcGenericName.BoundText = Empty
dtcTradeName.BoundText = Empty
txtReorderLeval.Text = Empty
txtReorderqty.Text = Empty
dtcIssueUnit.BoundText = Empty
dtcPurchaseUnit.BoundText = Empty
dtcStrengthUnit.BoundText = Empty
txtIsueUnit.Text = Empty
txtPurchaseUnit.Text = Empty
txtStrength.Text = Empty
txtUnitRatio.Text = Empty
dtcMnufacture.BoundText = Empty
dtcImporter.BoundText = Empty
dtcDistributor.BoundText = Empty
dtlDistributors.ListField = Empty
dtcGenericName.Locked = False
txtComments.Text = Empty
TemItemID = Empty
End Sub

Private Sub CheckEmptyValues()
Dim A
CheckEmptys = False
If txtDisplayName.Text = Empty Then A = MsgBox("Enter Display Name", vbCritical + vbOKOnly, "Error"): txtDisplayName.SetFocus: Exit Sub
If txtItemname.Text = Empty Then A = MsgBox("Enter Item Name", vbCritical + vbOKOnly, "Error"): txtItemname.SetFocus: Exit Sub
If txtItemCode.Text = Empty Then A = MsgBox("Enter Item Code Name", vbCritical + vbOKOnly, "Error"): txtItemCode.SetFocus: Exit Sub
If dtcCatogeryName.BoundText = Empty Then A = MsgBox("Select Catogery Name Name", vbCritical + vbOKOnly, "Error"): dtcCatogeryName.SetFocus: Exit Sub
If dtcTradeName.BoundText = Empty Then A = MsgBox("Select Trade Name", vbCritical + vbOKOnly, "Error"): dtcTradeName.SetFocus: Exit Sub
If dtcGenericName.BoundText = Empty Then A = MsgBox("Select Genaric Name", vbCritical + vbOKOnly, "Error"): dtcGenericName.SetFocus: Exit Sub
If txtReorderLeval.Text = Empty Then A = MsgBox("Enter Re-Order Leval", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: txtReorderLeval.SetFocus: Exit Sub
If txtReorderqty.Text = Empty Then A = MsgBox("Enter Re-Order Quantity", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: txtReorderqty.SetFocus: Exit Sub
If txtStrength.Text = Empty Then A = MsgBox("Enter Strenth Unit", vbCritical + vbOKOnly, "Error"): txtStrength.SetFocus: Exit Sub
If dtcStrengthUnit.BoundText = Empty Then A = MsgBox("Select Strenth Unit Name", vbCritical + vbOKOnly, "Error"): dtcStrengthUnit.SetFocus: Exit Sub
If txtIsueUnit.Text = Empty Then A = MsgBox("Enter Issue Unit", vbCritical + vbOKOnly, "Error"): txtIsueUnit.SetFocus: Exit Sub
If dtcIssueUnit.BoundText = Empty Then A = MsgBox("Select Issue Unit Name", vbCritical + vbOKOnly, "Error"): dtcIssueUnit.SetFocus: Exit Sub
If txtPurchaseUnit.Text = Empty Then A = MsgBox("Enter Purchase Issue Unit", vbCritical + vbOKOnly, "Error"): txtPurchaseUnit.SetFocus: Exit Sub
If dtcPurchaseUnit.BoundText = Empty Then A = MsgBox("Select Purchase Unit Name", vbCritical + vbOKOnly, "Error"): dtcPurchaseUnit.SetFocus: Exit Sub
If txtUnitRatio.Text = Empty Then A = MsgBox("Enter Unit Ratio", vbCritical + vbOKOnly, "Error"): txtUnitRatio.SetFocus: Exit Sub
If dtcMnufacture.BoundText = Empty Then A = MsgBox("Select Manufacture Name", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: dtcMnufacture.SetFocus: Exit Sub
If dtcImporter.BoundText = Empty Then A = MsgBox("Select Importer Name", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: dtcImporter.SetFocus: Exit Sub
If dtlDistributors.Appearance = Empty Then A = MsgBox("Select Distributor Name", vbCritical + vbOKOnly, "Error"): SSTab1.Tab = 1: dtcDistributor.SetFocus: Exit Sub
CheckEmptys = True

End Sub

Private Sub txtStrength_LostFocus()
Call CreateName

End Sub

Private Sub txtUnitRatio_LostFocus()
Call CreateName
End Sub
