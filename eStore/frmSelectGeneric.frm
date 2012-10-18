VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSelectGeneric 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select By Generic Name"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
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
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11760
   Begin VB.ListBox lstItemID 
      Height          =   2700
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstItem 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   9375
   End
   Begin btButtonEx.ButtonEx btnOK 
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   5400
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
   Begin MSDataListLib.DataCombo cmbGeneric 
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   714
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   405
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   714
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   405
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   714
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   8438015
      Text            =   ""
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
   Begin VB.Label Label4 
      Caption         =   "Trade Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblDristributor 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   4800
      Width           =   9375
   End
   Begin VB.Label Label3 
      Caption         =   "&Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&Generic Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmSelectGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsItem As New ADODB.Recordset
    Dim rsGeneric As New ADODB.Recordset
    Dim rsCategory As New ADODB.Recordset
    Dim rsDI As New ADODB.Recordset
    Dim TemDI As Long
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset
    Dim MyItem As New Item
    
Private Sub btnOK_Click()
    frmHospitalSale.dtcCatogery.BoundText = cmbCategory.BoundText
    frmHospitalSale.dtcItem.BoundText = Val(lstItemID.Text)
    Unload Me
End Sub


Private Sub cmbCategory_Change()
    Call ListItems
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstItem.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub ListItems()
    With rsItem
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
                temSql = "SELECT tblItem.Display , tblItem.ItemID , Sum(tblBatchStock.Stock) AS SumOfStock, (tblCurrentSalePrice.SPrice) AS LastOfSPrice " & _
                            "FROM ((tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) RIGHT JOIN tblItem ON tblBatch.ItemID = tblItem.ItemID) LEFT JOIN tblCurrentSalePrice ON tblItem.ItemID = tblCurrentSalePrice.ItemID " & _
                            "Where (((tblItem.GenericNameID) = " & Val(cmbGeneric.BoundText) & " )  AND ((tblItem.ItemCategoryID)=" & Val(cmbCategory.BoundText) & ") ) " & _
                            "GROUP BY tblItem.Display , tblItem.ItemID, tblCurrentSalePrice.SPrice " & _
                            "ORDER BY tblItem.Display"
        Else
                temSql = "SELECT tblItem.Display , tblItem.ItemID , Sum(tblBatchStock.Stock) AS SumOfStock, (tblCurrentSalePrice.SPrice) AS LastOfSPrice " & _
                            "FROM ((tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) RIGHT JOIN tblItem ON tblBatch.ItemID = tblItem.ItemID) LEFT JOIN tblCurrentSalePrice ON tblItem.ItemID = tblCurrentSalePrice.ItemID " & _
                            "Where (((tblItem.GenericNameID) = " & Val(cmbGeneric.BoundText) & " )) " & _
                            "GROUP BY tblItem.Display , tblItem.ItemID, tblCurrentSalePrice.SPrice " & _
                            "ORDER BY tblItem.Display"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        lstItem.Clear
        lstItemID.Clear
        While .EOF = False
            lstItem.AddItem (Left(!Display & Space(40), 40)) & vbTab & Right(Space(10) & !SumOfStock, 10) & vbTab & Right(Space(10) & Format(!LastOfSPrice, "0.00"), 10)
            lstItemID.AddItem !ItemID
            .MoveNext
        Wend
    End With
End Sub

Private Sub cmbGeneric_Change()
    Call ListItems
End Sub

Private Sub cmbGeneric_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbGeneric.Text = Empty
    End If
End Sub

Private Sub cmbItem_Change()
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    MyItem.ID = cmbItem.BoundText
    cmbGeneric.BoundText = MyItem.GenericID
End Sub

Private Sub Form_Load()
    Call FillCombos
    cmbCategory.BoundText = frmHospitalSale.dtcCatogery.BoundText
End Sub

Private Sub FillCombos()
    With rsViewItem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem order by Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbItem
        Set .RowSource = rsViewItem
        .ListField = "Display"
        .BoundColumn = "ItemID"
    End With
    With rsGeneric
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblGenericName order by GenericName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbGeneric
        Set .RowSource = rsGeneric
        .ListField = "GenericName"
        .BoundColumn = "GenericNameID"
    End With
    With rsCategory
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItemCategory order by categoryCode"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
End Sub


Private Sub lstItem_Click()
    lstItemID.ListIndex = lstItem.ListIndex
    DistributorDetails (Val(lstItemID.Text))
End Sub

Private Sub lstItem_DblClick()
    lstItemID.ListIndex = lstItem.ListIndex
    btnOK_Click
End Sub

Private Sub lstItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnOK_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbGeneric.SetFocus
    End If
End Sub
Private Sub DistributorDetails(ItemID As Long)
    With rsDI
        If .State = 1 Then .Close
        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & (Val(lstItemID.Text)) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        TemDI = !DistributorID
        End If
        .Close
    End With
    With rsTemDistributor
        If .State = 1 Then .Close
        temSql = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & TemDI & ""
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!DistributorName) Then lblDristributor.Caption = !DistributorName
        If .State = 1 Then .Close
    End With
End Sub
