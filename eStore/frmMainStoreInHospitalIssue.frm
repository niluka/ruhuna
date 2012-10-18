VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMainStoreInHospitalIssue 
   Caption         =   "In Hospital Issues"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   5505
   Begin MSFlexGridLib.MSFlexGrid GridIssues 
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.TextBox lblQuentityByIssueUnit 
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox lblQuentityByPurchaseUnit 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin MSDataListLib.DataCombo dtcRecevingStore 
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItemCatogery 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   5400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   8
      Top             =   5880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
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
   Begin btButtonEx.ButtonEx ButtonEx3 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Issue"
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
   Begin btButtonEx.ButtonEx ButtonEx4 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "E&xit"
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
   Begin VB.Label Label8 
      Caption         =   "Issued By :"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Checked By:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblIssueSize 
      Caption         =   "Label5"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPurchaseSize 
      Caption         =   "Label5"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Catogery"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Quentity"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "To Send to :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMainStoreInHospitalIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemIssueUnit As String
    Dim TemPurchaseUnit As String
    Dim TemIssuePurchaseRation As Double
    Dim TemQuentity As Double
    Dim TemSql As String
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCatogery As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemUnit As New ADODB.Recordset
    Dim rsStores As New ADODB.Recordset

Private Sub dtcItemCatogery_Click(Area As Integer)
    If IsNumeric(dtcItemCatogery.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
End Sub

Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    TemSql = "SELECT * from tblitem where itemcatogeryID = " & dtcItemCatogery.BoundText & " order by item"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly

End With
End Sub

Private Sub ListAllItems()

End Sub

Private Sub Form_Load()
    
    Call FillCombos
    
End Sub

Private Sub FillCombos()
With rsItem
    If .State = 1 Then .Close
    TemSql = "SELECT * from tblitem order by item"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Item"
    .BoundColumn = "ItemID"
End With
With rsItemCatogery
    If .State = 1 Then .Close
    TemSql = "SELECT * from tblitemcatogery order by itemcatogery"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItemCatogery
    Set .RowSource = rsItemCatogery
    .ListField = "ItemCatogery"
    .BoundColumn = "ItemCategoryID"
End With
With rsStaff
    If .State = 1 Then .Close
    TemSql = "SELECT * from tblstaff order by listedname"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcIssueStaff
    Set .RowSource = rsStaff
    .ListField = "ListedName"
    .BoundColumn = "StaffID"
End With
With dtcCheckedStaff
    Set .RowSource = rsStaff
    .ListField = "ListedName"
    .BoundColumn = "StaffID"
End With

' *************************************
' To Do

With rsStores
    If .State = 1 Then .Close
    TemSql = "SELECT tblMoveCatogery.MoveCatogery, tblMoveCatogery.IssueStore FROM tblMoveCatogery WHERE (((tblMoveCatogery.IssueStore)=1)) ORDER BY tblMoveCatogery.MoveCatogery"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcRecevingStore
    Set .RowSource = rsStore
    .ListField = "Store"
    .BoundColumn = "StoreID"
End With


' *************************************

End Sub
