VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Suppliers"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
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
   ScaleHeight     =   6795
   ScaleWidth      =   11100
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
   Begin VB.Frame framAdd 
      Height          =   6015
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   6495
      Begin MSDataListLib.DataCombo dtcDistributor 
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataList dtlDistributorName 
         Height          =   2940
         Left            =   1920
         TabIndex        =   7
         Top             =   2760
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5186
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx bttnAddDistributor 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   2280
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
      Begin btButtonEx.ButtonEx bttnDelete 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   2280
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
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Caption         =   "Item Name"
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Distributer Name"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame FrameSearch 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin MSDataListLib.DataCombo dtcItem 
         Height          =   5460
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9631
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
   End
End
Attribute VB_Name = "frmItemSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewItem As New ADODB.Recordset
    Dim rsViewDitributor  As New ADODB.Recordset
    Dim rsViewItemDitributor  As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String

Private Sub bttnAddDistributor_Click()
    Dim tr As Integer
    If IsNumeric(dtcItem.BoundText) = False Then
        tr = MsgBox("Select Item Name", vbCritical + vbOKOnly, "Error")
        dtcItem.SetFocus
        Exit Sub
    End If
    If IsNumeric(dtcDistributor.BoundText) = False Then
        tr = MsgBox("Select Distributor Name", vbCritical + vbOKOnly, "Error")
        dtcDistributor.SetFocus
        Exit Sub
    End If
    With rsTem
        If .State = 1 Then .Close
        .Open "Select* From tblItemDistributor", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ItemID = Val(dtcItem.BoundText)
        !DistributorID = Val(dtcDistributor.BoundText)
        .Update
        If .State = 1 Then .Close
        dtcDistributor.BoundText = Empty
        Call FillItemDistributors
    End With
End Sub

Private Sub ClearValues()
    lblItemName.Caption = Empty
    dtcDistributor.Text = Empty
    Set dtlDistributorName.RowSource = Nothing
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub
    
Private Sub FillItemCombo()
    With rsViewItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItem order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcItem.RowSource = rsViewItem
        dtcItem.BoundColumn = "ItemID"
        dtcItem.ListField = "display"
    End With
End Sub

Private Sub bttnDelete_Click()
    If IsNumeric(dtlDistributorName.BoundText) = False Then: MsgBox "Select Delete Distributor Name": Exit Sub
    With rsTem
        If .State = 1 Then .Close
        .Open "Delete From tblItemDistributor Where ItemDistributorID = " & Val(dtlDistributorName.BoundText) & " ", cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        MsgBox "Distributor Deleted"
        bttnDelete.Enabled = False
        Call FillItemDistributors
    End With
End Sub

Private Sub FillItemDistributors()
    If IsNumeric(dtcItem.BoundText) = False Then Exit Sub
    lblItemName.Caption = dtcItem.Text
    With rsViewItemDitributor
        If .State = 1 Then .Close
        temSql = "SELECT tblItemDistributor.*, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN tblItemDistributor ON tblDistrubutor.DistributorID = tblItemDistributor.DistributorID Where ItemID = " & dtcItem.BoundText & ""
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtlDistributorName.RowSource = rsViewItemDitributor
        dtlDistributorName.BoundColumn = "ItemDistributorID"
        dtlDistributorName.ListField = "DistributorName"
    End With
    bttnDelete.Enabled = False
End Sub

Private Sub dtcItem_Click(Area As Integer)
    Call FillItemDistributors
End Sub

Private Sub dtlDistributorName_Click()
    bttnDelete.Enabled = True
    bttnDelete.Visible = True
End Sub

Private Sub Form_Load()
    Call FillItemCombo
    Call fillDistributorCombo
End Sub

Private Sub fillDistributorCombo()
    With rsViewDitributor
        If .State = 1 Then .Close
        temSql = "SELECT * from tblDistrubutor order by DistributorName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcDistributor.RowSource = rsViewDitributor
        dtcDistributor.ListField = "DistributorName"
        dtcDistributor.BoundColumn = "DistributorID"
    End With
End Sub

