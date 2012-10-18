VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransfers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfers"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
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
   ScaleHeight     =   7560
   ScaleWidth      =   14310
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   7320
      TabIndex        =   22
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmInHospitalIssue.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAMPP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtVMPP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAMP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtVMP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtVTM"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDisplay"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtIStore"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtRStore"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtRStoreID"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Issue Store"
      TabPicture(1)   =   "frmInHospitalIssue.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Receiving Store"
      TabPicture(2)   =   "frmInHospitalIssue.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Total Stock"
      TabPicture(3)   =   "frmInHospitalIssue.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSFlexGrid3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtRStoreID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtRStore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox txtIStore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox txtDisplay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtVTM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtVMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtAMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtVMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtAMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2880
         Width           =   4095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   37
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin VB.Label Label15 
         Caption         =   "Receiving Store :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label Label14 
         Caption         =   "Issuing Store :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Display Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Virtual Therapeutic Moiety:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Virtual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Actual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Virtual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label13 
         Caption         =   "Actual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   3255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridIssues 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2640
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
   Begin VB.TextBox txtQuentityByIssueUnit 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtQuentityByPurchaseUnit 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dtcTransaction 
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItemCatogery 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   5880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   8
      Top             =   6360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx bttnIssue 
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   6840
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
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   6840
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
   Begin MSDataListLib.DataCombo dtcBatch 
      Height          =   360
      Left            =   1560
      TabIndex        =   20
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "Batch"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Issued By :"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Checked By:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblPUnit 
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   2160
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
      Top             =   2160
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
Attribute VB_Name = "frmTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim TemIssueUnit As String
    Dim TemPurchaseUnit As String
    Dim TemIssueUnitID As Long
    Dim TemPurchaseUnitID As Long
    Dim TemQuentity As Double
    
    Dim TemSql As String
    Dim BorderMargin As Integer
    
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCatogery As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemUnit As New ADODB.Recordset
    Dim rsStores As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsBatch As New ADODB.Recordset
    Dim rsTemBatch As New ADODB.Recordset
    Dim NewItem As New Item


Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    Call AddToGrid
    Call ClearAddValues
    dtcItemCatogery.SetFocus
End Sub

Private Sub ClearAddValues()
    dtcItem.Text = Empty
    dtcItemCatogery.Text = Empty
    dtcBatch.Text = Empty
    txtQuentityByIssueUnit.Text = Empty
    txtQuentityByPurchaseUnit.Text = Empty
    lblIssueSize.Caption = Empty
    lblPurchaseSize.Caption = Empty
End Sub

Private Sub AddToGrid()
    With GridIssues
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .CellAlignment = 7
        .Text = .Rows - 1
        .Col = 1
        .CellAlignment = 1
        .Text = dtcItem.Text
        .Col = 2
        .Text = dtcItem.BoundText
        .Col = 3
        .CellAlignment = 1
        .Text = dtcBatch.Text
        .Col = 4
        .Text = dtcBatch.BoundText
        .Col = 5
        .CellAlignment = 7
        .Text = txtQuentityByIssueUnit.Text & " " & lblIssueSize.Caption
        .Col = 6
        .Text = txtQuentityByIssueUnit.Text
        .Col = 7
        .Text = lblIssueSize.Caption
        .Col = 8
        .CellAlignment = 7
        .Text = txtQuentityByPurchaseUnit.Text & " " & lblPurchaseSize.Caption
        .Col = 9
        .Text = txtQuentityByPurchaseUnit.Text
        .Col = 10
        .Text = lblPurchaseSize.Caption
        
        
'serial 0
'Item 1
'itemid 2
'batch 3
'batchid 4
'iquentitystring 5
'iquentity 6
'issueunit 7
'pQuentityString 8
'purchaseunit 9
'pQuentity 10
    End With
End Sub

Private Sub FormatGrid()
With GridIssues
    .Cols = 11
    .Rows = 1
    
    BorderMargin = 100
    
    .ColWidth(0) = 600
    
    .ColWidth(2) = 1
    .ColWidth(3) = 800
    .ColWidth(4) = 1
    .ColWidth(5) = 1600
    .ColWidth(6) = 1
    .ColWidth(7) = 1
    .ColWidth(8) = 1400
    .ColWidth(9) = 1
    .ColWidth(10) = 1
    
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7) + .ColWidth(8) + .ColWidth(9) + .ColWidth(10) + BorderMargin)
    
    
    .Row = 0
    Dim TemNum As Integer
    
    For TemNum = 0 To 10
        .Col = TemNum
        .CellAlignment = 4
        Select Case TemNum
            Case 0: .Text = "No."
            Case 1: .Text = "Item"
            Case 3: .Text = "Batch"
            Case 5: .Text = "Issue Qty"
            Case 8: .Text = "Pur. Qty"
        End Select
        
        
    Next TemNum
    
'serial 0
'Item 1
'itemid 2
'batch 3
'batchid 4
'iquentitystring 5
'iquentity 6
'issueunit 7
'pQuentityString 8
'purchaseunit 9
'pQuentity 10
'



End With
End Sub

Private Function CanIssue() As Boolean
    Dim Tr As Integer
    CanIssue = False
        If GridIssues.Rows <= 1 Then
            Tr = MsgBox("There are no items to issue", vbCritical, "No items")
            dtcItem.SetFocus
            Exit Function
        End If
        If IsNumeric(dtcTransaction.BoundText) = False Then
            Tr = MsgBox("Please select a tranfer name", vbCritical, "Tranfer to?")
            dtcTransaction.SetFocus
            Exit Sub
        End If
    CanIssue = True
End Function

Private Sub bttnDelete_Click()
    If GridIssues.Rows <= 2 Then
        Call FormatGrid
    Else
        GridIssues.RemoveItem (GridIssues.Row)
    End If
End Sub

Private Sub bttnExit_Click()
    Unload Me
End Sub

Private Sub bttnIssue_Click()
    If CanIssue = False Then Exit Sub
    Call TransferStocks
End Sub

Private Sub TransferStocks1()
    Dim RowItemID As Long
    Dim RowBatchID As Long
    Dim RowQuentity As Double
    Dim TemNum As Integer
    Dim NoError As Boolean
    Dim Tr As Integer
    With GridIssues
        .Visible = False
        NoError = True
        For TemNum = 1 To GridIssues.Rows - 1
            .Row = TemNum
            .Col = 2
            RowItemID = Val(.Text)
            .Col = 4
            RowBatchID = Val(.Text)
            .Col = 6
            RowQuentity Val(.Text)
            End If
            If TransferStocks(UserStoreID, Val(txtRStoreID.Text), RowItemID, RowBatchID, RowQuentity) = False Then
                NoError = False
            Else
                If .Rows = 2 And .Row = 1 Then
                    FormatGrid
                Else
                    .RemoveItem (.Row)
                End If
            End If
        Next
        If NoError = True Then
            Tr = MsgBox("Transfer of all the items were successful", vbInformation, "Successfully transferred")
        Else
            Tr = MsgBox("Some items are not transferred successfully", vbCritical, "Some are NOT transferred successfully")
        End If
        .Visible = True
    End With
End Sub

Private Function TransferStocks(ByVal IStoreIDValue As Long, ByVal RStoreIDValue As Long, ByVal ItemIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim Tr As Integer
    On Error GoTo eh
    TransferStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblBatch where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " And ItemID = " & ItemIDValue
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            Tr = MsgBox("There is such drug batch", vbCritical, "Error")
            .Close
            Exit Function
        End If
        If !stock < Quentity Then
            Tr = MsgBox("There are no enough stocks in you store to transfer to another store", vbCritical, "No Enough Stocks")
            .Close
            Exit Function
        End If
        !stock = !stock - Quentity
        .Update
        .Close
        
    With rsTemBatch
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblBatch where batchid = " & BatchIDValue & " AND StoreID = " & RStoreIDValue & " And ItemID = " & ItemIDValue
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            .AddNew
            !storeID = RStoreIDValue
            !batchid = BatchIDValue
            !itemid = ItemIDValue
        End If
        !stock = !stock + Quentity
        .Update
        .Close
    TransferStocks = True
    Exit Function

eh:
    .CancelUpdate
    .Close
    Tr = MsgBox("Could not deduct stocks from your store" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Function
    End With
End Function


Private Sub dtcItem_Click(Area As Integer)
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim TemdtcVal As Long
    Dim Tr As Integer
    NewItem.ID = Val(dtcItem.BoundText)
    lblIUnit.Caption = NewItem.IUnit
    lblPUnit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
    With rsBatch
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblbatch where ItemID =" & Val(dtcItem.BoundText) & " and  storeID = " & UserStoreID & " AND stock > 0 order by DOE "
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then
            Tr = MsgBox("There are no stocks", vbCritical, "No Stocks")
            .Close
            dtcItem.SetFocus
            With dtcBatch
                Set .RowSource = Nothing
                .ListField = Empty
                .BoundColumn = Empty
                .BoundText = Empty
            End With
            Exit Sub
        Else
            .MoveFirst
            TemdtcVal = !batchid
        End If
    End With
    With dtcBatch
        Set .RowSource = rsBatch
        .ListField = "Batch"
        .BoundColumn = "BatchID"
        .BoundText = TemdtcVal
    End With
    
End Sub

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
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Item"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub ListAllItems()
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
End Sub

Private Sub dtctransaction_Click(Area As Integer)
If IsNumeric(dtcTransaction.BoundText) = True Then
    With rsTemStore
        If .State = 1 Then .Close
        TemSql = "SELECT tblTransferCatogery.TransferCatogeryID, tblTransferCatogery.TransferCatogery, tblTransferCatogery.IStoreID, tblTransferCatogery.RStoreID, tblIStore.Store, tblRStore.Store, tblTransferCatogery.TransferCatogeryID FROM tblTransferCatogery, tblStore AS tblIStore, tblStore AS tblRStore WHERE (((tblTransferCatogery.TransferCatogeryID)=" & dtcTransaction.BoundText & "))"
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            txtRStore.Text = Empty
            txtRStoreID.Text = Empty
            .Close
            Exit Sub
        Else
            txtRStore.Text = A
            txtIStore.Text = A
        End If
Else
    txtRStore.Text = Empty
    txtRStoreID.Text = Empty
End If

End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
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
    TemSql = "SELECT tblTransferCatogery.TransferCatogeryID, tblTransferCatogery.TransferCatogery FROM tblTransferCatogery WHERE (((tblTransferCatogery.IStoreID)=" & UserStoreID & ")) ORDER BY tblTransferCatogery.TransferCatogery"
    .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcTransaction
    Set .RowSource = rsStores
    .ListField = "TransferCatogery"
    .BoundColumn = "TransferCatogeryID"
End With


' *************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CanExit = False Then Cancel = True: Exit Sub
End Sub

Private Function CanExit() As Boolean
    CanExit = False
    Dim Tr As Integer
    If GridIssues.Rows > 1 Then
        Tr = MsgBox("There are " & (GridIssues.Rows - 1) & " items listed to be issued, but the transction is not finalised. You have to delete all the items of confirn the issue by clicking the issue button ", vbCritical, "Transaction NOT finalised")
        Me.ZOrder 0
        bttnIssue.SetFocus
        Exit Function
    End If
    CanExit = True
End Function

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim Tr As Integer
    If Not IsNumeric(dtcTransaction.BoundText) Then
        Tr = MsgBox("You have not selected the transaction", vbCritical, "Transaction?")
        dtcTransaction.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcItem.BoundText) Then
        Tr = MsgBox("You have not selected the item to transfer", vbCritical, "Item?")
        dtcItem.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcBatch.BoundText) Then
        Tr = MsgBox("You have not selected a batch to transfer", vbCritical, "Batch?")
        dtcBatch.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtQuentityByIssueUnit.Text) Or Not IsNumeric(txtQuentityByPurchaseUnit.Text) Then
        Tr = MsgBox("You have not entered a valid quentity to issue", vbCritical, "Quentity?")
        txtQuentityByIssueUnit.SetFocus
        SendKeys "{Home}+{end}"
        Exit Function
    End If
    CanAdd = True
End Function

Private Function CanDelete() As Boolean
CanDelete = False
Dim Tr As Integer
If GridIssues.Rows < 2 Then
    Tr = MsgBox("There are no items to be removed", vbCritical, "No items")
    GridIssues.SetFocus
    Exit Function
End If
GridIssues.Col = 0
If Not IsNumeric(GridIssues.Text) Then
    Tr = MsgBox("You have not selected an item to delete", vbCritical, "No item")
    GridIssues.SetFocus
    Exit Function
End If
CanDelete = True
End Function

Private Sub GridIssues_Click()
    If GridIssues.Rows > 1 Then
        bttnDelete.Enabled = True
    Else
        bttnDelete.Enabled = False
    End If
    GridIssues.Col = GridIssues.Cols - 1
    GridIssues.ColSel = 0
End Sub

Private Sub Tem()
Dim temitem As New Item
End Sub

