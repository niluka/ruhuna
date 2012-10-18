VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConsumption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumption"
   ClientHeight    =   6330
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
   ScaleHeight     =   6330
   ScaleWidth      =   14310
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   7320
      TabIndex        =   13
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmConsumption.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRStoreID"
      Tab(0).Control(1)=   "txtIStore"
      Tab(0).Control(2)=   "txtDisplay"
      Tab(0).Control(3)=   "txtVTM"
      Tab(0).Control(4)=   "txtVMP"
      Tab(0).Control(5)=   "txtAMP"
      Tab(0).Control(6)=   "txtVMPP"
      Tab(0).Control(7)=   "txtAMPP"
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(14)=   "Label13"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "History"
      TabPicture(1)   =   "frmConsumption.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridHistory"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Store Stock"
      TabPicture(2)   =   "frmConsumption.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "GridStore"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Total Stock"
      TabPicture(3)   =   "frmConsumption.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GridTotal"
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
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   20
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   14
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   15
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   16
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   17
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   18
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2880
         Width           =   4095
      End
      Begin MSFlexGridLib.MSFlexGrid GridStore 
         Height          =   4695
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridHistory 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin VB.Label Label14 
         Caption         =   "Store :"
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
         Left            =   -74880
         TabIndex        =   38
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
         Left            =   -74880
         TabIndex        =   37
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
         Left            =   -74880
         TabIndex        =   36
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
         Left            =   -74880
         TabIndex        =   35
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
         Left            =   -74880
         TabIndex        =   34
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
         Left            =   -74880
         TabIndex        =   33
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
         Left            =   -74880
         TabIndex        =   32
         Top             =   2880
         Width           =   3255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridIssues 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
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
   Begin VB.TextBox txtQuentityByIssueUnit 
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtQuentityByPurchaseUnit 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItemCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   4800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   1560
      TabIndex        =   10
      Top             =   5160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   2160
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
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&onsume"
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
      TabIndex        =   12
      Top             =   5760
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
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCCatogery 
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Reason"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Batch"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Issued By :"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Checked By:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   5880
      TabIndex        =   28
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPUnit 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Catogery"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Quentity"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmConsumption"
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
    Dim TemDouble As Double
    Dim temSql As String
    Dim BorderMargin As Integer
    Dim NewItem As New Item
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsStores As New ADODB.Recordset
    Dim rsBatch As New ADODB.Recordset
    Dim rsCCatogery As New ADODB.Recordset
    Dim rsTemConsume As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemUnit As New ADODB.Recordset
    Dim rsTemBatch As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset

Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    Call AddToGrid
    Call ClearAddValues
    dtcItemCategory.SetFocus
End Sub

Private Sub ClearAddValues()
    dtcItem.Text = Empty
    dtcItemCategory.Text = Empty
    dtcBatch.Text = Empty
    txtQuentityByIssueUnit.Text = Empty
    txtQuentityByPurchaseUnit.Text = Empty
    lblIUnit.Caption = Empty
    lblPunit.Caption = Empty
    dtcCCatogery.Text = Empty
    With GridHistory
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridStore
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridTotal
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    lblIUnit.Caption = NewItem.IUnit
    lblPunit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
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
        .Text = txtQuentityByIssueUnit.Text & " " & lblIUnit.Caption
        .Col = 6
        .Text = txtQuentityByIssueUnit.Text
        .Col = 7
        .Text = lblIUnit.Caption
        .Col = 8
        .CellAlignment = 7
        .Text = txtQuentityByPurchaseUnit.Text & " " & lblPunit.Caption
        .Col = 9
        .Text = lblPunit.Caption
        .Col = 10
        .Text = txtQuentityByPurchaseUnit.Text
        .Col = 11
        .CellAlignment = 1
        .Text = dtcCCatogery.Text
        .Col = 12
        .Text = dtcCCatogery.BoundText
        
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
'Reason 11
'ReasonID 12
    End With
End Sub

Private Sub FormatGrid()
With GridIssues
    .Cols = 13
    .Rows = 1
    
    BorderMargin = 100
    
    .ColWidth(0) = 600
    
    .ColWidth(2) = 1
    .ColWidth(3) = 700
    .ColWidth(4) = 1
    .ColWidth(5) = 1200
    .ColWidth(6) = 1
    .ColWidth(7) = 1
    .ColWidth(8) = 1000
    .ColWidth(9) = 1
    .ColWidth(10) = 1
    .ColWidth(11) = 1200
    .ColWidth(12) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7) + .ColWidth(8) + .ColWidth(9) + .ColWidth(10) + .ColWidth(11) + .ColWidth(12) + BorderMargin)
    
    
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
            Case 10: .Text = "Reason"
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
'reason 11
'CatogeryID 12



End With
End Sub

Private Function CanIssue() As Boolean
    Dim tr As Integer
    CanIssue = False
        If GridIssues.Rows <= 1 Then
            tr = MsgBox("There are no items to issue", vbCritical, "No items")
            dtcItem.SetFocus
            Exit Function
        End If
        If Not IsNumeric(dtcCheckedStaff.BoundText) Then
            tr = MsgBox("Please enter the staff member who checked the consumption", vbInformation, "Checked by?")
            dtcCheckedStaff.SetFocus
            Exit Function
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
    Call CheckConsume
    Call ClearIssue
End Sub

Private Sub ClearIssue()
    Me.txtAMP.Text = Empty
    Me.txtAMPP.Text = Empty
    Me.txtDisplay.Text = Empty
    Me.txtQuentityByIssueUnit.Text = Empty
    Me.txtIStore.Text = Empty
    Me.txtQuentityByPurchaseUnit.Text = Empty
    Me.txtRStoreID.Text = Empty
    Me.txtVMP.Text = Empty
    Me.txtVMPP.Text = Empty
    Me.txtVTM.Text = Empty
    Me.dtcBatch.Text = Empty
    Me.dtcCCatogery.Text = Empty
    Me.dtcItem.Text = Empty
    Me.dtcItemCategory.Text = Empty
    GridHistory.Clear
    GridStore.Clear
    GridTotal.Clear
End Sub

Private Sub CheckConsume()
    Dim RowItemID As Long
    Dim RowBatchID As Long
    Dim RowQuentity As Double
    Dim RowCCatogeryID As Long
    Dim TemNum As Integer
    Dim NoError As Boolean
    Dim tr As Integer
    With GridIssues
        .Visible = False
        NoError = True
        For TemNum = 1 To GridIssues.Rows - 1
            .Row = TemNum
            .Col = 2
            RowItemID = Val(.Text)
            NewItem.ID = RowItemID
            .Col = 4
            RowBatchID = Val(.Text)
            .Col = 6
            RowQuentity = Val(.Text)
            .Col = 12
            RowCCatogeryID = Val(.Text)
            If ConsumeStocks(UserStoreID, RowBatchID, RowQuentity) = False Then
                NoError = False
            Else
                With rsTemConsume
                    If .State = 1 Then .Close
                    temSql = "SELECT tblConsumption.ItemID, tblConsumption.Price, tblConsumption.BatchID, tblConsumption.StoreID, tblConsumption.StaffID, tblConsumption.CheckedStaffID, tblConsumption.Amount, tblConsumption.Date, tblConsumption.Time, tblConsumption.CategoryID FROM tblConsumption"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    .AddNew
                    !ItemID = RowItemID
                    !CategoryID = RowCCatogeryID
                    !BatchID = RowBatchID
                    !StoreID = UserStoreID
                    !StaffID = UserID
                    !CheckedStaffID = dtcCheckedStaff.BoundText
                    !Amount = RowQuentity
                    !Price = RowQuentity * NewItem.PPrice
                    !Date = Date
                    !Time = Now
                    .Update
                    .Close
                End With
                If .Rows = 2 And .Row = 1 Then
                    FormatGrid
                    Exit For
                Else
                    .RemoveItem (.Row)
                    TemNum = TemNum - 1
                End If
            End If
        Next
        If NoError = True Then
            tr = MsgBox("All the items were successfully added as consumed", vbInformation, "Successful")
        Else
            tr = MsgBox("One or more items were NOT successfully added as consumed", vbCritical, "Error")
        End If
        .Visible = True
    End With
End Sub

Private Function ConsumeStocks(ByVal IStoreIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
    On Error GoTo eh
    ConsumeStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            tr = MsgBox("There is no such drug batch", vbCritical, "Error")
            .Close
            Exit Function
        End If
        If !Stock < Quentity Then
            tr = MsgBox("There are no enough stocks in you store to transfer to another store", vbCritical, "No Enough Stocks")
            .Close
            Exit Function
        End If
        !Stock = !Stock - Quentity
        .Update
        .Close
    ConsumeStocks = True
    Exit Function

eh:
    If .State = 1 Then
        .CancelUpdate
        .Close
    End If
    tr = MsgBox("Could not deduct stocks from your store" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Function
    End With
End Function


Private Sub dtcItem_Change()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim TemdtcVal As Long
    Dim tr As Integer
    NewItem.ID = Val(dtcItem.BoundText)
    With GridHistory
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridStore
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridTotal
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    lblIUnit.Caption = NewItem.IUnit
    lblPunit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
    With rsBatch
        If .State = 1 Then .Close
        If DoNotAllowExpireConsumption = False Then
            temSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "WHERE (((tblBatch.ItemID)=" & Val(dtcItem.BoundText) & ") AND ((tblBatch.DOE)>'" & Format(Date, "MMMM dd yyyy") & "') AND ((tblBatchStock.Stock)>0) AND ((tblBatchStock.StoreID)=" & UserStoreID & " )) " & _
                        "ORDER BY tblBatch.DOE"
        Else
            temSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "WHERE (((tblBatch.ItemID)=" & Val(dtcItem.BoundText) & ") AND ((tblBatchStock.Stock)>0) AND ((tblBatchStock.StoreID)=" & UserStoreID & " )) " & _
                        "ORDER BY tblBatch.DOE"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then
            tr = MsgBox("There are no stocks", vbCritical, "No Stocks")
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
            TemdtcVal = !BatchID
        End If
    End With
    With dtcBatch
        Set .RowSource = rsBatch
        .ListField = "Batch"
        .BoundColumn = "BatchID"
        .BoundText = TemdtcVal
    End With
    
    txtQuentityByIssueUnit.Text = Empty
    txtQuentityByPurchaseUnit.Text = Empty
    Call FillStoreStocks
    Call FillAllStocks
    Call FillHistory
End Sub

Private Sub FillStoreStocks()
With GridStore
    .Visible = False

    .Clear
    .Cols = 4
    .Rows = 1
    .Row = 0
    .FixedCols = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Batch"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Stock (" & lblIUnit.Caption & ")"
    
    .Col = 2
    .CellAlignment = 4
    .Text = "Expiary"
    
    .Col = 3
    .CellAlignment = 4
    .Text = "Location"
    
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1600
    .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    
End With
With rsTemStore
    If .State = 1 Then .Close
    temSql = "SELECT tblBatch.*, tblBatchStock.*, tblLocation.Location " & _
                " FROM (tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID) LEFT JOIN tblLocation ON tblBatchStock.LocationID = tblLocation.LocationID " & _
                " WHERE tblBatch.ItemID=" & Val(dtcItem.BoundText) & " AND tblBatchStock.Stock > 0 AND tblBatchStock.StoreID = " & UserStoreID & " "
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            GridStore.Rows = GridStore.Rows + 1
            GridStore.Row = GridStore.Rows - 1
            GridStore.Col = 0
            GridStore.CellAlignment = 1
            GridStore.Text = !Batch
            GridStore.Col = 1
            GridStore.CellAlignment = 7
            GridStore.Text = !Stock
            GridStore.Col = 2
            GridStore.CellAlignment = 1
            GridStore.Text = Format(!DOE, ShortDateFormat)
            GridStore.Col = 3
            GridStore.CellAlignment = 1
            If IsNull(!Location) = False Then
                GridStore.Text = !Location
            Else
                GridStore.Text = Empty
            End If
            
            .MoveNext
        Wend
    End If
    GridStore.Visible = True
    .Close
End With
End Sub

Private Sub FillAllStocks()
With GridTotal
    .Visible = False

    .Clear
    .Cols = 4
    .Rows = 1
    .Row = 0
    .FixedCols = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Batch"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Stock (" & NewItem.IUnit & ")"
    
    .Col = 2
    .CellAlignment = 4
    .Text = "Expiary"
    
    .Col = 3
    .CellAlignment = 4
    .Text = "Department"
    
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1600
    .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    
End With
With rsTemStore
    If .State = 1 Then .Close
    temSql = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
                " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                " WHERE tblBatch.ItemID=" & Val(dtcItem.BoundText) & " AND tblBatchStock.Stock > 0 "
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            GridTotal.Rows = GridTotal.Rows + 1
            GridTotal.Row = GridTotal.Rows - 1
            GridTotal.Col = 0
            GridTotal.CellAlignment = 1
            GridTotal.Text = !Batch
            GridTotal.Col = 1
            GridTotal.CellAlignment = 7
            If Not IsNull(!Stock) Then
                GridTotal.Text = !Stock
            Else
                GridTotal.Text = 0
            End If
            GridTotal.Col = 2
            GridTotal.CellAlignment = 1
            GridTotal.Text = Format(!DOE, ShortDateFormat)
            GridTotal.Col = 3
            GridTotal.CellAlignment = 1
            If Not IsNull(!Store) Then
                GridTotal.Text = !Store
            Else
                GridTotal.Text = Empty
            End If
            .MoveNext
        Wend
    End If
    GridTotal.Visible = True
    .Close
End With

End Sub

Private Sub FillHistory()
With GridHistory
    .Visible = False

    .Clear
    .Cols = 4
    .Rows = 1
    .Row = 0
    .FixedCols = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Reason"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Date"
    
    .Col = 2
    .CellAlignment = 4
    .Text = "Stock (" & lblIUnit.Caption & ")"
    
    .Col = 3
    .CellAlignment = 4
    .Text = "Department"
    
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 1600
    .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    
End With
With rsTemStore
    If .State = 1 Then .Close
    temSql = "SELECT tblConsumptionCategory.ConsumptionCategory, tblConsumption.Date, tblConsumption.Amount, tblStaff.Name, tblConsumption.StoreID, * FROM (tblConsumptionCategory RIGHT JOIN tblConsumption ON tblConsumptionCategory.ConsumptionCategoryID = tblConsumption.CategoryID) LEFT JOIN tblStaff ON tblConsumption.StaffID = tblStaff.StaffID WHERE (((tblConsumption.StoreID)=" & UserStoreID & ") AND ((tblConsumption.ItemID)=" & dtcItem.BoundText & ")) ORDER BY tblConsumption.Date DESC"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    Dim i As Integer
    If .RecordCount > 0 Then
        While .EOF = False
            GridHistory.Rows = GridHistory.Rows + 1
            GridHistory.Row = GridHistory.Rows - 1
            GridHistory.Col = 0
            GridHistory.CellAlignment = 1
            GridHistory.Text = !ConsumptionCategory
            GridHistory.Col = 1
            GridHistory.CellAlignment = 7
            GridHistory.Text = Format(!Date, ShortDateFormat)
            GridHistory.Col = 2
            GridHistory.CellAlignment = 1
            GridHistory.Text = !Amount
            GridHistory.Col = 3
            GridHistory.CellAlignment = 1
            GridHistory.Text = !Name
            i = i + 1
            If i > 15 Then .MoveLast
            .MoveNext
        Wend
    End If
    GridHistory.Visible = True
    .Close
End With




End Sub

Private Sub dtcItemCategory_Click(Area As Integer)
    If IsNumeric(dtcItemCategory.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
End Sub

Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcItemCategory.BoundText & " order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub ListAllItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub dtcItemCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        dtcItemCategory.Text = Empty
        ListAllItems
    End If
End Sub


Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    With GridHistory
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridStore
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridTotal
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
End Sub

Private Sub FillCombos()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsItemCategory
    If .State = 1 Then .Close
    temSql = "SELECT * from tblItemCategory order by ItemCategory"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItemCategory
    Set .RowSource = rsItemCategory
    .ListField = "ItemCategory"
    .BoundColumn = "ItemCategoryID"
End With
With rsStaff
    If .State = 1 Then .Close
    temSql = "SELECT * from tblstaff order by listedname"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcIssueStaff
    Set .RowSource = rsStaff
    .ListField = "ListedName"
    .BoundColumn = "StaffID"
    .BoundText = UserID
End With
With dtcCheckedStaff
    Set .RowSource = rsStaff
    .ListField = "ListedName"
    .BoundColumn = "StaffID"
    .BoundText = UserID
End With

' *************************************
' To Do

With rsCCatogery
    If .State = 1 Then .Close
    temSql = "SELECT tblConsumptionCategory.* FROM tblConsumptionCategory ORDER BY tblConsumptionCategory.ConsumptionCategory"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCCatogery
    Set .RowSource = rsCCatogery
    .ListField = "ConsumptionCategory"
    .BoundColumn = "ConsumptionCategoryID"
End With


' *************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CanExit = False Then Cancel = True: Exit Sub
End Sub

Private Function CanExit() As Boolean
    CanExit = False
    Dim tr As Integer
    If GridIssues.Rows > 1 Then
        tr = MsgBox("There are " & (GridIssues.Rows - 1) & " items listed to be consumed, but the transction is not finalised. You have to delete all the items of confirn the consumption by clicking the issue button ", vbCritical, "Transaction NOT finalised")
        Me.ZOrder 0
        bttnIssue.SetFocus
        Exit Function
    End If
    CanExit = True
End Function

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If Not IsNumeric(dtcCCatogery.BoundText) Then
        tr = MsgBox("You have not selected the reason", vbCritical, "Reason?")
        dtcCCatogery.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcItem.BoundText) Then
        tr = MsgBox("You have not selected the item to consume", vbCritical, "Item?")
        dtcItem.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcBatch.BoundText) Then
        tr = MsgBox("You have not selected a batch to consume", vbCritical, "Batch?")
        dtcBatch.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtQuentityByIssueUnit.Text) Or Not IsNumeric(txtQuentityByPurchaseUnit.Text) Then
        tr = MsgBox("You have not entered a valid quentity to consume", vbCritical, "Quentity?")
        txtQuentityByIssueUnit.SetFocus
        SendKeys "{Home}+{end}"
        Exit Function
    End If
    Dim i As Integer
    GridIssues.Visible = False
    For i = 1 To GridIssues.Rows - 1
        GridIssues.Row = i
        GridIssues.Col = 2
        If GridIssues.Text = dtcItem.BoundText Then
            tr = MsgBox("You can't enter the same item twice for a single occasion. If you want to adjust the quentity, double click the item on the Grid and then make alterations", vbCritical, "Same item twice")
            GridIssues.Visible = True
            GridIssues.SetFocus
            Exit Function
        End If
    Next i
    GridIssues.Visible = True
    If CalculateStock(dtcItem.BoundText, dtcBatch.BoundText, UserStoreID).Amount < Val(txtQuentityByIssueUnit.Text) Then
        tr = MsgBox("There are no adequate stocks", vbCritical, "No adequate stocks")
        txtQuentityByIssueUnit.SetFocus
        SendKeys "{home}+{end}"
        Exit Function
    End If
    CanAdd = True
    
End Function

Private Function CanDelete() As Boolean
CanDelete = False
Dim tr As Integer
If GridIssues.Rows < 2 Then
    tr = MsgBox("There are no items to be removed", vbCritical, "No items")
    GridIssues.SetFocus
    Exit Function
End If
GridIssues.Col = 0
If Not IsNumeric(GridIssues.Text) Then
    tr = MsgBox("You have not selected an item to delete", vbCritical, "No item")
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


Private Sub GridIssues_DblClick()
    If GridIssues.Rows > 1 Then
        dtcItemCategory.Text = Empty
        With GridIssues
            .Col = 2
            dtcItem.BoundText = Val(.Text)
            
            .Col = 4
            dtcBatch.BoundText = Val(.Text)
            
            .Col = 10
            txtQuentityByPurchaseUnit.Text = Empty
            txtQuentityByPurchaseUnit.SetFocus
            txtQuentityByPurchaseUnit.Text = (.Text)
            
            .Col = 6
            txtQuentityByIssueUnit.Text = Empty
            txtQuentityByIssueUnit.SetFocus
            txtQuentityByIssueUnit.Text = Val(.Text)
            
            .Col = 12
            dtcCCatogery.BoundText = Val(.Text)
        End With
        dtcItem.SetFocus
        SendKeys "{Home}+{End}"
        bttnDelete_Click
    End If
   
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
'reason 11
'CatogeryID 12
    
End Sub

Private Sub txtQuentityByIssueUnit_GotFocus()
    TemDouble = Val(txtQuentityByIssueUnit.Text)
End Sub

Private Sub txtQuentityByIssueUnit_LostFocus()
    If TemDouble = Val(txtQuentityByIssueUnit.Text) Then Exit Sub
    If NewItem.IssueUnitsPerPack = 0 Then Exit Sub
    txtQuentityByPurchaseUnit.Text = Val(txtQuentityByIssueUnit.Text) / NewItem.IssueUnitsPerPack
End Sub

Private Sub txtQuentityByPurchaseUnit_GotFocus()
    TemDouble = Val(txtQuentityByPurchaseUnit.Text)
End Sub

Private Sub txtQuentityByPurchaseUnit_LostFocus()
    If TemDouble = Val(txtQuentityByPurchaseUnit.Text) Then Exit Sub
    txtQuentityByIssueUnit.Text = txtQuentityByPurchaseUnit.Text * NewItem.IssueUnitsPerPack
End Sub
