VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14490
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
   ScaleHeight     =   7395
   ScaleWidth      =   14490
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   4680
      Width           =   5655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   7320
      TabIndex        =   22
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmTransfer.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRStoreID"
      Tab(0).Control(1)=   "txtRStore"
      Tab(0).Control(2)=   "txtIStore"
      Tab(0).Control(3)=   "txtDisplay"
      Tab(0).Control(4)=   "txtVTM"
      Tab(0).Control(5)=   "txtVMP"
      Tab(0).Control(6)=   "txtAMP"
      Tab(0).Control(7)=   "txtVMPP"
      Tab(0).Control(8)=   "txtAMPP"
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(13)=   "Label10"
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(15)=   "Label12"
      Tab(0).Control(16)=   "Label13"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Issue Store"
      TabPicture(1)   =   "frmTransfer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridIStore"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Receiving Store"
      TabPicture(2)   =   "frmTransfer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GridRStore"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Total Stock"
      TabPicture(3)   =   "frmTransfer.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "GridTotal"
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
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3840
         Visible         =   0   'False
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2880
         Width           =   4095
      End
      Begin MSFlexGridLib.MSFlexGrid GridRStore 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridIStore 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   3375
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
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
         Left            =   -74880
         TabIndex        =   29
         Top             =   2880
         Width           =   3255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridIssues 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6376
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
      Style           =   2
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
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItemCategory 
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
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   8640
      TabIndex        =   7
      Top             =   5880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   8640
      TabIndex        =   8
      Top             =   6360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
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
      Left            =   12000
      TabIndex        =   9
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Transfer"
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
      Left            =   13200
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
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label16 
      Caption         =   "Comments :"
      Height          =   255
      Left            =   7320
      TabIndex        =   43
      Top             =   4680
      Width           =   1215
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
      Left            =   7320
      TabIndex        =   19
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Checked By:"
      Height          =   255
      Left            =   7320
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
Attribute VB_Name = "frmTransfer"
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
    Dim NewItem As New Item
    Dim TemDouble As Double
    Dim temSql As String
    Dim BorderMargin As Integer
    
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsStores As New ADODB.Recordset
    Dim rsBatch As New ADODB.Recordset
    Dim rsTemBatch As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemUnit As New ADODB.Recordset
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
    With GridRStore
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridIStore
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
        .Text = txtQuentityByPurchaseUnit.Text
        .Col = 10
        .Text = lblPunit.Caption
        
        
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
    Dim tr As Integer
    CanIssue = False
        If GridIssues.Rows <= 1 Then
            tr = MsgBox("There are no items to issue", vbCritical, "No items")
            dtcItem.SetFocus
            Exit Function
        End If
        If IsNumeric(dtcTransaction.BoundText) = False Then
            tr = MsgBox("Please select a tranfer name", vbCritical, "Tranfer to?")
            dtcTransaction.SetFocus
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
    Call TransferItems
    If GridIssues.Rows > 1 Then
        
    Else
        Call ClearIssueValues
    End If
End Sub

Private Sub ClearIssueValues()
    dtcTransaction.Text = Empty
    txtComments.Text = Empty
    ClearAddValues
End Sub

Private Sub TransferItems()
    Dim RowItemID As Long
    Dim RowBatchID As Long
    Dim RowQuentity As Double
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
            .Col = 4
            RowBatchID = Val(.Text)
            .Col = 6
            RowQuentity = Val(.Text)
            
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
            tr = MsgBox("Transfer of all the items were successful", vbInformation, "Successfully transferred")
        Else
            tr = MsgBox("Some items are not transferred successfully", vbCritical, "Some are NOT transferred successfully")
        End If
        .Visible = True
    End With
End Sub

Private Function TransferStocks(ByVal IStoreIDValue As Long, ByVal RStoreIDValue As Long, ByVal ItemIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
'    On Error GoTo eh
    TransferStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            tr = MsgBox("There is such drug batch", vbCritical, "Error")
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
    temSql = "SELECT tblTransfer.TransferID, tblTransfer.price, tblTransfer.TransferCategoryID, tblTransfer.ItemID,tblTransfer.Amount, tblTransfer.EStoreID ,tblTransfer.BatchID, tblTransfer.SDate, tblTransfer.STime, tblTransfer.SStoreID, tblTransfer.SStaffID, tblTransfer.SCheckedUserID, tblTransfer.Issued, tblTransfer.IssueComments FROM tblTransfer"
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    .AddNew
    ![TransferCategoryID] = dtcTransaction.BoundText
    ![ItemID] = ItemIDValue
    NewItem.ID = ItemIDValue
    ![BatchID] = BatchIDValue
    ![SDate] = Date
    ![STime] = Time
    ![SStoreID] = UserStoreID
    ![SStaffID] = UserID
    If IsNumeric(dtcCheckedStaff.BoundText) = True Then ![SCheckedUserID] = dtcCheckedStaff.BoundText
    ![Issued] = True
    ![Amount] = Quentity
    !Price = Quentity * NewItem.PPrice
    ![EStoreID] = Val(txtRStoreID.Text)
    ![IssueComments] = txtComments.Text
    .Update
    .Close
    TransferStocks = True
    Exit Function

eh:
    .CancelUpdate
    .Close
    tr = MsgBox("Could not deduct stocks from your store" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Function
    End With
End Function


Private Sub dtcItem_Click(Area As Integer)
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim TemdtcVal As Long
    Dim tr As Integer
    NewItem.ID = Val(dtcItem.BoundText)
    lblIUnit.Caption = NewItem.IUnit
    lblPunit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
    With GridIStore
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridRStore
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
'    With rsBatch
'        If .State = 1 Then .Close
'        If DoNotAllowExpireConsumption = False Then
'            TemSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
'                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
'                        "WHERE (((tblBatch.ItemID)=" & Val(dtcItem.BoundText) & ") AND ((tblBatch.DOE)>'" & Format(Date, "MMMM dd yyyy") & "') AND ((tblBatchStock.Stock)>0) AND ((tblBatchStock.StoreID)=" & UserStoreID & " )) " & _
'                        "ORDER BY tblBatch.DOE"
'        Else
'            TemSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
'                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
'                        "WHERE (((tblBatch.ItemID)=" & Val(dtcItem.BoundText) & ") AND ((tblBatchStock.Stock)>0) AND ((tblBatchStock.StoreID)=" & UserStoreID & " )) " & _
'                        "ORDER BY tblBatch.DOE"
'        End If
'        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount < 1 Then
'            TR = MsgBox("There are no stocks", vbCritical, "No Stocks")
'            .Close
'            dtcItem.SetFocus
'            With dtcBatch
'                Set .RowSource = Nothing
'                .ListField = Empty
'                .BoundColumn = Empty
'                .BoundText = Empty
'            End With
'            Exit Sub
'        Else
'            .MoveFirst
'            TemdtcVal = !BatchID
'        End If
'    End With
'    With dtcBatch
'        Set .RowSource = rsBatch
'        .ListField = "Batch"
'        .BoundColumn = "BatchID"
'        .BoundText = TemdtcVal
'    End With
'
'    txtQuentityByIssueUnit.Text = Empty
'    txtQuentityByPurchaseUnit.Text = Empty
    
    Call FillIStoreStocks
    Call FillRStoreStocks
    Call FillAllStocks

End Sub


Private Sub FillIStoreStocks()
With GridIStore
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
            GridIStore.Rows = GridIStore.Rows + 1
            GridIStore.Row = GridIStore.Rows - 1
            GridIStore.Col = 0
            GridIStore.CellAlignment = 1
            GridIStore.Text = !Batch
            GridIStore.Col = 1
            GridIStore.CellAlignment = 7
            GridIStore.Text = !Stock
            GridIStore.Col = 2
            GridIStore.CellAlignment = 1
            GridIStore.Text = Format(!DOE, ShortDateFormat)
            GridIStore.Col = 3
            GridIStore.CellAlignment = 1
            If IsNull(!Location) = False Then
                GridIStore.Text = !Location
            Else
                GridIStore.Text = Empty
            End If
            
            .MoveNext
        Wend
    End If
    GridIStore.Visible = True
    .Close
End With
End Sub

Private Sub FillRStoreStocks()
If Not IsNumeric(txtRStoreID.Text) Then Exit Sub

With GridRStore
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
                " WHERE tblBatch.ItemID=" & Val(dtcItem.BoundText) & " AND tblBatchStock.Stock > 0 AND tblBatchStock.StoreID = " & Val(txtRStoreID.Text) & " "
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            GridRStore.Rows = GridRStore.Rows + 1
            GridRStore.Row = GridRStore.Rows - 1
            GridRStore.Col = 0
            GridRStore.CellAlignment = 1
            GridRStore.Text = !Batch
            GridRStore.Col = 1
            GridRStore.CellAlignment = 7
            GridRStore.Text = !Stock
            GridRStore.Col = 2
            GridRStore.CellAlignment = 1
            GridRStore.Text = Format(!DOE, ShortDateFormat)
            GridRStore.Col = 3
            GridRStore.CellAlignment = 1
            If IsNull(!Location) = False Then
                GridRStore.Text = !Location
            Else
                GridRStore.Text = Empty
            End If
            
            .MoveNext
        Wend
    End If
    GridRStore.Visible = True
    .Close
End With
End Sub



Private Sub dtcItem_LostFocus()
    Dim tr As Integer
    Dim TemdtcVal As Long
    
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

End Sub

Private Sub dtcItemCategory_Click(Area As Integer)
    If Trim(dtcItemCategory.Text) <> "" And IsNumeric(dtcItemCategory.BoundText) = True Then
        ListSelectedItems
    Else
        ListAllItems
    End If
    dtcItem.Text = Empty
End Sub

Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcItemCategory.BoundText & " order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
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
    If KeyCode = vbKeyEscape Then dtcItemCategory.Text = Empty
End Sub

Private Sub dtctransaction_Click(Area As Integer)
If IsNumeric(dtcTransaction.BoundText) = True Then
    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT tblStore.Store, tblTransferCategory.RStoreID " & _
                    "FROM tblStore RIGHT JOIN tblTransferCategory ON tblStore.StoreID = tblTransferCategory.RStoreID " & _
                    "WHERE (((tblTransferCategory.TransferCategoryID)=" & dtcTransaction.BoundText & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            txtRStore.Text = Empty
            txtRStoreID.Text = Empty
            txtIStore.Text = Empty
            .Close
            Exit Sub
        Else
            If Not IsNull(!Store) Then txtRStore.Text = !Store
            txtIStore.Text = UserStore
            If Not IsNull(!RStoreID) Then txtRStoreID.Text = !RStoreID
        End If
    End With
Else
    txtRStore.Text = Empty
    txtRStoreID.Text = Empty
End If

End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    dtcIssueStaff.BoundText = UserID
    dtcCheckedStaff.BoundText = UserID
End Sub

Private Sub FillCombos()
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
    temSql = "SELECT tblTransferCategory.TransferCategoryID, tblTransferCategory.TransferCategory FROM tblTransferCategory WHERE (((tblTransferCategory.IStoreID)=" & UserStoreID & ")) ORDER BY tblTransferCategory.TransferCategory"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcTransaction
    Set .RowSource = rsStores
    .ListField = "TransferCategory"
    .BoundColumn = "TransferCategoryID"
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
        tr = MsgBox("There are " & (GridIssues.Rows - 1) & " items listed to be issued, but the transction is not finalised. You have to delete all the items of confirn the issue by clicking the issue button ", vbCritical, "Transaction NOT finalised")
        Me.ZOrder 0
        bttnIssue.SetFocus
        Exit Function
    End If
    CanExit = True
End Function

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If Not IsNumeric(dtcTransaction.BoundText) Then
        tr = MsgBox("You have not selected the transaction", vbCritical, "Transaction?")
        dtcTransaction.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcItem.BoundText) Then
        tr = MsgBox("You have not selected the item to transfer", vbCritical, "Item?")
        dtcItem.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcBatch.BoundText) Then
        tr = MsgBox("You have not selected a batch to transfer", vbCritical, "Batch?")
        dtcBatch.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtQuentityByIssueUnit.Text) Or Not IsNumeric(txtQuentityByPurchaseUnit.Text) Then
        tr = MsgBox("You have not entered a valid quentity to issue", vbCritical, "Quentity?")
        txtQuentityByIssueUnit.SetFocus
        SendKeys "{Home}+{end}"
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

