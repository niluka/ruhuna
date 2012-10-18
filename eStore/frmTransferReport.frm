VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransferReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trnandfer Report"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
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
   ScaleHeight     =   5145
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   39566
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   39566
      End
      Begin VB.Label Label1 
         Caption         =   "Date From"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "By Department"
      TabPicture(0)   =   "frmTransferReport.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "bttnViewbyStore"
      Tab(0).Control(2)=   "dtcDepartment"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "By Cons. Category"
      TabPicture(1)   =   "frmTransferReport.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "dtcConsumption"
      Tab(1).Control(2)=   "ButtonEx1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "By Staff"
      TabPicture(2)   =   "frmTransferReport.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "dtcStaff"
      Tab(2).Control(2)=   "bttnViewbyStaff"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "By All"
      TabPicture(3)   =   "frmTransferReport.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "bttnAllbyDate"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin btButtonEx.ButtonEx bttnAllbyDate 
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View Report"
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
      Begin MSDataListLib.DataCombo dtcDepartment 
         Height          =   360
         Left            =   -71040
         TabIndex        =   8
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnViewbyStore 
         Height          =   375
         Left            =   -71040
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View Report"
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
      Begin btButtonEx.ButtonEx ButtonEx1 
         Height          =   375
         Left            =   -71040
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View Report"
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
      Begin MSDataListLib.DataCombo dtcConsumption 
         Height          =   360
         Left            =   -71040
         TabIndex        =   11
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnViewbyStaff 
         Height          =   375
         Left            =   -71040
         TabIndex        =   12
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View Report"
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
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   360
         Left            =   -71040
         TabIndex        =   13
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Department Name"
         Height          =   255
         Left            =   -73800
         TabIndex        =   16
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Consumption Name"
         Height          =   255
         Left            =   -73800
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Staff Name"
         Height          =   255
         Left            =   -73800
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmTransferReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsStore As New ADODB.Recordset
Dim rsTem As New ADODB.Recordset
Dim rsViewDepartment As New ADODB.Recordset
Dim rsTem1 As New ADODB.Recordset
Dim rsStaff As New ADODB.Recordset
Dim rsViewCatogeryName As New ADODB.Recordset
Dim temSQL As String

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub bttnAllbyDate_Click()

With rsTem
    If .State = 1 Then .Close
    temSQL = ""

    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    
    If .RecordCount = 0 Then Exit Sub

        With dtrConsumsionAll
        Set .DataSource = rsTem
        .Sections("Section4").Controls("lblTopic").Caption = "Discard Products"
        .Sections("Section4").Controls("lblSubTopic").Caption = "Date From  :  " & DTPicker1.Value & "To  :  " & DTPicker2.Value
        .Show
        
        End With
End With
End Sub

Private Sub bttnViewbyStore_Click()
If IsNumeric(dtcDepartment.BoundText) = False Then Exit Sub
With rsTem
    If .State = 1 Then .Close
    temSQL = "SELECT tblTransfer.*, tblSStaff.Name AS SendStaff, tblBatch.Batch, ([tblEStore].[store]) AS SStore, tblEStaff.Name AS EndStaff, ([tblEStore].[Store]) AS EStore, tblItem.Display FROM (tblStore AS tblEStore RIGHT JOIN (tblStaff AS tblEStaff RIGHT JOIN (tblStore AS tblSStore RIGHT JOIN (tblStaff AS tblSStaff RIGHT JOIN (tblTransfer LEFT JOIN tblBatch ON tblTransfer.BatchID = tblBatch.BatchID) ON tblSStaff.StaffID = tblTransfer.SStaffID) ON tblSStore.StoreID = tblTransfer.SStoreID) ON tblEStaff.StaffID = tblTransfer.EStaffID) ON tblEStore.StoreID = tblTransfer.EStoreID) LEFT JOIN tblItem ON tblTransfer.ItemID = tblItem.ItemID Where SStoreID = " & dtcDepartment.BoundText & " "
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    
    If .RecordCount = 0 Then Exit Sub
    
    
        With dtrTrnsferByStore
        Set .DataSource = rsTem
        .Sections("Section4").Controls("lblTopic").Caption = "Trnsfer Products Order By Depatment"
        .Sections("Section4").Controls("lblSubTopic").Caption = "Department Name    : " & dtcDepartment.Text
        .Show
        
        End With
End With
End Sub

Private Sub bttnViewbyStaff_Click()
If IsNumeric(dtcStaff.BoundText) = False Then Exit Sub
With rsTem1
    If .State = 1 Then .Close
    temSQL = ""
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    
    If .RecordCount = 0 Then Exit Sub

        With dtrConsumsionByStaff
        Set .DataSource = rsTem1
        .Sections("Section4").Controls("lblTopic").Caption = "Discard Products Order By Staff Name"
        .Sections("Section4").Controls("lblSubTopic").Caption = "Staff Name    : " & dtcStaff.Text
        .Show
        
        End With
End With

End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
Call FillDepartment
Call FillStaff
Call FillConsumtionCategory
End Sub

Private Sub FillConsumtionCategory()
    With rsViewCatogeryName
        If .State = 1 Then .Close
        .Open "Select* From tblConsumptionCategory Order By ConsumptionCategory", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcConsumption.RowSource = rsViewCatogeryName
            dtcConsumption.ListField = "ConsumptionCategory"
            dtcConsumption.BoundColumn = "ConsumptionCategoryID"
        End If
    End With
End Sub

Private Sub FillDepartment()
    With rsViewDepartment
        If .State = 1 Then .Close
        .Open "Select tblStore.* From tblStore Order By Store", cnnStores, adOpenStatic, adLockReadOnly
    
        If .RecordCount = 0 Then Exit Sub
        Set dtcDepartment.RowSource = rsViewDepartment
        dtcDepartment.ListField = "Store"
        dtcDepartment.BoundColumn = "StoreID"
    End With
End Sub

Private Sub FillStaff()

    With rsStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff order by name"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        
        Set dtcStaff.RowSource = rsStaff
        dtcStaff.ListField = "ListedName"
        dtcStaff.BoundColumn = "StaffID"
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsStore.State = 1 Then rsStore.Close: Set rsStore = Nothing
If rsTem.State = 1 Then rsTem.Close: Set rsTem = Nothing
If rsViewDepartment.State = 1 Then rsViewDepartment.Close: Set rsViewDepartment = Nothing
If rsTem1.State = 1 Then rsTem1.Close: Set rsTem1 = Nothing
If rsStaff.State = 1 Then rsStaff.Close: Set rsStaff = Nothing
If rsViewCatogeryName.State = 1 Then: Set rsViewCatogeryName = Nothing
End Sub

