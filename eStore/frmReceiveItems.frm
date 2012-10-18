VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmReceiveItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receive Items"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15225
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
   ScaleHeight     =   7770
   ScaleWidth      =   15225
   Begin VB.TextBox txtComments 
      Height          =   855
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   5640
      Width           =   5655
   End
   Begin btButtonEx.ButtonEx bttnReceive 
      Default         =   -1  'True
      Height          =   495
      Left            =   12480
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Receive"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   13800
      TabIndex        =   2
      Top             =   5640
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
   Begin VB.ListBox lstTx 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   720
      Width           =   14895
   End
   Begin VB.ListBox lstTxID 
      Height          =   4860
      Left            =   14280
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   7200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   6600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Checked By:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Issued By :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   $"frmReceiveItems.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   14415
   End
   Begin VB.Label Label1 
      Caption         =   "Items to receive"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmReceiveItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTxItems As New ADODB.Recordset
    Dim rsBatchStock As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    
    Dim temSql As String
    Dim NewItem As New Item
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnReceive_Click()
    Dim tr As Integer
    
    If lstTx.ListCount = 0 Or lstTxID.ListCount = 0 Then
        tr = MsgBox("There are no items to be received", vbInformation, "No Items")
        Exit Sub
    End If
    
    Dim i As Integer
    Dim Selected As Boolean
    Dim TemAmount  As Double
    Dim TemBatchID As Long
    
    Selected = False
    For i = 0 To lstTx.ListCount - 1
        If lstTx.Selected(i) = True Then Selected = True
    Next i
    If Selected = False Then
        tr = MsgBox("You have not selected any item to receive", vbCritical, "Not Selected")
        lstTx.SetFocus
        Exit Sub
    End If
    
    For i = 0 To lstTx.ListCount - 1
        If lstTx.Selected(i) = True Then
            lstTxID.ListIndex = i
            With rsTxItems
                If .State = 1 Then .Close
                temSql = "SELECT tblTransfer.EDate, tblTransfer.Amount, tblTransfer.BatchID, tblTransfer.ETime, tblTransfer.EStaffID, tblTransfer.ECheckedStaffID, tblTransfer.Received, tblTransfer.ReceiveComments " & _
                            "From tblTransfer " & _
                            "WHERE (((tblTransfer.TransferID)=" & Val(lstTxID.Text) & "))"
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                !EDate = Date
                !ETime = Now
                !EStaffID = dtcIssueStaff.BoundText
                !ECheckedStaffID = dtcCheckedStaff.BoundText
                !Received = True
                !ReceiveComments = txtComments.Text
                TemAmount = !Amount
                TemBatchID = !BatchID
                .Update
            End With
            With rsBatchStock
                If .State = 1 Then .Close
                temSql = "SELECT tblBatchStock.BatchID, tblBatchStock.StoreID, tblBatchStock.Stock " & _
                            "From tblBatchStock " & _
                            "WHERE (((tblBatchStock.BatchID)=" & TemBatchID & ") AND ((tblBatchStock.StoreID)=" & UserStoreID & ")) "
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount >= 1 Then
                    !Stock = !Stock + TemAmount
                Else
                    .AddNew
                    !BatchID = TemBatchID
                    !StoreID = UserStoreID
                    !Stock = TemAmount
                End If
                .Update
                .Close
            End With
        End If
    Next
    Call FillList
End Sub

Private Sub Form_Load()
    Call FillList
    Call FillCombos
    dtcCheckedStaff.BoundText = UserID
    dtcIssueStaff.BoundText = UserID
End Sub

Private Sub FillList()
    lstTx.Clear
    lstTxID.Clear
    Dim temText As String
    With rsTxItems
        If .State = 2 Then .Close
        temSql = "SELECT tblTransferCategory.TransferCategory, tblItem.Display, tblBatch.Batch,tblTransfer.TransferID, tblTransfer.SDate, tblTransfer.STime, tblTransfer.Amount, tblTransfer.ItemID, tblBatch.BatchID, tblTransfer.SStoreID, tblTransfer.Comments " & _
                    "FROM (tblBatch RIGHT JOIN (tblTransfer LEFT JOIN tblTransferCategory ON tblTransfer.TransferCategoryID = tblTransferCategory.TransferCategoryID) ON tblBatch.BatchID = tblTransfer.BatchID) LEFT JOIN tblItem ON tblBatch.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblTransfer.Issued)=1) AND ((tblTransfer.Received)=0) AND ((tblTransfer.EStoreID)=" & UserStoreID & "))" & _
                    "ORDER BY tblTransfer.SDate, tblTransferCategory.TransferCategory, tblTransfer.TransferID"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            While .EOF = False
                NewItem.ID = !ItemID
                temText = Left(!TransferCategory, 20)
                temText = temText & vbTab & Left(!Display & Space(20), 40)
                temText = temText & vbTab & Right(Space(20) & !Batch, 10)
                temText = temText & vbTab & Left(Format(!SDate, ShortDateFormat) & Space(20), 10)
                temText = temText & vbTab & Left(!STime & Space(20), 10)
                temText = temText & vbTab & Right(Space(20) & !Amount & Space(1), 12)
                temText = temText & Left(NewItem.IUnit & Space(20), 10)
'                temText = temText & vbTab & Left(!BatchID & Space(20), 10)
'                temText = temText & vbTab & Left(!SStoreID & Space(20), 10)
'                temText = temText & vbTab & Left(!Comments & Space(20), 10)
                lstTx.AddItem temText
                lstTxID.AddItem !TransferID
                .MoveNext
            Wend
        End If
    
    
    End With
End Sub

Private Sub FillCombos()
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

End Sub

