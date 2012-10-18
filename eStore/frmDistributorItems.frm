VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDistributorItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributor Items"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
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
   ScaleHeight     =   8100
   ScaleWidth      =   7185
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   7440
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11668
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dtcDistributor 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "&Distributor"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDistributorItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim rsViewDistributor As New ADODB.Recordset
    Dim CsetPrinter As New cSetDfltPrinter
    Dim temSql As String
    
Private Sub btnPrint_Click()
    If Not IsNumeric(dtcDistributor.BoundText) Then Exit Sub
    Dim RetVal As Integer
    Dim TemResponce As Integer
    Dim temTopic As String
    Dim temSubTopic As String
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            With dtrDistributorItem
                Set .DataSource = rsItem
                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
                temTopic = "Items Purchased from " & dtcDistributor.Text
                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
                .Sections("Section1").Controls.Item("txtItem").DataField = "Display"
                .Caption = temTopic & " - " & temSubTopic
                .Show
            End With
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select





End Sub

Private Sub dtcDistributor_Click(Area As Integer)
    GridItem.Clear
    GridItem.Rows = 1
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
With rsViewDistributor
    If .State = 1 Then .Close
    temSql = "Select * from tblDistrubutor Order by DistributorName"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcDistributor
    Set .RowSource = rsViewDistributor
    .ListField = "DistributorName"
    .BoundColumn = "DistributorID"
End With

End Sub

Private Sub FillGrid()
    If Not IsNumeric(dtcDistributor.BoundText) Then Exit Sub
    Dim i As Integer
    Dim r As Integer
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.* FROM tblItem RIGHT JOIN tblItemDistributor ON tblItem.ItemID = tblItemDistributor.ItemID WHERE tblItemDistributor.DistributorID = " & Val(dtcDistributor.BoundText) & "  ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            GridItem.Cols = 2
            GridItem.ColWidth(1) = GridItem.Width - GridItem.ColWidth(0) - 150
            GridItem.Rows = .RecordCount + 1
            r = 0
            For i = 1 To .RecordCount
                If IsNull(!Display) = False And IsNull(!ItemID) = False Then
                    r = r + 1
                    GridItem.TextMatrix(r, 0) = r
                    GridItem.TextMatrix(r, 1) = !Display
                End If
                .MoveNext
            Next
        End If
    End With
    GridItem.Rows = r + 1
End Sub
