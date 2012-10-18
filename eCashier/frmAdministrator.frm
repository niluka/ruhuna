VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdministrator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator"
   ClientHeight    =   8055
   ClientLeft      =   2085
   ClientTop       =   -2445
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10575
   Begin VB.CommandButton btnAddIDs 
      Caption         =   "Add IDs"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtUpto 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtTable 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   3120
      TabIndex        =   2
      Top             =   3600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Name"
      TabPicture(0)   =   "frmAdministrator.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Age"
      TabPicture(1)   =   "frmAdministrator.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Sex"
      TabPicture(2)   =   "frmAdministrator.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnAddForms 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      Caption         =   "Add Forms"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Upto"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

    
    
    

Private Sub btnAddForms_Click()
    Dim rsForm As New ADODB.Recordset
    Dim rsControl As New ADODB.Recordset
    Dim MyForm As Form
    Dim FormID As Long
    Dim i As Integer
    Dim MyControl As Control
    Dim temText As String
    For Each MyForm In Forms
        FormID = GetFormID(MyForm.Name, MyForm.Caption)
        For Each MyControl In MyForm.Controls
            With rsForm
                If TypeOf MyControl Is SSTab Then
                    For i = 0 To MyControl.Tabs - 1
                        If .State = 1 Then .Close
                        temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "' AND ControlIndex = " & i
                        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                        MyControl.Tab = i
                        If .RecordCount > 0 Then
                            !ControlText = GetControlText(MyControl)
                        Else
                            .AddNew
                            !FormID = FormID
                            !Control = MyControl.Name
                            !ControlType = GetControlType(MyControl)
                            !ControlText = GetControlText(MyControl)
                            !ControlIndex = i
                        End If
                        .Update
                        .Close
                    Next i
                Else
                    If .State = 1 Then .Close
                    temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "'"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ControlText = GetControlText(MyControl)
                    Else
                        .AddNew
                        !FormID = FormID
                        !Control = MyControl.Name
                        !ControlType = GetControlType(MyControl)
                        !ControlText = GetControlText(MyControl)
                    End If
                    .Update
                    .Close
                End If
            End With
        Next
    Next
End Sub



Private Sub AddControl()

End Sub

Private Sub btnAddIDs_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Long
    With rsTem
        temSql = "Select * from tbl" & txtTable.Text
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            For i = 1 To Val(txtUpto.Text - .RecordCount)
                .AddNew
                !IncomeBillID = 0
                .Update
            Next i
        Else
            For i = 1 To Val(txtUpto.Text)
                .AddNew
                !IncomeBillID = 0
                .Update
            Next i
        End If
    End With
    MsgBox "OK"
End Sub
