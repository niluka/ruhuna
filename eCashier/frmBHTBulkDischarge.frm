VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBHTBulkDischarge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT Bulk Discharge"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9285
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
   ScaleHeight     =   6810
   ScaleWidth      =   9285
   Begin VB.CommandButton btnNone 
      Caption         =   "None"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton btnAll 
      Caption         =   "All"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ListBox lstBHT 
      Height          =   6000
      Left            =   1080
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
   Begin VB.ListBox lstBHTID 
      Height          =   6060
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtDetails 
      Height          =   5535
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   12120
      TabIndex        =   1
      Top             =   9840
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
   Begin btButtonEx.ButtonEx btnDischarge 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Discharge"
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
   Begin VB.Label Label1 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBHTBulkDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim MyBHT As New clsBHT
    
Private Sub btnAll_Click()
    Dim i As Integer
    With lstBHT
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
End Sub

Private Sub btnDischarge_Click()
    
    Dim i As Integer
    
    i = MsgBox("Are you sure?", vbYesNo)
    
    If i = vbNo Then Exit Sub
    
    With lstBHT
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                DiachargeBHT (Val(lstBHTID.List(i)))
                DischargeFromRoom (Val(lstBHTID.List(i)))
            End If
        Next
    End With
    
    Call FillLists
    
End Sub
    
    
    
Private Sub DiachargeBHT(BHTID As Long)
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim MyBHTID As Long
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where BHTID = " & BHTID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Discharge = True
            !DOD = Date
            !TOD = Format(Date, "dd MMMM yyyy") & " " & Time
            !DisStaffID = UserID
            !Comments = "Bulk Discharge"
            .Update
        End If
        .Close
    End With

End Sub

    
    
Private Sub DisplayDetails(BHTID As Long)
    MyBHT.BHTID = BHTID
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
    
    
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub DischargeFromRoom(BHTID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomPatient where BHTID = " & BHTID & " Order by RoomPatientID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ToDate = Date
            !ToTime = Format(Date, "dd MMMM yyyy") & " " & Time
            .Update
        End If
        .Close
    End With

End Sub

Private Sub btnNone_Click()
    Dim i As Integer
    With lstBHT
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With

End Sub

Private Sub Form_Load()
    Call FillLists
End Sub

Private Sub FillLists()
    lstBHT.Clear
    lstBHTID.Clear
    Dim rsBHT As New ADODB.Recordset
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 0 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstBHT.AddItem !BHT
            lstBHTID.AddItem !BHTID
            .MoveNext
        Wend
    End With
End Sub

Private Sub lstBHT_Click()
    Call DisplayDetails(Val(lstBHTID.List(lstBHT.ListIndex)))
End Sub
