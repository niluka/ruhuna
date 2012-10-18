VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmChangeRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Room"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
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
   ScaleHeight     =   5850
   ScaleWidth      =   6165
   Begin VB.TextBox txtDetails 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3840
      Width           =   5895
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67043331
      CurrentDate     =   39958
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   3240
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
   Begin btButtonEx.ButtonEx btnChange 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Change"
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
   Begin MSDataListLib.DataCombo cmbCR 
      Height          =   360
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbNR 
      Height          =   360
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67043330
      CurrentDate     =   39958
   End
   Begin VB.Label Label6 
      Caption         =   "Details"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "New Room"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Current Room"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmChangeRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim MyBHT As New clsBHT
    Dim rsBHT As New ADODB.Recordset
    
Private Sub btnChange_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "Please select a BHT"
        cmbBHT.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbNR.BoundText) = False Then
        MsgBox "Pelase select a room"
        cmbNR.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomPatient where BHTID = " & Val(cmbBHT.BoundText) & " Order by RoomPatientID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ToDate = dtpDate.Value
            !ToTime = Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTime.Value
            .Update
        End If
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomPatient where RoomPatientID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !FromDate = dtpDate.Value
        !FromTime = Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTime.Value
        !RoomID = Val(cmbNR.BoundText)
        !BHTID = Val(cmbBHT.BoundText)
'        !ToDate = adFieldIsNull
'        !ToTime = adFieldIsNull
        .Update
    End With
    Call ClearValues
    cmbBHT.SetFocus
End Sub

Private Sub ClearValues()
    cmbBHT.Text = Empty
    cmbNR.Text = Empty
    cmbCR.Text = Empty
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbBHT_Change()
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "Select * from tblRoomPatient where BHTID = " & Val(cmbBHT.BoundText) & " order by RoomPatientID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbCR.BoundText = !RoomID
        Else
            cmbCR.Text = Empty
        End If
        .Close
    End With
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call DisplayDetails
End Sub

Private Sub DisplayDetails(): On Error Resume Next
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



Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    dtpDate.MaxDate = Date

End Sub

Private Sub FillCombos()
'    Dim BHT As New clsFillCombos
'    BHT.FillAnyCombo cmbBHT, "BHT", False
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 0 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With

    Dim CR As New clsFillCombos
    CR.FillAnyCombo cmbCR, "Room", False
    Dim NR As New clsFillCombos
    NR.FillAnyCombo cmbNR, "Room", False
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub
