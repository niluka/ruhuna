VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmReverseDischarge1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Discharge"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
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
   ScaleHeight     =   7800
   ScaleWidth      =   9375
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
   Begin btButtonEx.ButtonEx btnReverse 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Reverse Discharge"
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
   Begin VB.TextBox txtDetails 
      Height          =   5175
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   5895
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   6900
      Left            =   120
      TabIndex        =   0
      Tag             =   "Select"
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   12171
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin VB.Label Label8 
      Caption         =   "Details"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmReverseDischarge1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewBHT As New ADODB.Recordset
    Dim MyBHT As New clsBHT
    Dim temSql As String
        
Private Sub btnReverse_Click()

    On Error GoTo eh

    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    
    Dim MyBHTID As Long
    
    i = MsgBox("Are you sure you want to reverse of the discharge?", vbYesNo)
    
    If i = vbNo Then Exit Sub
    
    

    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            
            UpdateCompanyBalance MyBHT.HealthSchemeSupplierID, !Balance, True, False, False
            
            !Discharge = False
            !DOD = Null
            !TOD = Null
            !DisStaffID = UserID
            !AdmissionRate = 0
            !InitialLinanRate = 0
            !LaterLinanRate = 0
            !MaintananceRate = 0
            !NursingRate = 0
            
            
            !AdditionalCharge = 0
            !Balance = 0
            
            !AdmissionCharge = 0
            !LinanCharge = 0
            !RoomCharge = 0
            !ServicesCharge = 0
            !ProfessionalCharge = 0
            !MaintananceCharge = 0
            !NursingCharge = 0
            !MedicineCharge = 0
            !TotalCharge = 0
            !Payments = 0

            
            !FAdmissionCharge = 0
            !FLinanCharge = 0
            !FRoomCharge = 0
            !FServicesCharge = 0
            !FProfessionalCharge = 0
            !FMaintananceCharge = 0
            !FNursingCharge = 0
            !FMedicineCharge = 0
            !FTotalCharge = 0
            !FPayments = 0
            !FAdditionalCharge = 0
            
            !Price = 0
            !Discount = 0
            !DiscountPercent = 0
            !NetPrice = 0
            
            .Update
        End If
        .Close
    End With
    
    
    MyBHTID = Val(cmbBHT.BoundText)
    Call FillCombos
    cmbBHT.Text = Empty

    MsgBox "DIscharge is successfully reversed."
    Exit Sub

eh:

    MsgBox "Error. Discharge NOT Reversed" & vbNewLine & Err.Number & vbNewLine & Err.Description
    Exit Sub


End Sub

Private Sub cmbBHT_Change()
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call DisplayDetails
End Sub


Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsViewBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 1 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsViewBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With

End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    'temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
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

