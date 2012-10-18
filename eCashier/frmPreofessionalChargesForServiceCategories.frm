VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPreofessionalChargesForServiceCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professional Charges for Service Categories"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12285
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
   ScaleHeight     =   6240
   ScaleWidth      =   12285
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   10920
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSDataListLib.DataList lstSPC 
      Height          =   3900
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6879
      _Version        =   393216
   End
   Begin VB.TextBox txtComments 
      Height          =   1320
      Left            =   7560
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox txtFee 
      Height          =   360
      Left            =   7560
      TabIndex        =   7
      Top             =   2160
      Width           =   4575
   End
   Begin VB.CheckBox chkCanChange 
      Caption         =   "Can C&hange"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo cmbServiceCategory 
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10920
      TabIndex        =   12
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
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   7560
      TabIndex        =   5
      Top             =   1680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Speciality"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Fee"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Service Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmPreofessionalChargesForServiceCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSPC As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    If IsNumeric(cmbServiceCategory.BoundText) = False Then
        MsgBox "Service Category?"
        cmbServiceCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSpeciality.BoundText) = False Then
        MsgBox "Speciality?"
        cmbSpeciality.SetFocus
        Exit Sub
    End If
    If IsNumeric(lstSPC.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call ClearValues
    Call FillList
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceProfessionalCharges"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ServiceCategoryID = Val(cmbServiceCategory.BoundText)
        !SpecialityID = Val(cmbSpeciality.BoundText)
        !Fee = Val(txtFee.Text)
        !Comments = txtComments.Text
        If chkCanChange.Value = 1 Then
            !CanChange = True
        Else
            !CanChange = False
        End If
        .Update
        .Close
    End With
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceProfessionalCharges where ServiceProfessionalChargesID = " & Val(lstSPC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SpecialityID = Val(cmbSpeciality.BoundText)
            !Fee = Val(txtFee.Text)
            !Comments = txtComments.Text
            If chkCanChange.Value = 1 Then
                !CanChange = True
            Else
                !CanChange = False
            End If
            .Update
        End If
        .Close
    End With
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(lstSPC.BoundText) = False Then
        MsgBox "Select Professional Charge"
        lstSPC.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceProfessionalCharges where ServiceProfessionalChargesID = " & Val(lstSPC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedDate = Date
            !DeletedTime = Time
            .Update
        End If
        .Close
    End With
    Call ClearValues
    Call FillList
    
End Sub

Private Sub cmbServiceCategory_Change()
    Call ClearValues
    Call FillList
End Sub

Private Sub ClearValues()
    cmbSpeciality.Text = Empty
    txtFee.Text = Empty
    txtComments.Text = Empty
    chkCanChange.Value = 0
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbServiceCategory, "ServiceCategory", True
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    
End Sub

Private Sub FillList()
    With rsSPC
        If .State = 1 Then .Close
        temSql = "SELECT tblSpeciality.Speciality & ' ' & format(tblServiceProfessionalCharges.Fee, '0.00') as SpecialityFee  , tblServiceProfessionalCharges.ServiceProfessionalChargesID " & _
                    "FROM tblSpeciality RIGHT JOIN tblServiceProfessionalCharges ON tblSpeciality.SpecialityID = tblServiceProfessionalCharges.SpecialityID " & _
                    "Where (((tblServiceProfessionalCharges.ServiceCategoryID) = " & Val(cmbServiceCategory.BoundText) & ") AND ((tblServiceProfessionalCharges.Deleted)=False ))"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstSPC
        Set .RowSource = rsSPC
        .ListField = "SpecialityFee"
        .BoundColumn = "ServiceProfessionalChargesID"
    End With
End Sub

Private Sub lstSPC_Click()
    Call ClearValues
    Call DisplayDetails
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceProfessionalCharges where ServiceProfessionalChargesID = " & Val(lstSPC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSpeciality.BoundText = !SpecialityID
            txtFee.Text = Format(!Fee, "0.00")
            txtComments.Text = Format(!Comments, "")
            If !CanChange = True Then
                chkCanChange.Value = 1
            Else
                chkCanChange.Value = 0
            End If
        End If
        .Close
    End With
End Sub
