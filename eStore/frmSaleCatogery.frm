VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSaleCatogeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Categories"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
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
   ScaleHeight     =   8160
   ScaleWidth      =   9780
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   4680
      TabIndex        =   25
      Top             =   6720
      Width           =   4815
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Save"
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Cancel"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Save"
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
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   24
      Top             =   6720
      Width           =   4335
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Edit"
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
   End
   Begin VB.Frame Frame4 
      Height          =   6495
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   4335
      Begin MSDataListLib.DataCombo DtcSaleCategory 
         Height          =   5940
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   10478
         _Version        =   393216
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   4680
      TabIndex        =   18
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtProfitMargin 
         Height          =   360
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pay Mode"
         Height          =   1455
         Left            =   240
         TabIndex        =   26
         Top             =   4080
         Width           =   4335
         Begin VB.OptionButton optOther 
            Caption         =   "Other"
            Height          =   375
            Left            =   2280
            TabIndex        =   30
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optionCreditCard 
            Caption         =   "Credit Card"
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optionBankSlip 
            Caption         =   "Slip"
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optionCash 
            Caption         =   "Cash"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optionCredit 
            Caption         =   "Credit"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optionCheque 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo dtcPaymode 
         Height          =   360
         Left            =   240
         TabIndex        =   13
         Top             =   6000
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtPercentage 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   4335
      End
      Begin VB.TextBox txtCategoryName 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Customer type"
         Height          =   1455
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   4335
         Begin VB.OptionButton optUnit 
            Caption         =   "Hospital Unit"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton optionStaff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   2415
         End
         Begin VB.OptionButton optionBHT 
            Caption         =   "BHT"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton optionOutPation 
            Caption         =   "Out Patient"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Additional Profit Margin"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Discount Percent"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   5640
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Category Name"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmSaleCatogeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim rsViewCategory As New ADODB.Recordset
    Dim rsViewPaymode As New ADODB.Recordset
    Dim temSQL As String
    Dim A As Byte

Private Sub FillPaymentMode()
    With rsViewPaymode
        If .State = 1 Then .Close
        temSQL = "Select * From tblPaymentMethod Order by PaymentMethod"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcPaymode.RowSource = rsViewPaymode
        dtcPaymode.BoundColumn = "PaymentMethodID"
        dtcPaymode.ListField = "PaymentMethod"
    End With
End Sub

Private Sub FillSaleCategory()
    With rsViewCategory
        If .State = 1 Then .Close
        temSQL = "Select * From tblSaleCategory Order by SaleCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        Set DtcSaleCategory.RowSource = rsViewCategory
        DtcSaleCategory.BoundColumn = "SaleCategoryID"
        DtcSaleCategory.ListField = "SaleCategory"
    End With
End Sub

Private Sub bttnAdd_Click()
    Call BeforAdd
    txtCategoryName.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call BeforAddEdit
End Sub

Private Sub bttnChange_Click()
    If CheckValues = False Then Exit Sub
    Call EditCategoryDetails
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    txtCategoryName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Function ClearValues()
    txtCategoryName.Text = Empty
    optionOutPation.Value = False
    optionBHT.Value = False
    optionStaff.Value = False
    optUnit.Value = False
    dtcPaymode.BoundText = Empty
    txtPercentage.Text = Empty
    optionCash.Value = False
    optionCredit.Value = False
    optionCreditCard.Value = False
    optionBankSlip.Value = False
    optionCheque.Value = False
    optOther.Value = False
    txtProfitMargin.Text = Empty
End Function

Private Function CheckValues() As Boolean
    CheckValues = False
    If txtCategoryName.Text = Empty Then
        A = MsgBox("Enter Category Name", vbCritical + vbOKOnly, "Error"): txtCategoryName.SetFocus
        txtCategoryName.SetFocus
        Exit Function
    End If
    If optUnit.Value = False And optionOutPation.Value = False And optionBHT.Value = False And optionStaff.Value = False Then
        A = MsgBox("Select Customer Type", vbCritical + vbOKOnly, "Error")
        optionOutPation.SetFocus
        Exit Function
    End If
    If dtcPaymode.BoundText = Empty Then
        A = MsgBox("Select Pay Mode", vbCritical + vbOKOnly, "Error")
        dtcPaymode.SetFocus
        Exit Function
    End If
    If txtPercentage.Text = Empty Then
        A = MsgBox("Enter Percentage", vbCritical + vbOKOnly, "Error")
        txtPercentage.SetFocus
        Exit Function
    End If
    If optOther.Value = False And optionCash.Value = False And optionCredit.Value = False And optionCredit.Value = False And optionCheque.Value = False And optionBankSlip.Value = False And optionCreditCard.Value = False Then
        A = MsgBox("Select Payment Option", vbCritical + vbOKOnly, "Error")
        optionCash.SetFocus
        Exit Function
    End If
    CheckValues = True
End Function

Private Sub BeforAddEdit()
    Frame4.Enabled = True
    Frame1.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
End Sub

Private Sub BeforAdd()
    Frame4.Enabled = False
    Frame1.Enabled = True
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Call ClearValues
End Sub

Private Sub AfterEdit()
    Frame4.Enabled = False
    Frame1.Enabled = True
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
End Sub


Private Sub bttnSave_Click()
    If CheckValues = False Then Exit Sub
    Call SaveCategoryDetails
End Sub

Private Sub DisplaySeleted()
    Call ClearValues
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * From tblSaleCategory Where (SaleCategoryID = " & DtcSaleCategory.BoundText & ")"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            txtCategoryName.Text = !SaleCategory
            txtPercentage.Text = !SaleDiscountPercent
            dtcPaymode.BoundText = !PaymentMethodID
            If Not IsNull(!ProfitMargin) Then txtProfitMargin.Text = !ProfitMargin
            If !Cash = True Then optionCash.Value = True
            If !Credit = True Then optionCredit.Value = True
            If !Cheque = True Then optionCheque.Value = True
            If !Slips = True Then optionBankSlip.Value = True
            If !CreditCard = True Then optionCreditCard.Value = True
            If !Other = True Then optOther.Value = True
            If !OutPatient = True Then optionOutPation.Value = True
            If !InPatient = True Then optionBHT.Value = True
            If !Staff = True Then optionStaff.Value = True
            If !Unit = True Then optUnit.Value = True
        Else
        
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub SaveCategoryDetails()
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * From tblSaleCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SaleCategory = txtCategoryName.Text
        !SaleDiscountPercent = Val(txtPercentage.Text)
        !PaymentMethodID = Val(dtcPaymode.BoundText)
        If optionCash.Value = True Then
            !Cash = True
        Else
            !Cash = False
        End If
        If optionCredit.Value = True Then
            !Credit = True
        Else
            !Credit = False
        End If
        If optionCheque.Value = True Then
            !Cheque = True
        Else
            !Cheque = False
        End If
        If optionBankSlip.Value = True Then
            !Slips = True
        Else
            !Slips = False
        End If
        If optionCreditCard.Value = True Then
            !CreditCard = True
        Else
            !CreditCard = False
        End If
        If optOther.Value = True Then
            !Other = True
        Else
            !Other = False
        End If
        If optionOutPation.Value Then
            !OutPatient = True
        Else
            !OutPatient = False
        End If
        If optionBHT.Value Then
            !InPatient = True
        Else
            !InPatient = False
        End If
        If optionStaff.Value Then
            !Staff = True
        Else
            !Staff = False
        End If
        If optUnit.Value = True Then
            !Unit = True
        Else
            !Unit = False
        End If
        !ProfitMargin = Val(txtProfitMargin.Text)
        .Update
        Call BeforAddEdit
        Call ClearValues
        Call FillSaleCategory
        If .State = 1 Then .Close
        DtcSaleCategory.SetFocus
        DtcSaleCategory.Text = Empty
    End With
End Sub

Private Sub EditCategoryDetails()
With rsTem
    If .State = 1 Then .Close
    temSQL = "Select * From tblSaleCategory Where (SaleCategoryID = " & DtcSaleCategory.BoundText & ")"
    .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
    If .RecordCount > 0 Then
        !SaleCategory = txtCategoryName.Text
        !SaleDiscountPercent = Val(txtPercentage.Text)
        !PaymentMethodID = Val(dtcPaymode.BoundText)
        If optionCash.Value = True Then
            !Cash = True
        Else
            !Cash = False
        End If
        If optionCredit.Value = True Then
            !Credit = True
        Else
            !Credit = False
        End If
        If optionCheque.Value = True Then
            !Cheque = True
        Else
            !Cheque = False
        End If
        If optionBankSlip.Value = True Then
            !Slips = True
        Else
            !Slips = False
        End If
        If optionCreditCard.Value = True Then
            !CreditCard = True
        Else
            !CreditCard = False
        End If
        If optOther.Value = True Then
            !Other = True
        Else
            !Other = False
        End If
        If optionOutPation.Value Then
            !OutPatient = True
        Else
            !OutPatient = False
        End If
        If optionBHT.Value Then
            !InPatient = True
        Else
            !InPatient = False
        End If
        If optionStaff.Value Then
            !Staff = True
        Else
            !Staff = False
        End If
        If optUnit.Value = True Then
            !Unit = True
        Else
            !Unit = False
        End If
        !ProfitMargin = Val(txtProfitMargin.Text)
        .Update
    End If
    Call BeforAddEdit
    Call ClearValues
    Call FillSaleCategory
    If .State = 1 Then .Close
    DtcSaleCategory.SetFocus
    DtcSaleCategory.Text = Empty
End With
End Sub

Private Sub DtcSaleCategory_Click(Area As Integer)
    If IsNumeric(DtcSaleCategory.BoundText) = False Then Exit Sub
    Call DisplaySeleted
End Sub

Private Sub Form_Load()
    Call FillSaleCategory
    Call FillPaymentMode
    Call BeforAddEdit
End Sub

Private Sub optionBankSlip_Click()
    dtcPaymode.Text = optionBankSlip.Caption
End Sub

Private Sub optionBHT_Click()
    Call CheckOther
End Sub

Private Sub optionCash_Click()
    dtcPaymode.Text = optionCash.Caption
End Sub

Private Sub optionCheque_Click()
    dtcPaymode.Text = optionCheque.Caption
End Sub

Private Sub optionCredit_Click()
    dtcPaymode.Text = optionCredit.Caption
End Sub

Private Sub optionCreditCard_Click()
    dtcPaymode.Text = optionCreditCard.Caption
End Sub

Private Sub optionOutPation_Click()
    CheckOther
End Sub

Private Sub CheckOther()
    If optUnit.Value = True Then
        optOther.Value = True
        optionCash.Enabled = False
        optionCredit.Enabled = False
        optionCreditCard.Enabled = False
        optionCheque.Enabled = False
        optionBankSlip.Enabled = False
        optOther.Enabled = True
    Else
        optOther.Value = False
        optionCash.Enabled = True
        optionCredit.Enabled = True
        optionCreditCard.Enabled = True
        optionCheque.Enabled = True
        optionBankSlip.Enabled = True
        optOther.Enabled = False
    End If
End Sub

Private Sub optionStaff_Click()
    Call CheckOther
End Sub

Private Sub optUnit_Click()
    Call CheckOther
End Sub
