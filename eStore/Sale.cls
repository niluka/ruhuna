VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim SaleDiscountPercentValue As Double
    Dim ProfitMarginValue As Double
    Dim CashValue As Boolean
    Dim CreditValue As Boolean
    Dim CreditCardValue As Boolean
    Dim ChequeValue As Boolean
    Dim OtherValue As Boolean
    Dim SaleCategoryIDValue As Long
    Dim SaleCategoryValue As String
    Dim OutPatientValue As Boolean
    Dim InPatientValue As Boolean
    Dim StaffValue As Boolean
    Dim UnitValue As Boolean
    Dim temSQL As String
    Dim rsTemSale As New ADODB.Recordset

Public Property Let SaleCategoryID(IDValue As Long)
    SaleCategoryIDValue = IDValue
    With rsTemSale
        If .State = 1 Then .Close
        temSQL = "SELECT tblSaleCategory.SaleCategoryID, tblSaleCategory.SaleCategory,  tblSaleCategory.ProfitMargin,tblSaleCategory.SaleDiscountPercent, tblSaleCategory.Cash, tblSaleCategory.Credit, tblSaleCategory.Cheque, tblSaleCategory.CreditCard, tblSaleCategory.Other, tblSaleCategory.OutPatient, tblSaleCategory.InPatient, tblSaleCategory.Staff, tblSaleCategory.Unit " & _
                    " FROM tblSaleCategory " & _
                    " WHERE (((tblSaleCategory.SaleCategoryID)=" & SaleCategoryIDValue & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            SaleCategoryValue = .Fields("SaleCategory").Value
            SaleDiscountPercentValue = .Fields("SaleDiscountPercent").Value
            CashValue = .Fields("Cash").Value
            CreditValue = .Fields("Credit").Value
            ChequeValue = .Fields("Cheque").Value
            CreditCardValue = .Fields("CreditCard").Value
            OutPatientValue = .Fields("OutPatient").Value
            InPatientValue = !InPatient
            If IsNull(!ProfitMargin) = False Then
                ProfitMarginValue = !ProfitMargin
            Else
                ProfitMarginValue = 0
            End If
            StaffValue = .Fields("Staff").Value
            UnitValue = .Fields("Unit").Value
            OtherValue = .Fields("Other").Value
        End If
        If .State = 1 Then .Close
    End With
End Property



Public Property Get SaleCategoryID() As Long
    SaleCategoryID = SaleCategoryIDValue
End Property

Public Property Get SaleCategory() As String
    SaleCategory = SaleCategoryValue
End Property

Public Property Get ProfitMargin() As Double
    ProfitMargin = ProfitMarginValue
End Property

Public Property Get SaleDiscountPercent() As Double
    SaleDiscountPercent = SaleDiscountPercentValue
End Property

Public Property Get Cash() As Boolean
    Cash = CashValue
End Property

Public Property Get Credit() As Boolean
    Credit = CreditValue
End Property

Public Property Get Cheque() As Boolean
    Cheque = ChequeValue
End Property

Public Property Get CreditCard() As Boolean
    CreditCard = CreditCardValue
End Property

Public Property Get Other() As Boolean
    Other = OtherValue
End Property

Public Property Get InPatient() As Boolean
    InPatient = InPatientValue
End Property

Public Property Get OutPatient() As Boolean
    OutPatient = OutPatientValue
End Property

Public Property Get Staff() As Boolean
    Staff = StaffValue
End Property

Public Property Get Unit() As Boolean
    Unit = UnitValue
End Property
