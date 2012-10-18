VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim IDValue As Long
    Dim SaleCategoryValue As String
    Dim SaleCategoryIDValue As Long
    Dim SatffValue As String
    Dim StaffIDValue As Long
    Dim StaffRateValue As Double
    Dim HospitalRateValue As Double
    Dim TotalRateValue As Double

Public Property Get ID() As Long

End Property

Public Property Get CategoryID() As Long

End Property

Public Property Let ID(ID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleCategory.Category, tblStaff.Staff, tblItem.* FROM (tblItem LEFT JOIN tblStaff ON tblItem.StaffID = tblStaff.StaffID) LEFT JOIN tblSaleCategory ON tblItem.SaleCategoryID = tblSaleCategory.CategoryID Where ItemID =" & ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            SaleCategoryIDValue = !SaleCategoryID
            SaleCategoryValue = !SaleCategory
            SatffValue = !Staff
            StaffIDValue = !StaffID
            StaffRateValue = !StaffRate
            HospitalRateValue = !HospitalRate
            TotalRateValue = !TotalRate
        End If
        .Close
    End With
End Property
 
 
