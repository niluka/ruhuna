VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatientCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim CategoryV As String
    Dim IDV As Long
    Dim PaymentMethodV As String
    Dim PaymentMethodIDV As Long
    Dim SurchargeV As Double
    Dim IndoorPatientV As Boolean
    Dim OutdoorPatientV As Boolean
    Dim temSql As String
    
Public Property Let ID(IDGiven As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientCategory.PatientCategoryID, tblPatientCategory.PatientCategory, tblPaymentMethod.PaymentMethod, tblPatientCategory.PaymentMethodID, tblPatientCategory.Surcharge, tblPatientCategory.IndoorPatient, tblPatientCategory.OutdoorPatient " & _
                    "FROM tblPatientCategory LEFT JOIN tblPaymentMethod ON tblPatientCategory.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblPatientCategory.PatientCategoryID)=" & IDGiven & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        IDV = IDGiven
        If .RecordCount > 0 Then
            CategoryV = !PatientCategory
            IDV = !PatientCategoryID
            PaymentMethodV = !PaymentMethod
            PaymentMethodIDV = !PaymentMethodID
            SurchargeV = !Surcharge
            IndoorPatientV = !IndoorPatient
            OutdoorPatientV = !OutdoorPatient
        End If
        .Close
    End With
End Property



Public Property Get ID() As Long
    ID = IDV
End Property

Public Property Get Category() As String
    Category = CategoryV
End Property

Public Property Get PaymentMethod() As String
    PaymentMethod = PaymentMethodV
End Property

Public Property Get PaymentMethodID() As Long
    PaymentMethodID = PaymentMethodIDV
End Property

Public Property Get Surcharge() As Double
    Surcharge = SurchargeV
End Property

Public Property Get IndoorPatient() As Boolean
    IndoorPatient = IndoorPatientV
End Property

Public Property Get OutdoorPatient() As Boolean
    OutdoorPatient = OutdoorPatientV
End Property
