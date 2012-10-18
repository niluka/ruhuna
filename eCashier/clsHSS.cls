VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim HealthSchemeSupplierV As String
    Dim HealthSchemeSupplierIDV As Long
    Dim AddressV As String

Public Property Let HSSID(ID As Long)
    Dim rsTem As New ADODB.Recordset
    HealthSchemeSupplierV = 0
    HealthSchemeSupplierIDV = 0
    AddressV = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblHealthSchemeSuppliers.* " & _
                    "FROM tblHealthSchemeSuppliers " & _
                    "WHERE (((tblHealthSchemeSuppliers.HealthSchemeSupplierID)=" & ID & "))"
        HealthSchemeSupplierIDV = ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            HealthSchemeSupplierV = !HealthSchemeSupplierName
            AddressV = !HealthSchemeSupplierAddress
        End If
        .Close
    End With
End Property

Public Property Get HSSID() As Long
    HSSID = HealthSchemeSupplierIDV
End Property

Public Property Get Name() As String
    Name = HealthSchemeSupplierV
End Property

Public Property Get Address() As String
    Address = AddressV
End Property
