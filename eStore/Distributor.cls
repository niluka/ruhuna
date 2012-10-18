VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Distributor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Private varDistributorID
    Private varDistributorName As String
    Private varDistributorAddress As String
    Private varDistributorCityID As Long
    Private varDistributorCountryID As Long
    Private varDistributorTelephone As String
    Private varDistributorFax As String
    Private varDistributorEmail As String
    Private varDistributorWebsite As String
    Private varDistributorComments As String
    Private varbalance As Double
    Private varupsize_ts

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    Dim newEntry As Boolean
    With rsTem
        temSQL = "SELECT * FROM tblDistrubutor Where DistributorID = " & varDistributorID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            newEntry = True
        Else
            newEntry = False
        End If
        !DistributorName = varDistributorName
        !DistributorAddress = varDistributorAddress
        !DistributorCityID = varDistributorCityID
        !DistributorCountryID = varDistributorCountryID
        !DistributorTelephone = varDistributorTelephone
        !DistributorFax = varDistributorFax
        !DistributorEmail = varDistributorEmail
        !DistributorWebsite = varDistributorWebsite
        !DistributorComments = varDistributorComments
        !balance = varbalance
        .Update
        If newEntry = True Then
            .Close
            temSQL = "SELECT @@IDENTITY AS NewID"
           .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            varDistributorID = !NewID
        Else
            varDistributorID = !DistributorID
        End If
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblDistrubutor WHERE DistributorID = " & varDistributorID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!DistributorID) Then
               varDistributorID = !DistributorID
            End If
            If Not IsNull(!DistributorName) Then
               varDistributorName = !DistributorName
            End If
            If Not IsNull(!DistributorAddress) Then
               varDistributorAddress = !DistributorAddress
            End If
            If Not IsNull(!DistributorCityID) Then
               varDistributorCityID = !DistributorCityID
            End If
            If Not IsNull(!DistributorCountryID) Then
               varDistributorCountryID = !DistributorCountryID
            End If
            If Not IsNull(!DistributorTelephone) Then
               varDistributorTelephone = !DistributorTelephone
            End If
            If Not IsNull(!DistributorFax) Then
               varDistributorFax = !DistributorFax
            End If
            If Not IsNull(!DistributorEmail) Then
               varDistributorEmail = !DistributorEmail
            End If
            If Not IsNull(!DistributorWebsite) Then
               varDistributorWebsite = !DistributorWebsite
            End If
            If Not IsNull(!DistributorComments) Then
               varDistributorComments = !DistributorComments
            End If
            If Not IsNull(!balance) Then
               varbalance = !balance
            End If
            If Not IsNull(!upsize_ts) Then
               varupsize_ts = !upsize_ts
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varDistributorID = Empty
    varDistributorName = Empty
    varDistributorAddress = Empty
    varDistributorCityID = 0
    varDistributorCountryID = 0
    varDistributorTelephone = Empty
    varDistributorFax = Empty
    varDistributorEmail = Empty
    varDistributorWebsite = Empty
    varDistributorComments = Empty
    varbalance = 0
    varupsize_ts = Empty
End Sub

Public Property Let DistributorID(ByVal vDistributorID)
    Call clearData
    varDistributorID = vDistributorID
    Call loadData
End Property

Public Property Get DistributorID()
    DistributorID = varDistributorID
End Property

Public Property Let DistributorName(ByVal vDistributorName As String)
    varDistributorName = vDistributorName
End Property

Public Property Get DistributorName() As String
    DistributorName = varDistributorName
End Property

Public Property Let DistributorAddress(ByVal vDistributorAddress As String)
    varDistributorAddress = vDistributorAddress
End Property

Public Property Get DistributorAddress() As String
    DistributorAddress = varDistributorAddress
End Property

Public Property Let DistributorCityID(ByVal vDistributorCityID As Long)
    varDistributorCityID = vDistributorCityID
End Property

Public Property Get DistributorCityID() As Long
    DistributorCityID = varDistributorCityID
End Property

Public Property Let DistributorCountryID(ByVal vDistributorCountryID As Long)
    varDistributorCountryID = vDistributorCountryID
End Property

Public Property Get DistributorCountryID() As Long
    DistributorCountryID = varDistributorCountryID
End Property

Public Property Let DistributorTelephone(ByVal vDistributorTelephone As String)
    varDistributorTelephone = vDistributorTelephone
End Property

Public Property Get DistributorTelephone() As String
    DistributorTelephone = varDistributorTelephone
End Property

Public Property Let DistributorFax(ByVal vDistributorFax As String)
    varDistributorFax = vDistributorFax
End Property

Public Property Get DistributorFax() As String
    DistributorFax = varDistributorFax
End Property

Public Property Let DistributorEmail(ByVal vDistributorEmail As String)
    varDistributorEmail = vDistributorEmail
End Property

Public Property Get DistributorEmail() As String
    DistributorEmail = varDistributorEmail
End Property

Public Property Let DistributorWebsite(ByVal vDistributorWebsite As String)
    varDistributorWebsite = vDistributorWebsite
End Property

Public Property Get DistributorWebsite() As String
    DistributorWebsite = varDistributorWebsite
End Property

Public Property Let DistributorComments(ByVal vDistributorComments As String)
    varDistributorComments = vDistributorComments
End Property

Public Property Get DistributorComments() As String
    DistributorComments = varDistributorComments
End Property

Public Property Let balance(ByVal vbalance As Double)
    varbalance = vbalance
End Property

Public Property Get balance() As Double
    balance = varbalance
End Property

Public Property Let upsize_ts(ByVal vupsize_ts)
    varupsize_ts = vupsize_ts
End Property

Public Property Get upsize_ts()
    upsize_ts = varupsize_ts
End Property


