VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Private varItemCategoryID
    Private varItemCategory As String
    Private varComments As String
    Private varSalesMargin As Double
    Private varCategoryCode
    Private varupsize_ts

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    Dim newEntry As Boolean
    With rsTem
        temSQL = "SELECT * FROM tblItemCategory Where ItemCategoryID = " & varItemCategoryID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            newEntry = True
        Else
            newEntry = False
        End If
        !ItemCategory = varItemCategory
        !Comments = varComments
        !SalesMargin = varSalesMargin
        !CategoryCode = varCategoryCode
        .Update
        If newEntry = True Then
            .Close
            temSQL = "SELECT @@IDENTITY AS NewID"
           .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            varItemCategoryID = !NewID
        Else
            varItemCategoryID = !ItemCategoryID
        End If
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblItemCategory WHERE ItemCategoryID = " & varItemCategoryID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!ItemCategoryID) Then
               varItemCategoryID = !ItemCategoryID
            End If
            If Not IsNull(!ItemCategory) Then
               varItemCategory = !ItemCategory
            End If
            If Not IsNull(!Comments) Then
               varComments = !Comments
            End If
            If Not IsNull(!SalesMargin) Then
               varSalesMargin = !SalesMargin
            End If
            If Not IsNull(!CategoryCode) Then
               varCategoryCode = !CategoryCode
            End If
            If Not IsNull(!upsize_ts) Then
               varupsize_ts = !upsize_ts
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varItemCategoryID = Empty
    varItemCategory = Empty
    varComments = Empty
    varSalesMargin = 0
    varCategoryCode = Empty
    varupsize_ts = Empty
End Sub

Public Property Let ItemCategoryID(ByVal vItemCategoryID)
    Call clearData
    varItemCategoryID = vItemCategoryID
    Call loadData
End Property

Public Property Get ItemCategoryID()
    ItemCategoryID = varItemCategoryID
End Property

Public Property Let ItemCategory(ByVal vItemCategory As String)
    varItemCategory = vItemCategory
End Property

Public Property Get ItemCategory() As String
    ItemCategory = varItemCategory
End Property

Public Property Let Comments(ByVal vComments As String)
    varComments = vComments
End Property

Public Property Get Comments() As String
    Comments = varComments
End Property

Public Property Let SalesMargin(ByVal vSalesMargin As Double)
    varSalesMargin = vSalesMargin
End Property

Public Property Get SalesMargin() As Double
    SalesMargin = varSalesMargin
End Property

Public Property Let CategoryCode(ByVal vCategoryCode)
    varCategoryCode = vCategoryCode
End Property

Public Property Get CategoryCode()
    CategoryCode = varCategoryCode
End Property

Public Property Let upsize_ts(ByVal vupsize_ts)
    varupsize_ts = vupsize_ts
End Property

Public Property Get upsize_ts()
    upsize_ts = varupsize_ts
End Property


