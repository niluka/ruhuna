VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Private varItemID
    Private varVTM As String
    Private varVMP As String
    Private varAMP As String
    Private varAMPP As String
    Private varVMPP As String
    Private varDisplay As String
    Private varCode As String
    Private varTradeNameID
    Private varGenericNameID
    Private varItemCategoryID
    Private varStrengthUnitID
    Private varIssueUnitID
    Private varPackUnitID
    Private varStrengthOfIssueUnit As Double
    Private varIssueUnitsPerPack As Double
    Private varROL As Double
    Private varROQ As Double
    Private varMinQty As Double
    Private varManufacturerID
    Private varImporterID
    Private varComments As String
    Private varupsize_ts

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    Dim newEntry As Boolean
    With rsTem
        temSQL = "SELECT * FROM tblItem Where ItemID = " & varItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            newEntry = True
        Else
            newEntry = False
        End If
        !VTM = varVTM
        !VMP = varVMP
        !AMP = varAMP
        !AMPP = varAMPP
        !VMPP = varVMPP
        !Display = varDisplay
        !Code = varCode
        !TradeNameID = varTradeNameID
        !GenericNameID = varGenericNameID
        !ItemCategoryID = varItemCategoryID
        !StrengthUnitID = varStrengthUnitID
        !IssueUnitID = varIssueUnitID
        !PackUnitID = varPackUnitID
        !StrengthOfIssueUnit = varStrengthOfIssueUnit
        !IssueUnitsPerPack = varIssueUnitsPerPack
        !ROL = varROL
        !ROQ = varROQ
        !MinQty = varMinQty
        !ManufacturerID = varManufacturerID
        !ImporterID = varImporterID
        !Comments = varComments
        .Update
        If newEntry = True Then
            .Close
            temSQL = "SELECT @@IDENTITY AS NewID"
           .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            varItemID = !NewID
        Else
            varItemID = !ItemID
        End If
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblItem WHERE ItemID = " & varItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!ItemID) Then
               varItemID = !ItemID
            End If
            If Not IsNull(!VTM) Then
               varVTM = !VTM
            End If
            If Not IsNull(!VMP) Then
               varVMP = !VMP
            End If
            If Not IsNull(!AMP) Then
               varAMP = !AMP
            End If
            If Not IsNull(!AMPP) Then
               varAMPP = !AMPP
            End If
            If Not IsNull(!VMPP) Then
               varVMPP = !VMPP
            End If
            If Not IsNull(!Display) Then
               varDisplay = !Display
            End If
            If Not IsNull(!Code) Then
               varCode = !Code
            End If
            If Not IsNull(!TradeNameID) Then
               varTradeNameID = !TradeNameID
            End If
            If Not IsNull(!GenericNameID) Then
               varGenericNameID = !GenericNameID
            End If
            If Not IsNull(!ItemCategoryID) Then
               varItemCategoryID = !ItemCategoryID
            End If
            If Not IsNull(!StrengthUnitID) Then
               varStrengthUnitID = !StrengthUnitID
            End If
            If Not IsNull(!IssueUnitID) Then
               varIssueUnitID = !IssueUnitID
            End If
            If Not IsNull(!PackUnitID) Then
               varPackUnitID = !PackUnitID
            End If
            If Not IsNull(!StrengthOfIssueUnit) Then
               varStrengthOfIssueUnit = !StrengthOfIssueUnit
            End If
            If Not IsNull(!IssueUnitsPerPack) Then
               varIssueUnitsPerPack = !IssueUnitsPerPack
            End If
            If Not IsNull(!ROL) Then
               varROL = !ROL
            End If
            If Not IsNull(!ROQ) Then
               varROQ = !ROQ
            End If
            If Not IsNull(!MinQty) Then
               varMinQty = !MinQty
            End If
            If Not IsNull(!ManufacturerID) Then
               varManufacturerID = !ManufacturerID
            End If
            If Not IsNull(!ImporterID) Then
               varImporterID = !ImporterID
            End If
            If Not IsNull(!Comments) Then
               varComments = !Comments
            End If
            If Not IsNull(!upsize_ts) Then
               varupsize_ts = !upsize_ts
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varItemID = Empty
    varVTM = Empty
    varVMP = Empty
    varAMP = Empty
    varAMPP = Empty
    varVMPP = Empty
    varDisplay = Empty
    varCode = Empty
    varTradeNameID = Empty
    varGenericNameID = Empty
    varItemCategoryID = Empty
    varStrengthUnitID = Empty
    varIssueUnitID = Empty
    varPackUnitID = Empty
    varStrengthOfIssueUnit = 0
    varIssueUnitsPerPack = 0
    varROL = 0
    varROQ = 0
    varMinQty = 0
    varManufacturerID = Empty
    varImporterID = Empty
    varComments = Empty
    varupsize_ts = Empty
End Sub

Public Property Let ItemID(ByVal vItemID)
    Call clearData
    varItemID = vItemID
    Call loadData
End Property

Public Property Get ItemID()
    ItemID = varItemID
End Property

Public Property Let VTM(ByVal vVTM As String)
    varVTM = vVTM
End Property

Public Property Get VTM() As String
    VTM = varVTM
End Property

Public Property Let VMP(ByVal vVMP As String)
    varVMP = vVMP
End Property

Public Property Get VMP() As String
    VMP = varVMP
End Property

Public Property Let AMP(ByVal vAMP As String)
    varAMP = vAMP
End Property

Public Property Get AMP() As String
    AMP = varAMP
End Property

Public Property Let AMPP(ByVal vAMPP As String)
    varAMPP = vAMPP
End Property

Public Property Get AMPP() As String
    AMPP = varAMPP
End Property

Public Property Let VMPP(ByVal vVMPP As String)
    varVMPP = vVMPP
End Property

Public Property Get VMPP() As String
    VMPP = varVMPP
End Property

Public Property Let Display(ByVal vDisplay As String)
    varDisplay = vDisplay
End Property

Public Property Get Display() As String
    Display = varDisplay
End Property

Public Property Let Code(ByVal vCode As String)
    varCode = vCode
End Property

Public Property Get Code() As String
    Code = varCode
End Property

Public Property Let TradeNameID(ByVal vTradeNameID)
    varTradeNameID = vTradeNameID
End Property

Public Property Get TradeNameID()
    TradeNameID = varTradeNameID
End Property

Public Property Let GenericNameID(ByVal vGenericNameID)
    varGenericNameID = vGenericNameID
End Property

Public Property Get GenericNameID()
    GenericNameID = varGenericNameID
End Property

Public Property Let ItemCategoryID(ByVal vItemCategoryID)
    varItemCategoryID = vItemCategoryID
End Property

Public Property Get ItemCategoryID()
    ItemCategoryID = varItemCategoryID
End Property

Public Property Let StrengthUnitID(ByVal vStrengthUnitID)
    varStrengthUnitID = vStrengthUnitID
End Property

Public Property Get StrengthUnitID()
    StrengthUnitID = varStrengthUnitID
End Property

Public Property Let IssueUnitID(ByVal vIssueUnitID)
    varIssueUnitID = vIssueUnitID
End Property

Public Property Get IssueUnitID()
    IssueUnitID = varIssueUnitID
End Property

Public Property Let PackUnitID(ByVal vPackUnitID)
    varPackUnitID = vPackUnitID
End Property

Public Property Get PackUnitID()
    PackUnitID = varPackUnitID
End Property

Public Property Let StrengthOfIssueUnit(ByVal vStrengthOfIssueUnit As Double)
    varStrengthOfIssueUnit = vStrengthOfIssueUnit
End Property

Public Property Get StrengthOfIssueUnit() As Double
    StrengthOfIssueUnit = varStrengthOfIssueUnit
End Property

Public Property Let IssueUnitsPerPack(ByVal vIssueUnitsPerPack As Double)
    varIssueUnitsPerPack = vIssueUnitsPerPack
End Property

Public Property Get IssueUnitsPerPack() As Double
    IssueUnitsPerPack = varIssueUnitsPerPack
End Property

Public Property Let ROL(ByVal vROL As Double)
    varROL = vROL
End Property

Public Property Get ROL() As Double
    ROL = varROL
End Property

Public Property Let ROQ(ByVal vROQ As Double)
    varROQ = vROQ
End Property

Public Property Get ROQ() As Double
    ROQ = varROQ
End Property

Public Property Let MinQty(ByVal vMinQty As Double)
    varMinQty = vMinQty
End Property

Public Property Get MinQty() As Double
    MinQty = varMinQty
End Property

Public Property Let ManufacturerID(ByVal vManufacturerID)
    varManufacturerID = vManufacturerID
End Property

Public Property Get ManufacturerID()
    ManufacturerID = varManufacturerID
End Property

Public Property Let ImporterID(ByVal vImporterID)
    varImporterID = vImporterID
End Property

Public Property Get ImporterID()
    ImporterID = varImporterID
End Property

Public Property Let Comments(ByVal vComments As String)
    varComments = vComments
End Property

Public Property Get Comments() As String
    Comments = varComments
End Property

Public Property Let upsize_ts(ByVal vupsize_ts)
    varupsize_ts = vupsize_ts
End Property

Public Property Get upsize_ts()
    upsize_ts = varupsize_ts
End Property


