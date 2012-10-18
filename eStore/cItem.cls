VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    
    Private ItemIDValue As Long
    Private ItemDisplayValue As String
    Private ItemCodeValue As String
    Private ItemCategoryValue As String
    Private ItemCategoryIDValue As Long
    Private ItemGenericValue As String
    Private ItemGenericIDValue As Long
    Private ItemTradeValue As String
    Private ItemTradeValueID As Long
    Private ItemManufacturerNameValue As String
    Private ItemManufacturerIDValue As Long
    Private ItemImporterNameValue As String
    Private ItemImporterIDValue As String
    Private ItemStrengthUnitValue As String
    Private ItemStrengthUnitIDValue As Double
    Private ItemIUnitValue As String
    Private ItemIUnitIDValue As Long
    Private ItemPUnitValue As String
    Private ItemPUnitIDValue As String
    Private ItemStrengthOfIssueUnitValue As Double
    Private ItemIssueUnitsPerPackValue As Double
    Private ItemROLValue As Double
    Private ItemROQValue As Double
    Private ItemMinQtyValue As Double
    Private ItemAMPValue As String
    Private ItemVMPValue As String
    Private ItemAMPPValue As String
    Private ItemVMPPValue As String
    Private ItemProfitMargine As Double
    Private ItemPPriceValue As Double
    Private ItemSPriceValue As Double
    
    
    Private temSql As String

Private Sub ClearValues()
            ItemDisplayValue = Empty
            ItemCodeValue = Empty
            ItemCategoryValue = Empty
            ItemCategoryIDValue = Empty
            ItemGenericValue = Empty
            ItemGenericIDValue = Empty
            ItemTradeValue = Empty
            ItemTradeValueID = Empty
            ItemManufacturerNameValue = Empty
            ItemManufacturerIDValue = Empty
            ItemImporterNameValue = Empty
            ItemImporterIDValue = Empty
            ItemROLValue = Empty
            ItemROQValue = Empty
            ItemMinQtyValue = Empty
            ItemStrengthUnitValue = Empty
            ItemStrengthUnitIDValue = Empty
            ItemIUnitValue = Empty
            ItemIUnitIDValue = Empty
            ItemPUnitValue = Empty
            ItemPUnitIDValue = Empty
            ItemAMPValue = Empty
            ItemVMPValue = Empty
            ItemVMPPValue = Empty
            ItemAMPPValue = Empty
            ItemStrengthOfIssueUnitValue = Empty
            ItemIssueUnitsPerPackValue = Empty
            ItemIssueUnitsPerPackValue = 1
            ItemProfitMargine = Empty
            ItemPPriceValue = Empty
            ItemSPriceValue = Empty

End Sub

Public Property Let ID(IDValue As Long)
    Call ClearValues
    ItemIDValue = IDValue
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.*, tblGenericName.GenericName, tblImporter.ImporterName, tblItemCategory.ItemCategory, tblItemCategory.SalesMargin, tblManufacturer.ManufacturerName, tblPackUnit.PackUnit, tblStrengthUnit.StrengthUnit, tblTradeName.TradeName, tblIssueUnit.IssueUnit " & _
                    " FROM tblStrengthUnit RIGHT JOIN (((tblManufacturer RIGHT JOIN (tblPackUnit RIGHT JOIN ((tblGenericName RIGHT JOIN (tblTradeName RIGHT JOIN tblItem ON tblTradeName.TradeNameID = tblItem.TradeNameID) ON tblGenericName.GenericNameID = tblItem.GenericNameID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) ON tblPackUnit.PackUnitID = tblItem.PackUnitID) ON tblManufacturer.ManufacturerID = tblItem.ManufacturerID) LEFT JOIN tblImporter ON tblItem.ImporterID = tblImporter.ImporterID) LEFT JOIN tblItemCategory ON tblItem.ItemCategoryID = tblItemCategory.ItemCategoryID) ON tblStrengthUnit.StrengthUnitID = tblItem.StrengthUnitID " & _
                    " WHERE (((tblItem.ItemID)=" & ItemIDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNumeric(!Display) Then ItemDisplayValue = .Fields("Display").Value
            If Not IsNull(!Code) Then ItemCodeValue = .Fields("code").Value
            If Not IsNull(!ItemCategory) Then ItemCategoryValue = .Fields("ItemCategory").Value
            If Not IsNull(!ItemCategoryID) Then ItemCategoryIDValue = .Fields("ItemCategoryID").Value
            If Not IsNull(!GenericName) Then ItemGenericValue = .Fields("GenericName").Value
            If Not IsNull(!GenericNameID) Then ItemGenericIDValue = .Fields("GenericNameID").Value
            If Not IsNull(!TradeName) Then ItemTradeValue = .Fields("TradeName").Value
            If Not IsNull(!TradeNameID) Then ItemTradeValueID = .Fields("TradeNameID").Value
            If Not IsNull(!ManufacturerName) Then ItemManufacturerNameValue = .Fields("ManufacturerName").Value
            If Not IsNull(!ManufacturerID) Then ItemManufacturerIDValue = .Fields("ManufacturerID").Value
            If Not IsNull(!ImporterName) Then ItemImporterNameValue = .Fields("ImporterName").Value
            If Not IsNull(!ImporterID) Then ItemImporterIDValue = .Fields("ImporterID").Value
            If Not IsNull(!ROL) Then ItemROLValue = .Fields("ROL").Value
            If Not IsNull(!ROQ) Then ItemROQValue = .Fields("ROQ").Value
            If Not IsNull(!MinQty) Then ItemMinQtyValue = .Fields("MinQty").Value
            If Not IsNull(!StrengthUnit) Then ItemStrengthUnitValue = .Fields("StrengthUnit").Value
            If Not IsNull(!StrengthUnitID) Then ItemStrengthUnitIDValue = .Fields("StrengthUnitID").Value
            If Not IsNull(!IssueUnit) Then ItemIUnitValue = .Fields("IssueUnit").Value
            If Not IsNull(!IssueUnitID) Then ItemIUnitIDValue = .Fields("IssueUnitID").Value
            If Not IsNull(!PackUnit) Then ItemPUnitValue = .Fields("PackUnit").Value
            If Not IsNull(!PackUnitID) Then ItemPUnitIDValue = .Fields("PackUnitID").Value
            If Not IsNull(!AMP) Then ItemAMPValue = .Fields("AMP").Value
            If Not IsNull(!VMP) Then ItemVMPValue = .Fields("VMP").Value
            If Not IsNull(!VMPP) Then ItemVMPPValue = .Fields("VMPP").Value
            If Not IsNull(!AMPP) Then ItemAMPPValue = .Fields("AMPP").Value
            If Not IsNull(!StrengthOfIssueUnit) Then ItemStrengthOfIssueUnitValue = .Fields("StrengthOfIssueUnit").Value
            If ItemStrengthOfIssueUnitValue = 0 Then ItemStrengthOfIssueUnitValue = 1
            If Not IsNull(!IssueUnitsPerPack) Then ItemIssueUnitsPerPackValue = .Fields("IssueUnitsPerPack").Value
            If ItemIssueUnitsPerPackValue = 0 Then ItemIssueUnitsPerPackValue = 1
            If Not IsNull(!SalesMargin) Then ItemProfitMargine = .Fields("SalesMargin").Value
        End If
        If .State = 1 Then .Close
    End With
    With rsPPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblCurrentPurchasePrice.PPrice " & _
                    "From tblCurrentPurchasePrice " & _
                    "Where (((tblCurrentPurchasePrice.ItemID) = " & ItemIDValue & ")) " & _
                    " Order By SetDate Desc, SetTime DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockBatchOptimistic
        If .RecordCount > 0 Then
            ItemPPriceValue = !PPrice
        End If
        .Close
    End With
    With rsSPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblCurrentSalePrice.SPrice " & _
                    "From tblCurrentSalePrice " & _
                    "Where (((tblCurrentSalePrice.ItemID) = " & ItemIDValue & ")) " & _
                    " Order By SetDate Desc, SetTime DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockBatchOptimistic
        If .RecordCount > 0 Then
            ItemSPriceValue = !SPrice
        End If
        .Close
    End With
    
End Property

Public Property Get PPrice() As Double
    PPrice = ItemPPriceValue
End Property

Public Property Get SPrice() As Double
    SPrice = ItemSPriceValue
End Property

Public Property Get ID() As Long
    ID = ItemIDValue
End Property
Public Property Get Display() As String
    Display = ItemDisplayValue
End Property
Public Property Get Code() As String
    Code = ItemCodeValue
End Property
Public Property Get Category() As String
    Category = ItemCategoryValue
End Property
Public Property Get CategoryID() As Long
    CategoryID = ItemCategoryIDValue
End Property
Public Property Get Generic() As String
    Generic = ItemGenericValue
End Property
Public Property Get GenericID() As Long
    GenericID = ItemGenericIDValue
End Property
Public Property Get Trade() As String
    Trade = ItemTradeValue
End Property
Public Property Get TradeID() As Long
    TradeID = ItemTradeValueID
End Property
Public Property Get ManufacturerName() As String
    ManufacturerName = ItemManufacturerNameValue
End Property
Public Property Get ManufacturerID() As String
    ManufacturerID = ItemManufacturerIDValue
End Property
Public Property Get ImporterName() As String
    ImporterName = ItemImporterNameValue
End Property
Public Property Get ImporterID() As String
    ImporterID = ItemImporterIDValue
End Property
Public Property Get ROL() As Double
    ROL = ItemROLValue
End Property
Public Property Get ROQ() As Double
    ROQ = ItemROQValue
End Property
Public Property Get MinQty() As Double
    MinQty = ItemMinQtyValue
End Property
Public Property Get StrengthUnit() As String
    StrengthUnit = ItemStrengthUnitValue
End Property
Public Property Get StrengthUnitID() As Double
    StrengthUnitID = ItemStrengthUnitIDValue
End Property
Public Property Get IUnit() As String
    IUnit = ItemIUnitValue
End Property
Public Property Get IUnitID() As Long
    IUnitID = ItemIUnitIDValue
End Property
Public Property Get PUnit() As String
    PUnit = ItemPUnitValue
End Property
Public Property Get PUnitID() As String
    PUnitID = ItemPUnitIDValue
End Property
Public Property Get StrengthOfIssueUnit()
    StrengthOfIssueUnit = ItemStrengthOfIssueUnitValue
End Property
Public Property Get IssueUnitsPerPack()
    IssueUnitsPerPack = ItemIssueUnitsPerPackValue
End Property
Public Property Get AMP() As String
    AMP = ItemAMPValue
End Property
Public Property Get VMP() As String
    VMP = ItemVMPValue
End Property
Public Property Get AMPP() As String
    AMPP = ItemAMPPValue
End Property
Public Property Get VMPP() As String
    VMPP = ItemVMPPValue
End Property
Public Property Get SalesMargin() As Double
    SalesMargin = ItemProfitMargine
End Property

