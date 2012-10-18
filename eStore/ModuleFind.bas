Attribute VB_Name = "ModuleFind"
Option Explicit
    Public Type ItemSaleAndReturn
        SaleValue As Double
        SaleQuentity As Double
        ReturnQuentity As Double
        ReturnValue As Double
    End Type

    Private rsTem1 As New ADODB.Recordset
    Private rsTem2 As New ADODB.Recordset
    Private rsTem3 As New ADODB.Recordset

Private temSql As String
Public Type Stock
    Amount As Double
    MaxDOE As Date
    MinDOE As Date
End Type


Public Function CalculateConsumption(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long, Optional ByVal StaffID As Long, Optional ByVal CategoryID As Long) As Double
    With rsTem1
        temSql = "SELECT Sum(Amount) AS Total FROM tblConsumption WHERE ItemID=" & ItemID & " AND Date Between '" & Format(FromDate, "MMMM dd YYYY") & "' And '" & Format(ToDate, "MMMM dd yyyy") & "' "
        If StoreID <> 0 Then temSql = temSql & " AND StoreID=" & StoreID & " "
        If StaffID <> 0 Then temSql = temSql & " AND StaffID =" & StaffID & " "
        If CategoryID <> 0 Then temSql = temSql & " AND CategoryID =" & CategoryID & " "
        If BatchID <> 0 Then temSql = temSql & " AND BatchID=" & BatchID & " "
        If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If Not IsNull(!Total) Then
                CalculateConsumption = !Total
            Else
                CalculateConsumption = 0
            End If
        .Close
    End With
End Function

Public Function CalculateSale(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long, Optional ByVal StaffID As Long, Optional ByVal CategoryID As Long) As Double
    With rsTem1
        temSql = "SELECT Sum(Amount) AS Total FROM tblSale WHERE ItemID=" & ItemID & " AND Date Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' "
        If StoreID <> 0 Then temSql = temSql & " AND StoreID=" & StoreID & " "
        If StaffID <> 0 Then temSql = temSql & " AND StaffID =" & StaffID & " "
        If CategoryID <> 0 Then temSql = temSql & " AND CategoryID =" & CategoryID & " "
        If BatchID <> 0 Then temSql = temSql & " AND BatchID=" & BatchID & " "
        If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If Not IsNull(!Total) Then
                CalculateSale = !Total
            Else
                CalculateSale = 0
            End If
        .Close
    End With
End Function

Public Function CalculateSalePrice(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long, Optional ByVal StaffID As Long, Optional ByVal CategoryID As Long) As Double
    With rsTem1
        temSql = "SELECT Sum(Price) AS Total FROM tblSale WHERE ItemID=" & ItemID & " AND Date Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' "
        If StoreID <> 0 Then temSql = temSql & " AND StoreID=" & StoreID & " "
        If StaffID <> 0 Then temSql = temSql & " AND StaffID =" & StaffID & " "
        If CategoryID <> 0 Then temSql = temSql & " AND CategoryID =" & CategoryID & " "
        If BatchID <> 0 Then temSql = temSql & " AND BatchID=" & BatchID & " "
        If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If Not IsNull(!Total) Then
                CalculateSalePrice = !Total
            Else
                CalculateSalePrice = 0
            End If
        .Close
    End With
End Function

Public Function CalculatePurchase(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date) As Double
    CalculatePurchase = 0
    With rsTem1
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblRefill.Amount) AS SumOfAmount, Sum(tblRefill.FreeAmount) AS SumOfFreeAmount, Sum(tblRefill.Price) AS SumOfPrice " & _
                    "FROM tblRefill RIGHT JOIN tblRefillBill ON tblRefill.RefillBillID = tblRefillBill.RefillBillID " & _
                    "WHERE (((tblRefill.ItemID)=" & ItemID & ") AND ((tblRefillBill.Date)Between '" & Format(FromDate, "DD MMMM yyyy") & "' And '" & Format(ToDate, "DD MMMM yyyy") & "'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfAmount) = False Then
                CalculatePurchase = !SumOfAmount
            End If
            If IsNull(!SumOfFreeAmount) = False Then
                CalculatePurchase = CalculatePurchase + !SumOfFreeAmount
            End If
        End If
        .Close
    End With
End Function

Public Function CalculatePurchaseValue(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date) As Double
    CalculatePurchaseValue = 0
    With rsTem1
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblRefill.Price) AS SumOfPrice " & _
                    "FROM tblRefill RIGHT JOIN tblRefillBill ON tblRefill.RefillBillID = tblRefillBill.RefillBillID " & _
                    "WHERE (((tblRefill.ItemID)=" & ItemID & ") AND ((tblRefillBill.Date)Between '" & Format(FromDate, "DD MMMM yyyy") & "' And '" & Format(ToDate, "DD MMMM yyyy") & "'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfPrice) = False Then
                CalculatePurchaseValue = !SumOfPrice
            End If
        End If
        .Close
    End With
End Function


Public Function CalculateAdjustment(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long, Optional ByVal StaffID As Long, Optional ByVal CategoryID As Long) As Double
    With rsTem1
        temSql = "SELECT Sum(Amount) AS Total FROM tblAdjustment WHERE ItemID=" & ItemID & " AND Date Between '" & Format(FromDate, "MMMM dd yyyy") & "' And '" & Format(ToDate, "MMMM dd yyyy") & "' "
        If StoreID <> 0 Then temSql = temSql & " AND StoreID=" & StoreID & " "
        If StaffID <> 0 Then temSql = temSql & " AND StaffID =" & StaffID & " "
        If CategoryID <> 0 Then temSql = temSql & " AND CategoryID =" & CategoryID & " "
        If BatchID <> 0 Then temSql = temSql & " AND BatchID=" & BatchID & " "
        If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If Not IsNull(!Total) Then
                CalculateAdjustment = !Total
            Else
                CalculateAdjustment = 0
            End If
        .Close
    End With
End Function

Public Function CalculateDiscard(ByVal ItemID As Long, ByVal FromDate As Date, ByVal ToDate As Date, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long, Optional ByVal StaffID As Long, Optional ByVal CategoryID As Long) As Double
    With rsTem1
        temSql = "SELECT Sum(Amount) AS Total FROM tblDiscard WHERE ItemID=" & ItemID & " AND Date Between '" & Format(FromDate, "MMMM dd yyyy") & "' And '" & Format(ToDate, "MMMM dd yyyy") & "' "
        If StoreID <> 0 Then temSql = temSql & " AND StoreID=" & StoreID & " "
        If StaffID <> 0 Then temSql = temSql & " AND StaffID =" & StaffID & " "
        If CategoryID <> 0 Then temSql = temSql & " AND CategoryID =" & CategoryID & " "
        If BatchID <> 0 Then temSql = temSql & " AND BatchID=" & BatchID & " "
        If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If Not IsNull(!Total) Then
                CalculateDiscard = !Total
            Else
                CalculateDiscard = 0
            End If
        .Close
    End With
End Function

Public Function CalculateStock(ByVal ItemID As Long, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long) As Stock
    With rsTem1
        temSql = "SELECT Sum([tblBatchStock].[Stock]) AS TotalStock, Min([tblBatch].[DOE]) AS MinDOE, Max([tblBatch].[DOE]) AS MaxDOE FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " "
        If BatchID <> 0 Then temSql = temSql & " AND tblBatch.BatchID =" & BatchID & " "
        If StoreID <> 0 Then temSql = temSql & "AND tblBatchStock.StoreID=" & StoreID & " "
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If Not IsNull(!TotalStock) Then
            CalculateStock.Amount = !TotalStock
        Else
            CalculateStock.Amount = 0
        End If
        If Not IsNull(!MinDOE) Then
            CalculateStock.MinDOE = !MinDOE
        Else
            CalculateStock.MinDOE = TimeSerial(0, 0, 0)
        End If
        If Not IsNull(!MaxDOE) Then
            CalculateStock.MaxDOE = !MaxDOE
        Else
            CalculateStock.MaxDOE = TimeSerial(0, 0, 0)
        End If
    End With
End Function


Public Function AddToStock(BatchID As Long, StoreID As Long, Amount As Long) As Boolean
    AddToStock = False
    With rsTem1
        If .State = 1 Then .Close
        temSql = "SELECT tblBatchstock.* From tblBatchstock WHERE BatchID=" & BatchID & " AND StoreID = " & StoreID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Stock = !Stock + Amount
            .Update
            AddToStock = True
        Else
            .AddNew
            !BatchID = BatchID
            !StoreID = StoreID
            !Stock = Amount
            .Update
            AddToStock = True
        End If
        .Close
    End With
End Function

Public Function BatchExist(Batch As String, ItemID As Long) As Long
    BatchExist = 0
    With rsTem2
        If .State = 1 Then .Close
        temSql = "SELECT tblBatch.* From tblBatch WHERE tblBatch.ItemID=" & ItemID & " AND Batch = '" & Batch & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            BatchExist = !BatchID
        Else
            BatchExist = 0
        End If
        .Close
    End With
End Function

Public Function AddBatch(Batch As String, ItemID As Long, DOM As Date, DOE As Date) As Long
AddBatch = 0
    With rsTem3
        If .State = 1 Then .Close
        temSql = "SELECT tblBatch.* From tblBatch"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ItemID = ItemID
        !Batch = Batch
        !DOM = DOM
        !DOE = DOE
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        AddBatch = !NewID
        .Close
    End With
End Function


Public Function LastDateOfMonth(ByVal SuppliedDate As Date) As Date
    If Month(SuppliedDate) = 12 Then
        LastDateOfMonth = DateSerial(Year(SuppliedDate), 12, 31)
    Else
        LastDateOfMonth = DateSerial(Year(SuppliedDate), Month(SuppliedDate) + 1, 1) - 1
    End If
End Function

Public Function PeriodSale(FromDate As Date, ToDate As Date, ItemID As Long, Optional ForBHTID As Long, Optional ForOPDPtID As Long, Optional ForStaffID As Long, Optional ForUnitID As Long, Optional SaleCategoryID As Long) As ItemSaleAndReturn
    Dim rsTem As New ADODB.Recordset
    PeriodSale.SaleQuentity = 0
    PeriodSale.SaleValue = 0
    PeriodSale.ReturnQuentity = 0
    PeriodSale.SaleValue = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblItem.ItemID, SUM(dbo.tblSale.Price) AS SumOfPrice, SUM(dbo.tblSale.Amount) AS SumOfAmount "
        temSql = temSql & "FROM dbo.tblSale RIGHT OUTER JOIN dbo.tblSaleBill ON dbo.tblSale.SaleBillID = dbo.tblSaleBill.SaleBillID LEFT OUTER JOIN dbo.tblItem ON dbo.tblSale.ItemID = dbo.tblItem.ItemID "
        temSql = temSql & "WHERE (dbo.tblItem.ItemID = " & ItemID & ") "
        If ForBHTID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledBHTID = " & ForBHTID & ")  "
        End If
        If ForStaffID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledStaffID = " & ForStaffID & ") "
        End If
        If ForOPDPtID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledOutPatientID = " & ForOPDPtID & ") "
        End If
        If ForUnitID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledUnitID = " & ForUnitID & ") "
        End If
        If SaleCategoryID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.SaleCategoryID = " & SaleCategoryID & ") "
        End If
        temSql = temSql & "AND (dbo.tblSaleBill.Date BETWEEN CONVERT(DATETIME, '" & Format(FromDate, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(ToDate, "dd MMMM yyyy") & "', 102)) "
        temSql = temSql & "GROUP BY dbo.tblItem.ItemID "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfPrice) = False Then
                PeriodSale.SaleValue = !SumOfPrice
            End If
            If IsNull(!SumOfAmount) = False Then
                PeriodSale.SaleQuentity = !SumOfAmount
            End If
        End If
        .Close
        
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblItem.ItemID, SUM(dbo.tblReturn.ReturnAmount) AS SumOfRQuentity, SUM(dbo.tblReturn.ReturnPrice) AS SumOfRValue "
        temSql = temSql & "FROM dbo.tblReturn LEFT OUTER JOIN dbo.tblItem ON dbo.tblReturn.ItemID = dbo.tblItem.ItemID LEFT OUTER JOIN dbo.tblSaleBill ON dbo.tblReturn.SaleBillID = dbo.tblSaleBill.SaleBillID RIGHT OUTER JOIN dbo.tblReturnBill ON dbo.tblReturn.ReturnBillID = dbo.tblReturnBill.ReturnBillID "
        temSql = temSql & "WHERE (dbo.tblItem.ItemID = " & ItemID & ") "
        If ForBHTID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledBHTID = " & ForBHTID & ")  "
        End If
        If ForStaffID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledStaffID = " & ForStaffID & ") "
        End If
        If ForOPDPtID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledOutPatientID = " & ForOPDPtID & ") "
        End If
        If ForUnitID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.BilledUnitID = " & ForUnitID & ") "
        End If
        If SaleCategoryID <> 0 Then
            temSql = temSql & " AND (dbo.tblSaleBill.SaleCategoryID = " & SaleCategoryID & ") "
        End If
        temSql = temSql & "AND (dbo.tblReturnBill.Date BETWEEN CONVERT(DATETIME, '" & Format(FromDate, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(ToDate, "dd MMMM yyyy") & "', 102)) "
        temSql = temSql & "GROUP BY dbo.tblItem.ItemID "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfRValue) = False Then
                PeriodSale.ReturnValue = !SumOfRValue
            End If
            If IsNull(!SumOfRQuentity) = False Then
                PeriodSale.ReturnQuentity = !SumOfRQuentity
            End If
        End If
        .Close
        
    End With
End Function


Public Function CalculateAgeInWords(PatientDOB As Date) As String
    Dim Age As Long
    Age = DateDiff("yyyy", PatientDOB, Now)
    If Age >= 12 Then
        CalculateAgeInWords = Age & " Years"
        Exit Function
    Else
        Age = DateDiff("m", PatientDOB, Now)
        
        If Age > 144 Then CalculateAgeInWords = "12" & " Years and " & Age - 144 & " months": Exit Function
        If Age = 144 Then CalculateAgeInWords = "12" & " Years": Exit Function
        If Age > 132 Then CalculateAgeInWords = "11" & " Years and " & Age - 132 & " months": Exit Function
        If Age = 132 Then CalculateAgeInWords = "11" & " Years": Exit Function
        
        If Age > 120 Then CalculateAgeInWords = "10" & " Years and " & Age - 120 & " months": Exit Function
        If Age = 120 Then CalculateAgeInWords = "10" & " Years": Exit Function
        If Age > 108 Then CalculateAgeInWords = "9" & " Years and " & Age - 108 & " months": Exit Function
        If Age = 108 Then CalculateAgeInWords = "9" & " Years": Exit Function
        If Age > 96 Then CalculateAgeInWords = "8" & " Years and " & Age - 96 & " months": Exit Function
        If Age = 96 Then CalculateAgeInWords = "8" & " Years": Exit Function
        
        If Age > 84 Then CalculateAgeInWords = "7" & " Years and " & Age - 84 & " months": Exit Function
        If Age = 84 Then CalculateAgeInWords = "7" & " Years": Exit Function
        If Age > 72 Then CalculateAgeInWords = "6" & " Years and " & Age - 72 & " months": Exit Function
        If Age = 72 Then CalculateAgeInWords = "6" & " Years": Exit Function
        If Age > 60 Then CalculateAgeInWords = "5" & " Years and " & Age - 60 & " months": Exit Function
        If Age = 60 Then CalculateAgeInWords = "5" & " Years": Exit Function
        
        
        If Age > 48 Then CalculateAgeInWords = "4" & " Years and " & Age - 48 & " months": Exit Function
        If Age = 48 Then CalculateAgeInWords = "4" & " Years": Exit Function
        If Age > 36 Then CalculateAgeInWords = "3" & " Years and " & Age - 36 & " months": Exit Function
        If Age = 36 Then CalculateAgeInWords = "3" & " Years": Exit Function
        If Age > 24 Then CalculateAgeInWords = "2" & " Years and " & Age - 24 & " months": Exit Function
        If Age = 24 Then CalculateAgeInWords = "2" & " Years": Exit Function
        If Age > 12 Then CalculateAgeInWords = "1" & " Years and " & Age - 12 & " months": Exit Function
        If Age = 12 Then CalculateAgeInWords = "1" & " Year": Exit Function
        If Age >= 1 Then CalculateAgeInWords = Age & " Months": Exit Function
        Age = DateDiff("d", PatientDOB, Now)
        CalculateAgeInWords = Age & " Days": Exit Function
        Exit Function
    End If
End Function

Public Function getUserName(ToFindUserID As Long) As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & ToFindUserID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            getUserName = DecreptedWord(!UserName)
        End If
        .Close
    End With
End Function


Public Function getAllItems() As Collection
    Dim allItems As New Collection
    Dim rsTem As New ADODB.Recordset
    Dim NewItem As Item
    
    With rsTem
        temSql = "SELECT * from tblItem order by Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            NewItem = New Item
            NewItem.ID = !ItemID
            allItems.Add NewItem
            .MoveNext
        Wend
        .Close
    End With
    
    Set getAllItems = allItems
    
End Function
