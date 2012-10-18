Attribute VB_Name = "ModuleFind"
Option Explicit
    Dim temSql As String
    Dim i As Integer
    
Public Function LastDateOfMonth(ByVal SuppliedDate As Date) As Date
    If Month(SuppliedDate) = 12 Then
        LastDateOfMonth = DateSerial(Year(SuppliedDate), 12, 31)
    Else
        LastDateOfMonth = DateSerial(Year(SuppliedDate), Month(SuppliedDate) + 1, 1) - 1
    End If
End Function

Public Function SpecialityID(StaffID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & StaffID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            SpecialityID = !SpecialityID
        Else
            SpecialityID = 0
        End If
        .Close
    End With
End Function

Public Function ServiceCategory(ID As Long) As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceCategory where ServiceCategoryID = " & ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ServiceCategory = !ServiceCategory
        Else
            ServiceCategory = ""
        End If
        .Close
    End With
End Function

Public Function ServiceSubcategory(ID As Long) As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubcategory where ServiceSubcategoryID = " & ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ServiceSubcategory = !ServiceSubcategory
        Else
            ServiceSubcategory = ""
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

Public Function NewOPDBillID(BillDate As Date, BillTime As Date) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IsOPDBill = 1 AND Completed = 0  AND StoreID = " & UserStoreID & "  AND UserID = " & UserID & " Order by IncomeBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewOPDBillID = !IncomeBillID
        Else
            .AddNew
            !IsOPDBill = True
            !Date = Date
            !Time = Now
            !UserID = UserID
            !StoreID = UserStoreID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewOPDBillID = !NewID
        End If
        .Close
    End With
End Function

Public Function NewOPDDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblOPDBill where OPDBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IncomeBillID = IncomeBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewOPDDisplayBillID = !NewID
        .Close
    End With
End Function

Public Function NewPharmacyDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPharmacyBill where PharmacyBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IncomeBillID = IncomeBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewPharmacyDisplayBillID = !NewID
        .Close
    End With
End Function

Public Function NewHSSPaymentDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblHSSPaymentBill where HSSPaymentBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewHSSPaymentDisplayBillID = !NewID
        .Close
    End With
End Function


Public Function NewAgentDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblAgentBill where AgentBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewAgentDisplayBillID = !NewID
        .Close
    End With
End Function

Public Function NewGSBDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblGSBBill where GSBBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewGSBDisplayBillID = !NewID
        .Close
    End With
End Function

Public Function NewBHTDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHTBill where BHTBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewBHTDisplayBillID = !NewID
        .Close
    End With
End Function

Public Function NewRDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRBill where RBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewRDisplayBillID = !NewID
        .Close
    End With
End Function


Public Function NewMedicalTestDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblMedicalTestBill where MedicalTestBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewMedicalTestDisplayBillID = !NewID
        .Close
    End With
End Function


Public Function NewLabDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblLabBill where LabBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewLabDisplayBillID = !NewID
        .Close
    End With
End Function


Public Function NewExpenceDisplayBillID(IncomeBillID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblExpenceBill where ExpenceBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeBillID = IncomeBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        NewExpenceDisplayBillID = !NewID
        .Close
    End With
End Function


Public Function NewRBillID(BillDate As Date, BillTime As Date) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IsRBill = 1 AND Completed = 0  AND StoreID = " & UserStoreID & "  AND UserID = " & UserID & " Order by IncomeBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewRBillID = !IncomeBillID
        Else
            .AddNew
            !IsRBill = True
            !Date = Date
            !Time = Now
            !UserID = UserID
            !StoreID = UserStoreID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewRBillID = !NewID
        End If
        .Close
    End With
End Function


Public Function NewMedicalTestBillID(BillDate As Date, BillTime As Date) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IsMedicalTestBill = 1 AND Completed = 0  AND StoreID = " & UserStoreID & "  AND UserID = " & UserID & " Order by IncomeBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewMedicalTestBillID = !IncomeBillID
        Else
            .AddNew
            !IsMedicalTestBill = True
            !Date = Date
            !Time = Now
            !UserID = UserID
            !StoreID = UserStoreID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewMedicalTestBillID = !NewID
        End If
        .Close
    End With
End Function

Public Function NewLabBillID(BillDate As Date, BillTime As Date) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IsLabBill = 1 AND Completed = 0  AND StoreID = " & UserStoreID & " AND UserID = " & UserID & " Order by IncomeBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewLabBillID = !IncomeBillID
        Else
            .AddNew
            !IsLabBill = True
            !Date = Date
            !Time = Now
            !UserID = UserID
            !StoreID = UserStoreID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewLabBillID = !NewID
        End If
        .Close
    End With
End Function

Public Function NewExpenceBillID(BillDate As Date, BillTime As Date) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IsExpenceBill = 1 AND Completed = 0  AND StoreID = " & UserStoreID & " AND UserID = " & UserID & " Order by IncomeBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewExpenceBillID = !IncomeBillID
        Else
            .AddNew
            !IsExpenceBill = True
            !Date = Date
            !Time = Now
            !UserID = UserID
            !StoreID = UserStoreID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            NewExpenceBillID = !NewID
        End If
        .Close
    End With
End Function

Public Function GetControlType(MyControl As Control) As ControlType
    GetControlType = Unknown
    If TypeOf MyControl Is TextBox Then
        GetControlType = TextBox
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlType = ComboBox
    ElseIf TypeOf MyControl Is Button Then
        GetControlType = Button
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlType = CheckBox
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlType = DataCombo
    ElseIf TypeOf MyControl Is DataList Then
        GetControlType = DataList
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlType = DateTimePicker
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlType = Grid
    ElseIf TypeOf MyControl Is Label Then
        GetControlType = Label
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlType = ListBox
    ElseIf TypeOf MyControl Is Menu Then
        GetControlType = MenuItem
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlType = OptionButton
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlType = SSTab
    End If

End Function

Public Function GetControlText(MyControl As Control) As String
    GetControlText = Empty
    If TypeOf MyControl Is TextBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Button Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DataList Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 3)
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 4)
    ElseIf TypeOf MyControl Is Label Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Menu Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlText = MyControl.Caption
    End If
End Function

Public Function GetFormID(FormName As String, FormText As String) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblForm where Form = '" & FormName & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FormText = FormText
        Else
            .AddNew
            !FormText = FormText
            !Form = FormName
        End If
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        GetFormID = !NewID
        .Close
    End With
End Function


Public Sub EnableControls(MyForm As Form)
    Dim MyControl As Control
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                MyControl.Enabled = !Enabled
            Else
                MyControl.Enabled = True
            End If
            .Close
        End With
    Next
End Sub

Public Sub VisibleControls(MyForm As Form)
    Dim MyControl As Control
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                MyControl.Visible = !Visible
            Else
                MyControl.Visible = True
            End If
            .Close
        End With
    Next
End Sub


Private Function GetControlID(FormID As Long, MyControl As Control) As Long
    GetControlID = 0
        On Error Resume Next

    Dim rsForm As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
            With rsForm
                If TypeOf MyControl Is SSTab Then
                    For i = 0 To MyControl.Tabs - 1
                        If .State = 1 Then .Close
                        temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "' AND ControlIndex = " & i
                        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                        MyControl.Tab = i
                        If .RecordCount > 0 Then
                            !ControlText = GetControlText(MyControl)
                        Else
                            .AddNew
                            !FormID = FormID
                            !Control = MyControl.Name
                            !ControlType = GetControlType(MyControl)
                            !ControlText = GetControlText(MyControl)
                            !ControlIndex = i
                        End If
                        .Update
                        temSql = "SELECT @@IDENTITY AS NewID"
                        .Close
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        GetControlID = !NewID
                        .Close
                    Next i
                Else
                    If .State = 1 Then .Close
                    temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "'"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ControlText = GetControlText(MyControl)
                    Else
                        .AddNew
                        !FormID = FormID
                        !Control = MyControl.Name
                        !ControlType = GetControlType(MyControl)
                        !ControlText = GetControlText(MyControl)
                    End If
                    .Update
                    temSql = "SELECT @@IDENTITY AS NewID"
                    .Close
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    GetControlID = !NewID
                    .Close
                End If
            End With
End Function

Public Sub UpdateCompanyBalance(CompanyID As Long, UpdateValue As Double, AddToSupplier As Boolean, DeductSupplier As Boolean, ResetSupplier As Boolean, Optional PaymentMethodID As Long, Optional PaymentComments As String)
    Dim rsSup As New ADODB.Recordset
    Dim OldBalance As Double
    Dim NewBalance As Double
    With rsSup
        If .State = 1 Then .Close
        temSql = "Select * from tblHealthSchemeSuppliers where HealthSchemeSupplierID = " & CompanyID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            OldBalance = !Balance
            NewBalance = UpdateValue
            If AddToSupplier = True Then
                !Balance = !Balance + UpdateValue
            ElseIf DeductSupplier = True Then
                !Balance = !Balance - UpdateValue
            ElseIf ResetSupplier = True Then
                !Balance = UpdateValue
            End If
            .Update
            .Close
            temSql = "Select * from tblHealthSchemeSupplierPayment"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !HealthSchemeSupplierID = CompanyID
            If AddToSupplier = True Then
                !PaymentValue = UpdateValue
            ElseIf DeductSupplier = True Then
                !PaymentValue = 0 - UpdateValue
            ElseIf ResetSupplier = True Then
                !PaymentValue = NewBalance - OldBalance
            End If
            
            !PaymentUserID = UserID
            !PaymentDate = Date
            !PaymentTime = Now
            !PaymentDateTime = Now
            !PaymentComments = PaymentComments
            !PaymentMethodID = PaymentMethodID
            .Update
            .Close
        End If
    End With
End Sub

Public Sub UpdateBHTBalance(BHTID As Long, UpdateValue As Double, AddToBHT As Boolean, DeductBHT As Boolean, ResetBHT As Boolean)
    Dim rsSup As New ADODB.Recordset
    Dim OldBalance As Double
    Dim NewBalance As Double
    With rsSup
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where BHTID = " & BHTID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            OldBalance = !Balance
            NewBalance = UpdateValue
            If AddToBHT = True Then
                !Balance = !Balance + UpdateValue
            ElseIf DeductBHT = True Then
                !Balance = !Balance - UpdateValue
            ElseIf ResetBHT = True Then
                !Balance = UpdateValue
            End If
            .Update
            .Close
        End If
    End With
End Sub


