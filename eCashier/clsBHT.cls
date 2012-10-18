VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim MyPt As New clsPatient
    Dim BHTIDV As Long
    Dim BHTV As String
    Dim PatientIDV As Long
    Dim DOAV As Date
    Dim TOAV As Date
    Dim DODV As Date
    Dim TODV As Date
    Dim CommentsV As String
    Dim ReferringDoctorIDV As Long
    Dim ReferringDoctorV As String
    Dim PatientCategoryIDV As Long
    Dim PatientCategoryV As String
    Dim PaymentMethodIDV As Long
    Dim PaymentMethodV As String
    Dim FirstNameV As String
    Dim DateOfBirthV As Date
    Dim SexV As String
    Dim AgeInWordsV As String
    Dim DIschargeV As Boolean
    Dim PtSurchargeV As Double
    Dim ComSurchargeV As Double
    Dim HealthSchemeSupplierIDV As Long
    Dim HealthSchemeSupplierV As String
    Dim GuardianNameV As String
    Dim GuardianAddressV As String
    Dim DiscountV As Double
    Dim IsBHTV As Boolean
    Dim IsGSBV As Boolean
    Dim ForeignerV As Boolean
    Dim PtAddressV As String


    Dim MaintananceRateV As Double
    Dim LaterLinanRateV As Double
    Dim NursingRateV As Double
    Dim ICUNursingRateV As Double
    Dim AdmissionRateV As Double
    Dim InitialLinanRateV As Double
    Dim MaintainaceCashDiscountRateV As Double

    
    Dim AdmissionFeeV As Double
    Dim AdmissionChargeV As Double
    Dim LinanChargeV As Double
    Dim RoomChargeV As Double
    Dim ServicesChargeV As Double
    Dim MaintananceChargeV As Double
    Dim NursingChargeV As Double
    Dim ProfessionalChargeV As Double


Public Property Let BHTID(ID As Long)
    Call ClearValues
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT MaintananceRate, LaterLinanRate, NursingRate, ICUNursingRate, AdmissionRate, InitialLinanRate, MaintainaceCashDiscountRate, PtSurcharge, " & _
                      "ComSurcharge, AdmissionFee, AdmissionCharge, LinanCharge, RoomCharge, ServicesCharge, MaintananceCharge, NursingCharge, " & _
                      "ProfessionalCharge , tblBHT.BHTID, tblBHT.IsGSB, tblBHT.Foreigner, tblBHT.IsBHT, tblBHT.Discount, tblBHT.GuardianAddress, tblBHT.GuardianName, tblBHT.PtSurcharge, tblBHT.ReferringDoctorID, tblBHT.Comments, tblBHT.ComSurcharge, tblBHT.Discharge, tblBHT.BHT, tblBHT.DOA, tblBHT.TOA, tblBHT.DOD, tblBHT.TOD, tblPatientCategory.PatientCategoryID, tblPatientCategory.PatientCategory, tblPaymentMethod.PaymentMethodID, tblPaymentMethod.PaymentMethod, tblPatientMainDetails.PatientID, tblPatientMainDetails.FirstName, tblPatientMainDetails.DateOfBirth, tblSex.Sex, tblBHT.HealthSchemeSupplierID, tblStaffTitle.Title, tblStaff.Name, tblHealthSchemeSuppliers.HealthSchemeSupplierName, tblPatientMainDetails.NICNo, tblPatientMainDetails.Phone, tblPatientMainDetails.Address " & _
                    "FROM (((((tblBHT LEFT JOIN tblPatientCategory ON tblBHT.PatientCategoryID = tblPatientCategory.PatientCategoryID) LEFT JOIN tblPatientMainDetails ON tblBHT.PatientID = tblPatientMainDetails.PatientID) LEFT JOIN tblPaymentMethod ON tblPatientCategory.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblSex ON tblPatientMainDetails.SexID = tblSex.SexID) LEFT JOIN tblHealthSchemeSuppliers ON tblBHT.HealthSchemeSupplierID = tblHealthSchemeSuppliers.HealthSchemeSupplierID) LEFT JOIN (tblStaff LEFT JOIN tblTitle AS tblStaffTitle ON tblStaff.TitleID = tblStaffTitle.TitleID) ON tblBHT.ReferringDoctorID = tblStaff.StaffID " & _
                    "WHERE (((tblBHT.BHTID)=" & ID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            BHTIDV = !BHTID
            BHTV = !BHT
            DIschargeV = !Discharge
            PatientIDV = !PatientID
            MyPt.ID = !PatientID
            AgeInWordsV = MyPt.AgeInWords
            DOAV = !DOA
            On Error Resume Next
            TOAV = Format(!TOA, "hh:mm:ss")
            On Error GoTo 0
            If Not IsNull(!DOD) Then DODV = !DOD
            If Not IsNull(!TOD) Then TODV = Format(!TOD, "hh:mm:ss")
            If Not IsNull(!PatientCategoryID) Then PatientCategoryIDV = !PatientCategoryID
            If Not IsNull(!PatientCategory) Then PatientCategoryV = !PatientCategory
            If Not IsNull(!PaymentMethodID) Then PaymentMethodIDV = !PaymentMethodID
            If Not IsNull(!PaymentMethod) Then PaymentMethodV = !PaymentMethod
            If Not IsNull(!PtSurcharge) Then PtSurchargeV = !PtSurcharge
            If Not IsNull(!ComSurcharge) Then ComSurchargeV = !ComSurcharge
            If Not IsNull(!Comments) Then CommentsV = !Comments
            If Not IsNull(!ReferringDoctorID) Then ReferringDoctorIDV = !ReferringDoctorID
            If Not IsNull(![Name]) Then ReferringDoctorV = ![Name]
            If Not IsNull(!Title) Then ReferringDoctorV = !Title & " " & ReferringDoctorV
            If Not IsNull(!HealthSchemeSupplierID) Then HealthSchemeSupplierIDV = !HealthSchemeSupplierID
            If Not IsNull(!HealthSchemeSupplierName) Then HealthSchemeSupplierV = !HealthSchemeSupplierName
            
            If Not IsNull(!Address) Then PtAddressV = !Address
            
            If Not IsNull(!GuardianAddress) Then GuardianAddressV = !GuardianAddress
            If Not IsNull(!GuardianName) Then GuardianNameV = !GuardianName
            If Not IsNull(!Discount) Then DiscountV = !Discount
            
            
            If Not IsNull(!MaintananceRate) Then MaintananceRateV = !MaintananceRate
            If Not IsNull(!LaterLinanRate) Then LaterLinanRateV = !LaterLinanRate
            If Not IsNull(!NursingRate) Then NursingRateV = !NursingRate
            If Not IsNull(!ICUNursingRate) Then ICUNursingRateV = !ICUNursingRate
            If Not IsNull(!AdmissionRate) Then AdmissionRateV = !AdmissionRate
            If Not IsNull(!InitialLinanRate) Then InitialLinanRateV = !InitialLinanRate
            If Not IsNull(!MaintainaceCashDiscountRate) Then MaintainaceCashDiscountRateV = !MaintainaceCashDiscountRate
            If Not IsNull(!PtSurcharge) Then PtSurchargeV = !PtSurcharge
            If Not IsNull(!ComSurcharge) Then ComSurchargeV = !ComSurcharge
            If Not IsNull(!AdmissionFee) Then AdmissionFeeV = !AdmissionFee
            If Not IsNull(!AdmissionCharge) Then AdmissionChargeV = !AdmissionCharge
            If Not IsNull(!LinanCharge) Then LinanChargeV = !LinanCharge
            If Not IsNull(!RoomCharge) Then RoomChargeV = !RoomCharge
            If Not IsNull(!ServicesCharge) Then ServicesChargeV = !ServicesCharge
            If Not IsNull(!MaintananceCharge) Then MaintananceChargeV = !MaintananceCharge
            If Not IsNull(!NursingCharge) Then NursingChargeV = !NursingCharge
            If Not IsNull(!ProfessionalCharge) Then ProfessionalChargeV = !ProfessionalCharge
            
            
            
            IsBHTV = !IsBHT
            IsGSBV = !IsGSB
            ForeignerV = !Foreigner
            PatientIDV = !PatientID
            FirstNameV = !FirstName
            DateOfBirthV = !DateOfBirth
            If Not IsNull(!Sex) Then SexV = !Sex
    End If
        .Close
    End With
End Property

Private Sub ClearValues()
    BHTIDV = Empty
    BHTV = Empty
    DIschargeV = Empty
    PatientIDV = Empty
    MyPt.ID = Empty
    AgeInWordsV = Empty
    DOAV = Empty
    TOAV = Empty
    DODV = Empty
    TODV = Empty
    PatientCategoryIDV = Empty
    PatientCategoryV = Empty
    PaymentMethodIDV = Empty
    PaymentMethodV = Empty
    PtSurchargeV = Empty
    ComSurchargeV = Empty
    CommentsV = Empty
    ReferringDoctorIDV = Empty
    ReferringDoctorV = Empty
    HealthSchemeSupplierIDV = Empty
    HealthSchemeSupplierV = Empty
    PatientIDV = Empty
    FirstNameV = Empty
    DateOfBirthV = Empty
    SexV = Empty
    GuardianNameV = Empty
    GuardianAddressV = Empty
    DiscountV = Empty
    IsBHTV = False
    IsGSBV = False
    ForeignerV = False
    
    MaintananceRateV = 0
    LaterLinanRateV = 0
    NursingRateV = 0
    ICUNursingRateV = 0
    AdmissionRateV = 0
    InitialLinanRateV = 0
    MaintainaceCashDiscountRateV = 0
    PtSurchargeV = 0
    ComSurchargeV = 0
    AdmissionFeeV = 0
    AdmissionChargeV = 0
    LinanChargeV = 0
    RoomChargeV = 0
    ServicesChargeV = 0
    MaintananceChargeV = 0
    NursingChargeV = 0
    ProfessionalChargeV = 0
    
End Sub

Public Property Get IsBHT() As Boolean
    IsBHT = IsBHTV
End Property

Public Property Get Foreigner() As Boolean
    Foreigner = ForeignerV
End Property

Public Property Get IsGSB() As Boolean
    IsGSB = IsGSBV
End Property

Public Property Get Discount() As String
    Discount = DiscountV
End Property

Public Property Get GuardianName() As String
    GuardianName = GuardianNameV
End Property

Public Property Get GuardianAddress() As String
    GuardianAddress = GuardianAddressV
End Property

Public Property Get PtAddress() As String
    PtAddress = PtAddressV
End Property


Public Property Get HealthSchemeSupplier() As String
    HealthSchemeSupplier = HealthSchemeSupplierV
End Property

Public Property Get HealthSchemeSupplierID() As String
    HealthSchemeSupplierID = HealthSchemeSupplierIDV
End Property

Public Property Get ReferringDoctor() As String
    ReferringDoctor = ReferringDoctorV
End Property

Public Property Get ReferringDoctorID() As Long
    ReferringDoctorID = ReferringDoctorIDV
End Property

Public Property Get Comments() As String
    Comments = CommentsV
End Property

Public Property Get BHTID() As Long
    BHTID = BHTIDV
End Property

Public Property Get BHT() As String
        BHT = BHTV
End Property

Public Property Get DOA() As Date
    DOA = DOAV
End Property

Public Property Get TOA() As Date
    TOA = TOAV
End Property

Public Property Get DOD() As Date
    DOD = DODV
End Property

Public Property Get TOD() As Date
    TOD = TODV
End Property

Public Property Get PatientCategoryID() As Long
    PatientCategoryID = PatientCategoryIDV
End Property

Public Property Get PatientCategory() As String
    PatientCategory = PatientCategoryV
End Property

Public Property Get PaymentMethodID() As Long
    PaymentMethodID = PaymentMethodIDV
End Property

Public Property Get PaymentMethod() As String
    PaymentMethod = PaymentMethodV
End Property

Public Property Get PatientID() As Long
    PatientID = PatientIDV
End Property

Public Property Get FirstName() As String
   FirstName = FirstNameV
End Property

Public Property Get DateOfBirth() As Date
    DateOfBirth = DateOfBirthV
End Property

Public Property Get Sex() As String
    Sex = SexV
End Property

Public Property Get AgeInWords() As String
    AgeInWords = AgeInWordsV
End Property

Public Property Get Discharge() As Boolean
    Discharge = DIschargeV
End Property

Public Property Get PtSurcharge() As Double
    PtSurcharge = PtSurchargeV
End Property

Public Property Get ComSurcharge() As Double
    ComSurcharge = ComSurchargeV
End Property



Public Property Get MaintananceRate() As Double
     MaintananceRate = MaintananceRateV
End Property

Public Property Get LaterLinanRate() As Double
   LaterLinanRate = LaterLinanRateV
End Property

Public Property Get NursingRate() As Double
         NursingRate = NursingRateV
End Property

Public Property Get ICUNursingRate() As Double
       ICUNursingRate = ICUNursingRateV
End Property

Public Property Get AdmissionRate() As Double
                 AdmissionRate = AdmissionRateV
End Property

Public Property Get InitialLinanRate() As Double
               InitialLinanRate = InitialLinanRateV
End Property

Public Property Get MaintainaceCashDiscountRate() As Double
  MaintainaceCashDiscountRate = MaintainaceCashDiscountRateV
End Property




Public Property Get AdmissionFee() As Double
      AdmissionFee = AdmissionFeeV
End Property

Public Property Get AdmissionCharge() As Double
     AdmissionCharge = AdmissionChargeV
End Property

Public Property Get LinanCharge() As Double
     LinanCharge = LinanChargeV
End Property

Public Property Get RoomCharge() As Double
    RoomCharge = RoomChargeV
End Property

Public Property Get ServicesCharge() As Double
   ServicesCharge = ServicesChargeV
End Property

Public Property Get MaintananceCharge() As Double
     MaintananceCharge = MaintananceChargeV
End Property

Public Property Get NursingCharge() As Double
      NursingCharge = NursingChargeV
End Property
          
Public Property Get ProfessionalCharge() As Double
   ProfessionalCharge = ProfessionalChargeV
End Property
