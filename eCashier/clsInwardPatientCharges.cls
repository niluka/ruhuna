VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInwardPatientCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim BHTIDV As Double
    Dim AdmissionRateV As Double
    Dim InitialLinanRateV As Double
    Dim LaterLinanRateV As Double
    Dim MaintananceRateV As Double
    Dim MaintananceCashDiscountRateV As Double
    Dim NursingRateV As Double
    Dim ICUNursingRateV As Double
    Dim ForBHTV As Long
    
Public Property Let BHTID(BHTID As Long)
    BHTIDV = BHTID
    Call GetDetails
End Property

Public Property Get MaintainaceCashDiscountRate() As Double
    If MaintananceCashDiscountRateV = 0 Then
        GetDetails
    End If
    MaintainaceCashDiscountRate = MaintananceCashDiscountRateV
End Property
    
Public Property Get AdimssionRate() As Double
    If AdmissionRateV = 0 Then Call GetDetails
    AdimssionRate = AdmissionRateV
End Property

Public Property Get InitialLinanRate() As Double
    If InitialLinanRateV = 0 Then Call GetDetails
    InitialLinanRate = InitialLinanRateV
End Property

Public Property Get LaterLinanRate() As Double
    If LaterLinanRateV = 0 Then Call GetDetails
    LaterLinanRate = LaterLinanRateV
End Property

Public Property Get MaintananceRate() As Double
    If MaintananceRateV = 0 Then Call GetDetails
    MaintananceRate = MaintananceRateV
End Property

Public Property Get NursingRate()
    If NursingRateV = 0 Then Call GetDetails
    NursingRate = NursingRateV
End Property

Public Property Get ICUNursingRate()
    If ICUNursingRateV = 0 Then Call GetDetails
    ICUNursingRate = ICUNursingRateV
End Property


Private Sub GetDefaults()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblInwardPatientRates"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            AdmissionRateV = !AdmissionRate
            InitialLinanRateV = !InitialLinanRate
            LaterLinanRateV = !LaterLinanRate
            MaintananceRateV = !MaintananceRate
            NursingRateV = !NursingRate
            ICUNursingRateV = !ICUNursingRate
            MaintananceCashDiscountRateV = !MaintananceCashDiscountRate
        Else
            .AddNew
            !AdmissionRate = 470
            !InitialLinanRate = 200
            !LaterLinanRate = 120
            !NursingRate = 100
            !MaintananceRate = 400
            !ICUNursingRate = 200
            !MaintananceCashDiscountRate = 100
            AdmissionRateV = !AdmissionRate
            InitialLinanRateV = !InitialLinanRate
            LaterLinanRateV = !LaterLinanRate
            MaintananceRateV = !MaintananceRate
            NursingRateV = !NursingRate
            ICUNursingRateV = !ICUNursingRate
            MaintananceCashDiscountRateV = !MaintananceCashDiscountRate
            .Update
        End If
        .Close
    End With
End Sub

Private Sub GetDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT Where BHTID = " & BHTIDV
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Discharge = True Then
                AdmissionRateV = !AdmissionRate
                InitialLinanRateV = !InitialLinanRate
                LaterLinanRateV = !LaterLinanRate
                MaintananceRateV = !MaintananceRate
                NursingRateV = !NursingRate
                ICUNursingRateV = !ICUNursingRate
                MaintananceCashDiscountRateV = !MaintainaceCashDiscountRate
            Else
                Call GetDefaults
            End If
        Else
            Call GetDefaults
        End If
        .Close
    End With
End Sub
