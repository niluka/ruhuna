VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim RoomV As String
    Dim RoomIDV As Long
    Dim CategoryV As String
    Dim ICUNursingV As Boolean
    Dim CategoryIDV As Long
    Dim GeneralChargeV As Double
    Dim DiscountForCashV As Double
    Dim SurchargeForCreditV As Double
     

Public Property Let RoomID(ID As Long)
    Dim rsTem As New ADODB.Recordset
    RoomV = 0
    RoomIDV = 0
    CategoryV = 0
    CategoryIDV = 0
    ICUNursingV = False
    GeneralChargeV = 0
    DiscountForCashV = 0
    SurchargeForCreditV = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblRoom.Room, tblRoom.RoomID,tblRoomCategory.ICUNursing , tblRoomCategory.RoomCategoryID,  tblRoomCategory.SurchargeForCredit, tblRoomCategory.RoomCategory, tblRoomCategory.GeneralCharge, tblRoomCategory.DiscountForCash " & _
                    "FROM tblRoom LEFT JOIN tblRoomCategory ON tblRoom.RoomCategoryID = tblRoomCategory.RoomCategoryID " & _
                    "WHERE (((tblRoom.RoomID)=" & ID & "))"
        RoomIDV = ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            RoomV = !Room
            If Not IsNull(!RoomCategory) Then CategoryV = !RoomCategory
            If Not IsNull(!RoomCategoryID) Then CategoryIDV = !RoomCategoryID
            If Not IsNull(!GeneralCharge) Then GeneralChargeV = !GeneralCharge
            If Not IsNull(!DiscountForCash) Then DiscountForCashV = !DiscountForCash
            If Not IsNull(!SurchargeForCredit) Then SurchargeForCreditV = !SurchargeForCredit
            If !ICUNursing = True Then ICUNursingV = True
        End If
        .Close
    End With
End Property

Public Property Get ICUNursing() As Boolean
    ICUNursing = ICUNursingV
End Property


Public Property Get Room() As String
    Room = RoomV
End Property

Public Property Get RoomID() As Long
    RoomID = RoomIDV
End Property

Public Property Get Category() As String
    Category = CategoryV
End Property

Public Property Get CategoryID() As Long
    CategoryID = CategoryIDV
End Property

Public Property Get GeneralCharge() As Double
    GeneralCharge = GeneralChargeV
End Property

Public Property Get DiscountForCash() As Double
    DiscountForCash = DiscountForCashV
End Property

Public Property Get SurchargeForCredit() As Double
    SurchargeForCredit = SurchargeForCreditV
End Property
