VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAutoOrdering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Ordering"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9645
   Begin btButtonEx.ButtonEx bttnCreate 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Create Master Order Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Usage"
      TabPicture(0)   =   "frmAutoOrdering.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "mvUsageFrom"
      Tab(0).Control(3)=   "mvUsageTo"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Processes"
      TabPicture(1)   =   "frmAutoOrdering.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Items"
      TabPicture(2)   =   "frmAutoOrdering.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Order For"
      TabPicture(3)   =   "frmAutoOrdering.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "Label5"
      Tab(3).Control(2)=   "mvOrderFrom"
      Tab(3).Control(3)=   "mvOrderTo"
      Tab(3).Control(4)=   "Frame2"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Other"
      TabPicture(4)   =   "frmAutoOrdering.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "chkIgnoreOrdersPending"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Department"
         Height          =   1215
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   3015
         Begin VB.OptionButton OptOneDept 
            Caption         =   "Our Department only"
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   2655
         End
         Begin VB.OptionButton OptAllDept 
            Caption         =   "All Departments"
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   8655
         Begin VB.OptionButton OptAllROL 
            Caption         =   "All items below ROL"
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton OptSelectedROL 
            Caption         =   "Selected Items Below ROL"
            Height          =   255
            Left            =   5760
            TabIndex        =   46
            Top             =   240
            Width           =   2775
         End
         Begin VB.ListBox lstSelectedItemID 
            Height          =   2460
            Left            =   7440
            TabIndex        =   45
            Top             =   960
            Visible         =   0   'False
            Width           =   375
         End
         Begin btButtonEx.ButtonEx bttnAddItem 
            Height          =   375
            Left            =   4320
            TabIndex        =   43
            Top             =   1680
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "["
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 3"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ListBox lstItemID 
            Height          =   2460
            IntegralHeight  =   0   'False
            Left            =   3240
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ListBox lstSelectedItem 
            Height          =   2460
            Left            =   4920
            TabIndex        =   41
            Top             =   960
            Width           =   3615
         End
         Begin VB.ListBox lstItem 
            Height          =   2460
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   40
            Top             =   960
            Width           =   4095
         End
         Begin MSDataListLib.DataCombo dtcICatogery 
            Height          =   360
            Left            =   1320
            TabIndex        =   39
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.OptionButton OptSelectedItems 
            Caption         =   "Selected Items"
            Height          =   255
            Left            =   3960
            TabIndex        =   38
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton OptAllItems 
            Caption         =   "All items"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   2295
         End
         Begin btButtonEx.ButtonEx bttnRemoveItem 
            Height          =   375
            Left            =   4320
            TabIndex        =   44
            Top             =   2160
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Z"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 3"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CheckBox chkCatogery 
            Caption         =   "Catogery"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Order for"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   2415
         Begin VB.OptionButton OptOOther 
            Caption         =   "Other"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   2880
            Width           =   1935
         End
         Begin VB.OptionButton OptOOneWeek 
            Caption         =   "One week"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton OptOTwoWeeks 
            Caption         =   "Two weeks"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton OptOOneMonths 
            Caption         =   "One month"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton OptOTwoMonths 
            Caption         =   "Two months"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton OptOThreeMonths 
            Caption         =   "Three months"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1935
         End
         Begin VB.OptionButton OptOSixMonths 
            Caption         =   "Six months"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   1935
         End
         Begin VB.OptionButton OptOOneYear 
            Caption         =   "One year"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkIgnoreOrdersPending 
         Caption         =   "Ignore Pending Orders"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         Caption         =   "Consider Usage of"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   2415
         Begin VB.CheckBox chkConsumption 
            Caption         =   "Consumption"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkSales 
            Caption         =   "Sales"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkDiscard 
            Caption         =   "Discard"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox chkAdjustments 
            Caption         =   "Stock Adjustments"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Consider Usage For"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   2415
         Begin VB.OptionButton OptUOther 
            Caption         =   "Other"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   2880
            Width           =   1935
         End
         Begin VB.OptionButton optUOneWeek 
            Caption         =   "One week"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton OptUTwoWeeks 
            Caption         =   "Two weeks"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton OptUOneMonth 
            Caption         =   "One month"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton OptUTwoMonths 
            Caption         =   "Two months"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton OptUThreeMonths 
            Caption         =   "Three months"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   1935
         End
         Begin VB.OptionButton OptUSixMOnths 
            Caption         =   "Six months"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   2160
            Width           =   1935
         End
         Begin VB.OptionButton OptUOneYear 
            Caption         =   "One year"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   2520
            Width           =   1935
         End
      End
      Begin MSComCtl2.MonthView mvUsageTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2820
         Left            =   -69240
         TabIndex        =   26
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   67174401
         CurrentDate     =   39536
      End
      Begin MSComCtl2.MonthView mvOrderTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2820
         Left            =   -69240
         TabIndex        =   27
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   67174401
         CurrentDate     =   39536
      End
      Begin MSComCtl2.MonthView mvOrderFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2820
         Left            =   -72360
         TabIndex        =   28
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   67174401
         CurrentDate     =   39536
      End
      Begin MSComCtl2.MonthView mvUsageFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2820
         Left            =   -72360
         TabIndex        =   29
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   67174401
         CurrentDate     =   39536
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   -69240
         TabIndex        =   33
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   -72360
         TabIndex        =   32
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   -69240
         TabIndex        =   31
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   255
         Left            =   -72360
         TabIndex        =   30
         Top             =   600
         Width           =   3015
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select Ordering Criteria"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmAutoOrdering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemICatogery As New ADODB.Recordset
    Dim rsTemDistributors As New ADODB.Recordset
    Dim rsTemOrders As New ADODB.Recordset
    Dim rsTemStocks As New ADODB.Recordset
    Dim rsICatogery As New ADODB.Recordset
    Dim temSql As String
    Dim NewItem As New Item

Private Sub bttnAddItem_Click()
    If lstItem.ListCount < 1 Then Exit Sub
    Dim i As Integer
    For i = lstItem.ListCount - 1 To 0 Step -1
        If lstItem.Selected(i) = True Then
            If AlreadyAdded(Val(lstItemID.List(i))) = False Then
                lstSelectedItem.AddItem lstItem.List(i)
                lstSelectedItemID.AddItem lstItemID.List(i)
            Else
                Beep
            End If
        End If
    Next
End Sub

Private Function AlreadyAdded(ItemID As Long) As Boolean
AlreadyAdded = True
Dim i As Integer

For i = 0 To lstSelectedItemID.ListCount - 1
    If Val(lstSelectedItemID.List(i)) = ItemID Then Exit Function
Next

AlreadyAdded = False
End Function

Private Sub bttnCancel_Click()
    Dim tr  As Integer
    tr = MsgBox("Are you sure you want to cancel the automatic ordering ?", vbQuestion + vbYesNo, "Cancel?")
    If tr = vbYes Then Unload Me
End Sub

Private Sub bttnCreate_Click()
    Dim tr As Integer
    If CanCreate = False Then Exit Sub
    Call WriteOrderBill
    If optAllItems.Value = True Then
        If OrderForAllItems = True Then
            Call FinishOrderBill
            Unload Me
            frmConfirmAutoOrdering.Show
        Else
            tr = MsgBox("There are no items to be ordered", vbCritical, "No items")
            Unload Me
        End If
    Else
        If OrderForSelectedItems = True Then
            Call FinishOrderBill
            Unload Me
            frmConfirmAutoOrdering.Show
        Else
            tr = MsgBox("There are no items to be ordered", vbCritical, "No items")
            Unload Me
        End If
    End If
End Sub

Private Sub FinishOrderBill()
    With rsTemOrders
        If .State = 1 Then .Close
        temSql = "SELECT tblOrderBill.* FROM tblOrderBill where orderbillID = " & OrderBillID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !AutoRequestComplete = True
        .Update
    End With
End Sub

Private Sub WriteOrderBill()
    With rsTemOrders
        If .State = 1 Then .Close
        temSql = "SELECT tblOrderBill.* FROM tblOrderBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Autorequest = True
        !AutoRequestDate = Date
        !AutoRequestTime = Now
        !AutoRequestStaffID = UserID
        !AutoRequestStoreID = UserStoreID
        !UsageFrom = mvUsageFrom.Value
        !UsageTo = mvUsageTo.Value
        !OrderFrom = mvOrderFrom.Value
        !OrderTo = mvOrderTo.Value
        If chkIgnoreOrdersPending.Value = 1 Then
            !IgnorePendingOrders = True
        Else
            !IgnorePendingOrders = False
        End If
        !AllStores = True
        If chkConsumption.Value = 1 Then
            !Consumption = True
        Else
            !Consumption = False
        End If
        If chkSales.Value = 1 Then
            !Sale = True
        Else
            !Sale = False
        End If
        If chkDiscard.Value = 1 Then
            !Discard = True
        Else
            !Discard = False
        End If
        If chkAdjustments.Value = 1 Then
            !Adjustment = True
        Else
            !Adjustment = False
        End If
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        OrderBillID = !NewID
        .Close
    End With
End Sub

Private Function OrderForAllItems() As Boolean
    
    OrderForAllItems = False
    
    Dim TotalRequest As Double
    Dim TotalConsumption As Double
    Dim TotalSale As Double
    Dim TotalAdjustment As Double
    Dim TotalDiscard As Double
    Dim TotalUsage As Double
    Dim TotalRequirment As Double
    Dim UsageDuration As Long
    Dim OrderDuration As Long
    Dim CurrentStock As Double
    
    
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.* FROM tblItem ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 0 Then Exit Function
        While .EOF = False
            NewItem.ID = !ItemID
            With rsTemOrders
                If .State = 1 Then rsTemOrders.Close
                temSql = "SELECT tblOrder.* FROM tblOrder WHERE tblOrder.ItemID =" & rsTemItem!ItemID & " AND tblOrder.RequestComplete = 1  And tblOrder.ReceivedComplete = 0 "
                rsTemOrders.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount < 1 Or chkIgnoreOrdersPending.Value = 1 Then
                    If .State = 1 Then rsTemOrders.Close
                    temSql = "SELECT tblOrder.* FROM tblOrder"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    TotalRequest = 0
                    TotalUsage = 0
                    .AddNew
                    !OrderBillID = OrderBillID
                    !Autorequest = True
                    !AutoRequestDate = Date
                    !AutoRequestTime = Now
                    !AutoRequestStaffID = UserID
                    !AutoRequestStoreID = UserStoreID
                    !ItemID = rsTemItem!ItemID
                    !UsageFrom = mvUsageFrom.Value
                    !UsageTo = mvUsageTo.Value
                    !OrderFrom = mvOrderFrom.Value
                    !OrderTo = mvOrderTo.Value
                    If chkIgnoreOrdersPending.Value = 1 Then
                        !IgnorePendingOrders = True
                    Else
                        !IgnorePendingOrders = False
                    End If
                    !AllStores = True
                    If chkConsumption.Value = 1 Then
                        !Consumption = True
                        TotalConsumption = CalculateConsumption(rsTemItem!ItemID, mvUsageFrom.Value, mvUsageTo.Value)
                        !ConsumptionAmount = TotalConsumption
                        TotalUsage = TotalUsage + TotalConsumption
                    Else
                        !Consumption = False
                    End If
                    If chkSales.Value = 1 Then
                        TotalSale = CalculateSale(rsTemItem!ItemID, mvUsageFrom.Value, mvUsageTo.Value)
                        !Sale = True
                        !SaleAmount = TotalSale
                        TotalUsage = TotalUsage + TotalSale
                    Else
                        !Sale = False
                    End If
                    If chkDiscard.Value = 1 Then
                        TotalDiscard = CalculateDiscard(rsTemItem!ItemID, mvUsageFrom.Value, mvUsageTo.Value)
                        !Discard = True
                        !DiscardAmount = TotalDiscard
                        TotalUsage = TotalUsage + TotalDiscard
                    Else
                        !Discard = False
                    End If
                    If chkAdjustments.Value = 1 Then
                        TotalAdjustment = CalculateAdjustment(rsTemItem!ItemID, mvUsageFrom.Value, mvUsageTo.Value)
                        !Adjustment = True
                        !AdjustmentAmount = TotalAdjustment
                        TotalUsage = TotalUsage + TotalAdjustment
                    Else
                        !Adjustment = False
                    End If
                    UsageDuration = Abs(DateDiff("d", mvUsageFrom.Value, mvUsageTo.Value))
                    OrderDuration = Abs(DateDiff("d", mvOrderFrom.Value, mvOrderTo.Value))
                    TotalRequest = Round((TotalUsage * OrderDuration) / UsageDuration)
                    If NewItem.MinQty > TotalRequest Then
                        TotalRequest = NewItem.MinQty
                    End If
                    TotalRequest = (TotalRequest \ NewItem.IssueUnitsPerPack) * NewItem.IssueUnitsPerPack
                    If CalculateStock(rsTemItem!ItemID).Amount >= TotalRequest Then
                        !AutoRequestAmount = 0
                        !RequestAmount = 0
                        !AutoRequestComments = "The required amount of stocks of " & NewItem.Display & " for the planned duration of " & Format(mvOrderFrom.Value, LongDateFormat) & " to " & Format(mvOrderTo.Value, LongDateFormat) & " is available in the stocks at the moment. So no request was generated. "
                    Else
                        !AutoRequestAmount = TotalRequest
                        !RequestAmount = TotalRequest
                    End If
                    With rsTemDistributors
                        If .State = 1 Then .Close
                        temSql = "SELECT tblItemDistributor.ItemDistributorID, tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & rsTemItem!ItemID & "))"
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount < 1 Then
                            .Close
                            temSql = "SELECT tblDistrubutor.* FROM tblDistrubutor"
                            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                            If .RecordCount < 1 Then Exit Function
                            rsTemOrders!AutoRequestDistributorID = !DistributorID
                            rsTemOrders!RequestDistributorID = !DistributorID
                            rsTemOrders!AutoRequestComments = rsTemOrders!AutoRequestComments & " There was no specific distributor selected for this item. "
                            .Close
                        Else
                            rsTemOrders!AutoRequestDistributorID = !DistributorID
                             rsTemOrders!RequestDistributorID = !DistributorID
                            .Close
                        End If
                    End With
                    !AutoRequestComplete = True
                    .Update
                    OrderForAllItems = True
                End If
            End With
            .MoveNext
        Wend
        .Close
    End With
End Function

Private Function OrderForSelectedItems() As Boolean
    OrderForSelectedItems = False
    Dim TotalRequest As Double
    Dim TotalConsumption As Double
    Dim TotalSale As Double
    Dim TotalAdjustment As Double
    Dim TotalDiscard As Double
    Dim TotalUsage As Double
    Dim TotalRequirment As Double
    Dim UsageDuration As Long
    Dim OrderDuration As Long
    Dim CurrentStock As Double
    
    Dim i As Integer
    For i = 0 To lstSelectedItemID.ListCount - 1
        NewItem.ID = Val(lstSelectedItemID.List(i))
        With rsTemOrders
            If .State = 1 Then .Close
            temSql = "SELECT tblOrder.* FROM tblOrder WHERE tblOrder.ItemID =" & Val(lstSelectedItemID.List(i)) & " AND tblOrder.RequestComplete = 1  And tblOrder.ReceivedComplete = 0 "
            rsTemOrders.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount < 1 Or chkIgnoreOrdersPending.Value = 1 Then
                If .State = 1 Then rsTemOrders.Close
                temSql = "SELECT tblOrder.* FROM tblOrder"
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                TotalRequest = 0
                TotalUsage = 0
                .AddNew
                !OrderBillID = OrderBillID
                !Autorequest = True
                !AutoRequestDate = Date
                !AutoRequestTime = Now
                !AutoRequestStaffID = UserID
                !AutoRequestStoreID = UserStoreID
                !ItemID = Val(lstSelectedItemID.List(i))
                !UsageFrom = mvUsageFrom.Value
                !UsageTo = mvUsageTo.Value
                !OrderFrom = mvOrderFrom.Value
                !OrderTo = mvOrderTo.Value
                If chkIgnoreOrdersPending.Value = 1 Then
                    !IgnorePendingOrders = True
                Else
                    !IgnorePendingOrders = False
                End If
                !AllStores = True
                If chkConsumption.Value = 1 Then
                    !Consumption = True
                    TotalConsumption = CalculateConsumption(Val(lstSelectedItemID.List(i)), mvUsageFrom.Value, mvUsageTo.Value)
                    !ConsumptionAmount = TotalConsumption
                    TotalUsage = TotalUsage + TotalConsumption
                Else
                    !Consumption = False
                End If
                If chkSales.Value = 1 Then
                    TotalSale = CalculateSale(Val(lstSelectedItemID.List(i)), mvUsageFrom.Value, mvUsageTo.Value)
                    !Sale = True
                    !SaleAmount = TotalSale
                    TotalUsage = TotalUsage + TotalSale
                Else
                    !Sale = False
                End If
                If chkDiscard.Value = 1 Then
                    TotalDiscard = CalculateDiscard(Val(lstSelectedItemID.List(i)), mvUsageFrom.Value, mvUsageTo.Value)
                    !Discard = True
                    !DiscardAmount = TotalDiscard
                    TotalUsage = TotalUsage + TotalDiscard
                Else
                    !Discard = False
                End If
                If chkAdjustments.Value = 1 Then
                    TotalAdjustment = CalculateAdjustment(Val(lstSelectedItemID.List(i)), mvUsageFrom.Value, mvUsageTo.Value)
                    !Adjustment = True
                    !AdjustmentAmount = TotalAdjustment
                    TotalUsage = TotalUsage + TotalAdjustment
                Else
                    !Adjustment = False
                End If
                UsageDuration = Abs(DateDiff("d", mvUsageFrom.Value, mvUsageTo.Value))
                OrderDuration = Abs(DateDiff("d", mvOrderFrom.Value, mvOrderTo.Value))
                TotalRequest = Round((TotalUsage * OrderDuration) / UsageDuration)
                If NewItem.MinQty > TotalRequest Then
                    TotalRequest = NewItem.MinQty
                End If
                TotalRequest = (TotalRequest \ NewItem.IssueUnitsPerPack) * NewItem.IssueUnitsPerPack
                If CalculateStock(Val(lstSelectedItemID.List(i))).Amount >= TotalRequest Then
                    !AutoRequestAmount = 0
                    !RequestAmount = 0
                    !AutoRequestComments = "The required amount of stocks of " & NewItem.Display & " for the planned duration of " & Format(mvOrderFrom.Value, LongDateFormat) & " to " & Format(mvOrderTo.Value, LongDateFormat) & " is available in the stocks at the moment. So no request was generated. "
                Else
                    !AutoRequestAmount = TotalRequest
                    !RequestAmount = TotalRequest
                End If
                With rsTemDistributors
                    If .State = 1 Then .Close
                    temSql = "SELECT tblItemDistributor.* FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(lstSelectedItemID.List(i)) & "))"
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If .RecordCount < 1 Then
                        .Close
                        temSql = "SELECT tblDistrubutor.* FROM tblDistrubutor"
                        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                        If .RecordCount < 1 Then Exit Function
                        rsTemOrders!AutoRequestDistributorID = !DistributorID
                        rsTemOrders!AutoRequestComments = rsTemOrders!AutoRequestComments & " There was no specific distributor selected for this item. "
                        rsTemOrders!RequestDistributorID = !DistributorID
                        .Close
                    Else
                        rsTemOrders!AutoRequestDistributorID = !DistributorID
                        rsTemOrders!RequestDistributorID = !DistributorID
                        .Close
                    End If
                End With
                !AutoRequestComplete = True
                .Update
                OrderForSelectedItems = True
            End If
        End With
    Next i
End Function


Private Function CanCreate() As Boolean
    CanCreate = False
    Dim tr As Integer
    If DateDiff("d", mvOrderFrom.Value, mvOrderTo.Value) < 7 Then
        tr = MsgBox("You can't automatically order for less than seven days of activiti", vbCritical, "Order duration NOT sufficient")
        SSTab1.Tab = 3
        optUOneWeek.SetFocus
        Exit Function
    End If
    If DateDiff("d", mvUsageFrom.Value, mvUsageTo.Value) < 7 Then
        tr = MsgBox("You can't automatically order considering less than seven days of activiti", vbCritical, "Usage duration NOT sufficient")
        SSTab1.Tab = 0
        optUOneWeek.SetFocus
        Exit Function
    End If
    If chkConsumption.Value <> 1 And chkAdjustments.Value <> 1 And chkDiscard.Value <> 1 And chkSales.Value <> 1 Then
        tr = MsgBox("You must at least select one out of consumption, sales, discard or adjustments", vbCritical, "Select one")
        SSTab1.Tab = 1
        Exit Function
    End If
    If optAllItems.Value = False And lstSelectedItem.ListCount < 1 Then
        tr = MsgBox("You must select the items to order", vbCritical, "No Itme")
        SSTab1.Tab = 2
        optAllItems.SetFocus
        Exit Function
    End If
    CanCreate = True
End Function

Private Sub bttnRemoveItem_Click()
Dim i As Integer
For i = 0 To lstSelectedItem.ListCount - 1
    If lstSelectedItem.Selected(i) = True Then
        lstSelectedItemID.RemoveItem (lstSelectedItem.ListIndex)
        lstSelectedItem.RemoveItem (lstSelectedItem.ListIndex)
    End If
Next

End Sub


Private Sub chkCatogery_Click()
If chkCatogery.Value = 1 Then
    ListSelectedItems
Else
    ListAllItems
End If
End Sub

Private Sub ListAllItems()
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.* FROM tblItem  ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        lstItem.Clear
        lstItemID.Clear
        If .RecordCount > 0 Then
            lstItem.Visible = False
            While .EOF = False
                lstItem.AddItem !Display
                lstItemID.AddItem !ItemID
                .MoveNext
            Wend
            lstItem.Visible = True
        End If
        .Close
    End With
End Sub

Private Sub ListSelectedItems()
    If Not IsNumeric(dtcICatogery.BoundText) Then Exit Sub
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.* FROM tblItem WHERE (((tblItem.ItemCategoryID)=" & Val(dtcICatogery.BoundText) & ")) ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        lstItem.Clear
        lstItemID.Clear
        If .RecordCount > 0 Then
            lstItem.Visible = False
            While .EOF = False
                lstItem.AddItem !Display
                lstItemID.AddItem !ItemID
                .MoveNext
            Wend
            lstItem.Visible = True
        End If
        .Close
    End With
End Sub

Private Sub dtcICatogery_Change()
    Call ListSelectedItems
End Sub

Private Sub dtcICatogery_Click(Area As Integer)
    dtcICatogery_Change
End Sub

Private Sub Form_Load()
    Call GetSettings
    Call FillCombos
    Call CalculatePeriods
End Sub

Private Sub FillCombos()
With rsICatogery
    If .State = 1 Then .Close
    temSql = "SELECT tblItemCategory.* FROM tblItemCategory ORDER BY tblItemCategory.ItemCategory"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcICatogery
    Set .RowSource = rsICatogery
    .ListField = "ItemCategory"
    .BoundColumn = "ItemCategoryID"
End With
End Sub

Private Sub GetSettings()
    mvOrderFrom.MaxDate = Date
    mvUsageTo.MaxDate = Date
    optAllItems.Value = GetSetting(App.EXEName, "Options", "OptAllItems", True)
    OptAllROL.Value = GetSetting(App.EXEName, "Options", "OptAllROL", False)
    OptOOneMonths.Value = GetSetting(App.EXEName, "Options", "OptOOneMonths", False)
    OptOOneWeek.Value = GetSetting(App.EXEName, "Options", "OptOOneWeek", True)
    OptOOneYear.Value = GetSetting(App.EXEName, "Options", "OptOOneYear", False)
    OptOOther.Value = GetSetting(App.EXEName, "Options", "OptOOther", False)
    OptOSixMonths.Value = GetSetting(App.EXEName, "Options", "OptOSixMonths", False)
    OptOThreeMonths.Value = GetSetting(App.EXEName, "Options", "OptOThreeMonths", False)
    OptOTwoWeeks.Value = GetSetting(App.EXEName, "Options", "OptOTwoWeeks", False)
    OptOTwoMonths.Value = GetSetting(App.EXEName, "Options", "OptOTwoMonths", False)
    OptSelectedItems.Value = GetSetting(App.EXEName, "Options", "OptSelectedItems", False)
    OptSelectedROL.Value = GetSetting(App.EXEName, "Options", "OptSelectedROL", False)
    optAllItems.Value = GetSetting(App.EXEName, "Options", "OptAllItems", False)
    OptAllROL.Value = GetSetting(App.EXEName, "Options", "OptAllROL", False)
    OptUOneMonth.Value = GetSetting(App.EXEName, "Options", "optuOneMonths", True)
    optUOneWeek.Value = GetSetting(App.EXEName, "Options", "optuOneWeek", False)
    OptUOneYear.Value = GetSetting(App.EXEName, "Options", "optuOneYear", False)
    OptUOther.Value = GetSetting(App.EXEName, "Options", "optuOther", False)
    OptUSixMOnths.Value = GetSetting(App.EXEName, "Options", "optuSixMonths", False)
    OptUThreeMonths.Value = GetSetting(App.EXEName, "Options", "optuThreeMonths", False)
    OptUTwoWeeks.Value = GetSetting(App.EXEName, "Options", "optuTwoWeeks", False)
    OptUTwoMonths.Value = GetSetting(App.EXEName, "Options", "optuTwoMonths", False)
    OptSelectedItems.Value = GetSetting(App.EXEName, "Options", "OptSelectedItems", False)
    OptSelectedROL.Value = GetSetting(App.EXEName, "Options", "OptSelectedROL", False)
    chkAdjustments.Value = GetSetting(App.EXEName, "Options", "chkAdjustments", 0)
    chkSales.Value = GetSetting(App.EXEName, "Options", "chkSales", 1)
    chkDiscard.Value = GetSetting(App.EXEName, "Options", "chkDiscard", 0)
    chkConsumption.Value = GetSetting(App.EXEName, "Options", "chkConsumption", 1)
    chkCatogery.Value = GetSetting(App.EXEName, "Options", "chkCatogery", 0)
    chkIgnoreOrdersPending.Value = GetSetting(App.EXEName, "Options", "chkIgnoreOrdersPending", 1)
    OptOneDept.Value = GetSetting(App.EXEName, "Options", "OptOneDept", False)
    OptAllDept.Value = GetSetting(App.EXEName, "Options", "OptAllDept", True)
End Sub

Private Sub CalculatePeriods()
    Dim SUsage As Date
    Dim EUsage As Date
    Dim SOrder As Date
    Dim EOrder As Date
    If optUOneWeek.Value = True Then
        SUsage = Date - 7
    ElseIf OptUTwoWeeks.Value = True Then
        SUsage = Date - 14
    ElseIf OptUOneMonth.Value = True Then
        SUsage = DateSerial(Year(Date), Month(Date) - 1, Day(Date))
        If Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            SUsage = DateSerial(Year(Date), Month(Date), 1)
        ElseIf DateSerial(Year(Date), Month(SUsage), 1) = DateSerial(Year(Date), Month(Date), 1) Then
            SUsage = DateSerial(Year(Date), Month(Date), 1) - 1
        End If
    ElseIf OptUTwoMonths.Value = True Then
        SUsage = DateSerial(Year(Date), Month(Date) - 2, Day(Date))
        If Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            SUsage = DateSerial(Year(Date), Month(Date) - 1, 1)
        ElseIf DateSerial(Year(Date), Month(SUsage) - 1, 1) = DateSerial(Year(Date), Month(Date) - 1, 1) Then
            SUsage = DateSerial(Year(Date), Month(Date) - 1, 1) - 1
        End If
    ElseIf OptUThreeMonths.Value = True Then
        SUsage = DateSerial(Year(Date), Month(Date) - 3, Day(Date))
        If Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            SUsage = DateSerial(Year(Date), Month(Date) - 2, 1)
        ElseIf DateSerial(Year(Date), Month(SUsage) - 2, 1) = DateSerial(Year(Date), Month(Date) - 2, 1) Then
            SUsage = DateSerial(Year(Date), Month(Date) - 2, 1) - 1
        End If
    ElseIf OptUSixMOnths.Value = True Then
        SUsage = DateSerial(Year(Date), Month(Date) - 6, Day(Date))
        If Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            SUsage = DateSerial(Year(Date), Month(Date) - 5, 1)
        ElseIf DateSerial(Year(Date), Month(SUsage) - 5, 1) = DateSerial(Year(Date), Month(Date) - 5, 1) Then
            SUsage = DateSerial(Year(Date), Month(Date) - 5, 1) - 1
        End If
    ElseIf OptUOneYear.Value = True Then
        SUsage = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
        If Month(Date) = 2 And Day(Date) = 29 Then
            SUsage = DateSerial(Year(Date) - 1, Month(Date), 28)
        End If
    ElseIf OptUOther.Value = True Then
    
    End If
    
    If OptOOneWeek.Value = True Then
        EOrder = Date + 7
    ElseIf OptOTwoWeeks.Value = True Then
        EOrder = Date + 14
    ElseIf OptOOneMonths.Value = True Then
        EOrder = DateSerial(Year(Date), Month(Date) + 1, Day(Date) - 1)
        If Date = DateSerial(Year(Date), Month(Date), 1) Then
            EOrder = DateSerial(Year(Date), Month(Date) + 1, 1) - 1
        ElseIf Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            EOrder = DateSerial(Year(Date), Month(Date) + 2, 1) - 1
        End If
    ElseIf OptOTwoMonths.Value = True Then
        EOrder = DateSerial(Year(Date), Month(Date) + 2, Day(Date) - 1)
        If Date = DateSerial(Year(Date), Month(Date), 1) Then
            EOrder = DateSerial(Year(Date), Month(Date) + 2, 1) - 1
        ElseIf Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            EOrder = DateSerial(Year(Date), Month(Date) + 3, 1) - 1
        End If
    ElseIf OptOThreeMonths.Value = True Then
        EOrder = DateSerial(Year(Date), Month(Date) + 3, Day(Date) - 1)
        If Date = DateSerial(Year(Date), Month(Date), 1) Then
            EOrder = DateSerial(Year(Date), Month(Date) + 3, 1) - 1
        ElseIf Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            EOrder = DateSerial(Year(Date), Month(Date) + 4, 1) - 1
        End If
    ElseIf OptOSixMonths.Value = True Then
        EOrder = DateSerial(Year(Date), Month(Date) + 6, Day(Date) - 1)
        If Date = DateSerial(Year(Date), Month(Date), 1) Then
            EOrder = DateSerial(Year(Date), Month(Date) + 6, 1) - 1
        ElseIf Date = DateSerial(Year(Date), Month(Date) + 1, 1) - 1 Then
            EOrder = DateSerial(Year(Date), Month(Date) + 7, 1) - 1
        End If
    ElseIf OptOOneYear.Value = True Then
        EOrder = DateSerial(Year(Date) + 1, Month(Date), Day(Date) - 1)
        If Month(Date) = 2 And Day(Date) = 29 Then
            EOrder = DateSerial(Year(Date) + 1, Month(Date), 28)
        End If
    ElseIf OptOOther.Value = True Then
    
    End If
    
    EUsage = Date
    SOrder = Date
    
    mvUsageFrom.Value = SUsage
    mvUsageTo.Value = EUsage
    mvOrderFrom.Value = SOrder
    mvOrderTo.Value = EOrder
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, "Options", "OptAllItems", optAllItems.Value
    SaveSetting App.EXEName, "Options", "OptAllROL", OptAllROL.Value
    SaveSetting App.EXEName, "Options", "OptOOneMonths", OptOOneMonths.Value
    SaveSetting App.EXEName, "Options", "OptOOneWeek", OptOOneWeek.Value
    SaveSetting App.EXEName, "Options", "OptOOneYear", OptOOneYear.Value
    SaveSetting App.EXEName, "Options", "OptOOther", OptOOther.Value
    SaveSetting App.EXEName, "Options", "OptOSixMonths", OptOSixMonths.Value
    SaveSetting App.EXEName, "Options", "OptOThreeMonths", OptOThreeMonths.Value
    SaveSetting App.EXEName, "Options", "OptOTwoWeeks", OptOTwoWeeks.Value
    SaveSetting App.EXEName, "Options", "OptOTwoMonths", OptOTwoMonths.Value
    SaveSetting App.EXEName, "Options", "OptSelectedItems", OptSelectedItems.Value
    SaveSetting App.EXEName, "Options", "OptSelectedROL", OptSelectedROL.Value
    SaveSetting App.EXEName, "Options", "OptAllItems", optAllItems.Value
    SaveSetting App.EXEName, "Options", "OptAllROL", OptAllROL.Value
    SaveSetting App.EXEName, "Options", "optuOneMonths", OptUOneMonth.Value
    SaveSetting App.EXEName, "Options", "optuOneWeek", optUOneWeek.Value
    SaveSetting App.EXEName, "Options", "optuOneYear", OptUOneYear.Value
    SaveSetting App.EXEName, "Options", "optuOther", OptUOther.Value
    SaveSetting App.EXEName, "Options", "optuSixMonths", OptUSixMOnths.Value
    SaveSetting App.EXEName, "Options", "optuThreeMonths", OptUThreeMonths.Value
    SaveSetting App.EXEName, "Options", "optuTwoWeeks", OptUTwoWeeks.Value
    SaveSetting App.EXEName, "Options", "optuTwoMonths", OptUTwoMonths.Value
    SaveSetting App.EXEName, "Options", "OptSelectedItems", OptSelectedItems.Value
    SaveSetting App.EXEName, "Options", "OptSelectedROL", OptSelectedROL.Value
    SaveSetting App.EXEName, "Options", "chkAdjustments", chkAdjustments.Value
    SaveSetting App.EXEName, "Options", "chkSales", chkSales.Value
    SaveSetting App.EXEName, "Options", "chkDiscard", chkDiscard.Value
    SaveSetting App.EXEName, "Options", "chkConsumption", chkConsumption.Value
    SaveSetting App.EXEName, "Options", "chkCatogery", chkCatogery.Value
    SaveSetting App.EXEName, "Options", "chkIgnoreOrdersPending", chkIgnoreOrdersPending.Value
    SaveSetting App.EXEName, "Options", "OptOneDept", OptOneDept.Value
    SaveSetting App.EXEName, "Options", "OptAllDept", OptAllDept.Value
End Sub

Private Sub lstItem_Click()
    lstItemID.Selected(lstItem.ListIndex) = lstItem.Selected(lstItem.ListIndex)
End Sub

Private Sub lstSelectedItem_Click()
    If lstSelectedItem.ListCount = 0 Then Exit Sub
    lstSelectedItemID.ListIndex = lstSelectedItem.ListIndex
End Sub

Private Sub mvOrderFrom_DateClick(ByVal DateClicked As Date)
    OptOOther.Value = True
End Sub

Private Sub mvOrderTo_DateClick(ByVal DateClicked As Date)
    OptOOther.Value = True
End Sub

Private Sub mvUsageFrom_DateClick(ByVal DateClicked As Date)
    OptUOther.Value = True
End Sub

Private Sub mvUsageTo_DateClick(ByVal DateClicked As Date)
    OptUOther.Value = True
End Sub

Private Sub OptAllItems_Click()
If optAllItems.Value = True Then
    lstItem.Enabled = False
    lstSelectedItem.Enabled = False
Else
    lstItem.Enabled = True
    lstSelectedItem.Enabled = True
End If
End Sub

Private Sub OptAllROL_Click()
If OptAllROL.Value = True Then
    lstItem.Enabled = False
    lstSelectedItem.Enabled = False
Else
    lstItem.Enabled = True
    lstSelectedItem.Enabled = True
End If
End Sub

Private Sub optoOneMonth_Click()
    Call CalculatePeriods
End Sub

Private Sub OptOOneMonths_Click()
    Call CalculatePeriods
End Sub

Private Sub optoOneWeek_Click()
    Call CalculatePeriods
End Sub

Private Sub optoOneYear_Click()
    Call CalculatePeriods
End Sub

Private Sub optoOther_Click()
    Call CalculatePeriods
End Sub

Private Sub optoSixMOnths_Click()
    Call CalculatePeriods
End Sub

Private Sub optoThreeMonths_Click()
    Call CalculatePeriods
End Sub

Private Sub optoTwoMonths_Click()
    Call CalculatePeriods
End Sub

Private Sub optoTwoWeeks_Click()
    Call CalculatePeriods
End Sub

Private Sub OptSelectedItems_Click()
If OptSelectedItems.Value = True Then
    lstItem.Enabled = True
    lstSelectedItem.Enabled = True
Else
    lstItem.Enabled = False
    lstSelectedItem.Enabled = False
End If

End Sub

Private Sub OptSelectedROL_Click()
If OptSelectedROL.Value = True Then
    lstItem.Enabled = True
    lstSelectedItem.Enabled = True
Else
    lstItem.Enabled = False
    lstSelectedItem.Enabled = False
End If

End Sub

Private Sub OptUOneMonth_Click()
    Call CalculatePeriods

End Sub

Private Sub optUOneWeek_Click()
    Call CalculatePeriods
End Sub

Private Sub OptUOneYear_Click()
    Call CalculatePeriods

End Sub

Private Sub OptUSixMOnths_Click()
    Call CalculatePeriods

End Sub

Private Sub OptUThreeMonths_Click()
    Call CalculatePeriods

End Sub

Private Sub OptUTwoMonths_Click()
    Call CalculatePeriods

End Sub

Private Sub OptUTwoWeeks_Click()
    Call CalculatePeriods

End Sub
