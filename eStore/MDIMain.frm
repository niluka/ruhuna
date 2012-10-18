VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Lakmedipro e-Store"
   ClientHeight    =   8820
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15120
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   7080
      Top             =   4200
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":29C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":302C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":38C75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":409BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":48BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4E3AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":55B28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   9960
         TabIndex        =   1
         Top             =   0
         Width           =   5295
         Begin VB.Label lblDateTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   4695
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5C800
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":62EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":6B863
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":735AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":7B7B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":80F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":88716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuMedicines 
         Caption         =   "Medicines"
         Begin VB.Menu mnuGenericNameS 
            Caption         =   "Generic Names"
         End
         Begin VB.Menu mnuTradeNames 
            Caption         =   "Trade Names"
         End
         Begin VB.Menu mnuItemCatogeries 
            Caption         =   "Item Catogeries"
         End
         Begin VB.Menu mnuItemMaster 
            Caption         =   "Item Master"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuReportItemSuppliers 
            Caption         =   "Item Suppliers"
         End
         Begin VB.Menu mnuBatch 
            Caption         =   "Batch"
         End
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "Units"
         Begin VB.Menu mnuStrengthUnits 
            Caption         =   "Strength Units"
         End
         Begin VB.Menu mnuIssueUnits 
            Caption         =   "Issue Units"
         End
         Begin VB.Menu mnuPackUnits 
            Caption         =   "Pack Units"
         End
         Begin VB.Menu mnuDoseUnits 
            Caption         =   "Dose Units"
         End
         Begin VB.Menu mnuFrequencies 
            Caption         =   "Frequencies"
         End
         Begin VB.Menu mnuDurations 
            Caption         =   "Durations"
         End
         Begin VB.Menu mnuMessages 
            Caption         =   "Messages"
         End
      End
      Begin VB.Menu mnuCategories 
         Caption         =   "Catogeries"
         Begin VB.Menu mnuRefillCatogeries 
            Caption         =   "Refill Catogeries"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSaleCatogeries 
            Caption         =   "Sale Catogeries"
         End
         Begin VB.Menu mnuTransferCatogeries 
            Caption         =   "Transfer Catogeries"
         End
         Begin VB.Menu mnuConsumptionCatogeries 
            Caption         =   "Consumption Catogeries"
         End
         Begin VB.Menu mnuDiscardCatogeries 
            Caption         =   "Discard Catogeries"
         End
         Begin VB.Menu mnuAdjustmentCatogeries 
            Caption         =   "Adjustment Catogeries"
         End
         Begin VB.Menu mnuIncomeCategory 
            Caption         =   "Income Categories"
         End
         Begin VB.Menu mnuExpenceCategory 
            Caption         =   "Expence Categories"
         End
      End
      Begin VB.Menu mnuHospital 
         Caption         =   "Hospital"
         Begin VB.Menu mnuStaff 
            Caption         =   "Staff"
         End
         Begin VB.Menu mnuDepartments 
            Caption         =   "Departments"
         End
         Begin VB.Menu mnuRooms 
            Caption         =   "Rooms"
         End
      End
      Begin VB.Menu mnuDistributorS 
         Caption         =   "Distributors"
      End
      Begin VB.Menu mnuImporters 
         Caption         =   "Importers"
      End
      Begin VB.Menu mnuManufactures 
         Caption         =   "Manufactures"
      End
      Begin VB.Menu mnuHealthSchemeSuppliers 
         Caption         =   "Health Scheme Suppliers"
      End
   End
   Begin VB.Menu mnuStore 
      Caption         =   "Store"
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Begin VB.Menu mnuHospitalIssue 
            Caption         =   "Sale"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuSale1 
            Caption         =   "Sale 1"
            Shortcut        =   ^{F3}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSale 
            Caption         =   "Old Sale"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOPDSale 
            Caption         =   "OPD Sale"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuGenericNameSale 
            Caption         =   "Generic Name Sale"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPrescreptionSale 
            Caption         =   "Prescreption Sale"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuReturn 
            Caption         =   "Returns"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuSalesCancellation 
            Caption         =   "Cancelllations"
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Purchase"
         Begin VB.Menu mnuAutomaticOrdering 
            Caption         =   "Ordering"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuPurchase 
            Caption         =   "Good Receive"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuPurchaseCancellations 
            Caption         =   "Cancellations"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuPurchaseReturns 
            Caption         =   "Returns"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuPurchaseReprints 
            Caption         =   "Reprints"
         End
         Begin VB.Menu mnuStoreOrders 
            Caption         =   "Orders"
         End
      End
      Begin VB.Menu mnuTransactions 
         Caption         =   "Transactions"
         Visible         =   0   'False
         Begin VB.Menu mnuConsume 
            Caption         =   "Consume"
         End
         Begin VB.Menu mnuTransfers 
            Caption         =   "Transfers"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuReceive 
            Caption         =   "Receive"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDiscard 
            Caption         =   "Discard"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuAdjustments 
         Caption         =   "Adjustments"
         Begin VB.Menu mnuStockAjustments 
            Caption         =   "Stock Ajustments"
         End
         Begin VB.Menu mnuPurchasePriceAdjustment 
            Caption         =   "Purchase Price Adjustments"
         End
         Begin VB.Menu mnuSalespriceAjustments 
            Caption         =   "Sales Price Ajustments"
         End
         Begin VB.Menu mnuPriceAdjustments 
            Caption         =   "Price Adjustments"
         End
      End
   End
   Begin VB.Menu mnuPatients 
      Caption         =   "Patients"
      Begin VB.Menu mnuAdmit 
         Caption         =   "Admit"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuBackOffice 
      Caption         =   "Back Office"
      Begin VB.Menu mnuApproveOrders 
         Caption         =   "Approve Orders"
      End
      Begin VB.Menu mnuBackOfficePurchaseBillSettling 
         Caption         =   "Purchase Bill Settling"
      End
      Begin VB.Menu mnuSaleReports 
         Caption         =   "Sale Reports"
         Begin VB.Menu mnuSaleShiftEndSummery 
            Caption         =   "Shift End Summery"
         End
         Begin VB.Menu mnuSaleDayEndSummery 
            Caption         =   "Day End Summery"
         End
         Begin VB.Menu mnuTotalSaleReport 
            Caption         =   "Total Sale Report"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSaleCategoryReports 
            Caption         =   "Sale Category Reports"
         End
         Begin VB.Menu mnuInpatientSale 
            Caption         =   "Inpatient Sale"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuStaffSale 
            Caption         =   "Staff Sale"
         End
      End
      Begin VB.Menu mnuPurchaseReports 
         Caption         =   "Purchase Reports"
         Begin VB.Menu mnuReportPurchaseBills 
            Caption         =   "Purchase Bills"
         End
         Begin VB.Menu mnuPurchaseItems 
            Caption         =   "Purchase Items"
         End
         Begin VB.Menu mnuPurchaseBillSettlements 
            Caption         =   "Purchase Bill Settlements"
         End
      End
      Begin VB.Menu mnuTranactionReports 
         Caption         =   "Transaction Reports"
         Visible         =   0   'False
         Begin VB.Menu mnuDiscardReports 
            Caption         =   "Discard"
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Treansfer"
         End
         Begin VB.Menu mnuConsumptionReport 
            Caption         =   "Consumption"
         End
      End
      Begin VB.Menu mnuStockReports 
         Caption         =   "Stock Reports"
         Begin VB.Menu mnuCurrentStock 
            Caption         =   "Current Stock"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBatchStocks 
            Caption         =   "Batch Stock"
         End
         Begin VB.Menu mnuBackOfficeCategoryStock 
            Caption         =   "Category Stock"
         End
         Begin VB.Menu mnuDistributorStock 
            Caption         =   "Distributor Stock"
         End
         Begin VB.Menu mnuExpiaringStocks 
            Caption         =   "Expiaring Stocks"
         End
         Begin VB.Menu mnuBatchStockBeforeVerification 
            Caption         =   "Batch Stock Before Verification"
         End
         Begin VB.Menu mnuBatchStockAfterVerification 
            Caption         =   "Batch Stock After Verification"
         End
      End
      Begin VB.Menu mnuItemReports 
         Caption         =   "Item Reports"
         Begin VB.Menu mnuItemsDetails 
            Caption         =   "Item Details"
         End
         Begin VB.Menu mnuAllItemIssue 
            Caption         =   "All Item Issues"
         End
         Begin VB.Menu mnuSaleCategoryViceItemIssue 
            Caption         =   "Item Issue By Sale Category"
         End
         Begin VB.Menu mnuItemIssueToUnits 
            Caption         =   "Item Issue to Units"
         End
         Begin VB.Menu mnuItemIssueToStaff 
            Caption         =   "Item Issue to Staff"
         End
         Begin VB.Menu mnuItemIssueToBHT 
            Caption         =   "Item Issue to BHT"
         End
         Begin VB.Menu mnuItemIssueToCustomers 
            Caption         =   "Item Issue to Customers"
         End
         Begin VB.Menu mnuFastAndSlowMovingItems 
            Caption         =   "Fast and Slow Moving Items"
         End
         Begin VB.Menu mnuNonMovingItems 
            Caption         =   "Non Moving Items"
         End
         Begin VB.Menu mnuItemSummery 
            Caption         =   "Item Summery"
         End
         Begin VB.Menu mnuItemSaleGraph 
            Caption         =   "Item Sale Graph"
         End
      End
      Begin VB.Menu mnuBackOfficeOtherReports 
         Caption         =   "Other Reports"
         Begin VB.Menu mnuReportItemDetails 
            Caption         =   "Item Details"
         End
         Begin VB.Menu mnuItemSuppliers 
            Caption         =   "Item Suppliers"
         End
         Begin VB.Menu mnuBackOfficeOtherReportsStockAdjustment 
            Caption         =   "Stock Adjustments"
         End
         Begin VB.Menu mnuDistributorItems 
            Caption         =   "Distributor Items"
         End
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Begin VB.Menu mnuMyShiftEndSaleSummery 
         Caption         =   "My Shift-End Sale Summery"
      End
      Begin VB.Menu mnuMyDayEndSaleSummery 
         Caption         =   "My Day-End Sale Summery"
      End
      Begin VB.Menu mnuMyShiftEndPurchaseSummery 
         Caption         =   "My Shift-End Purchase Summery"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMyDayEndPurchaseSummery 
         Caption         =   "My Day-End Purchase Summery"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPreferances 
      Caption         =   "Preferances"
      Begin VB.Menu mnuProgramPreferances 
         Caption         =   "Program Preferances"
      End
      Begin VB.Menu mnuPrintingPreferances 
         Caption         =   "Printing Preferances"
      End
      Begin VB.Menu mnuHospitalDetails 
         Caption         =   "Hospital Details"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTipOfTheDay 
         Caption         =   "Tip of the day"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuTableOfContants 
         Caption         =   "Table of Contents"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

  '  frmTem.Show
    

    If UserAuthority = 6 Then
        mnuBackOffice.Visible = False
        mnuStaff.Visible = False
        mnuAdjustments.Visible = False
        mnuHelp.Visible = False
        mnuHospitalDetails.Visible = False
    Else
        mnuBackOffice.Enabled = True
        mnuStaff.Visible = True
        mnuAdjustments.Visible = True
        mnuHelp.Visible = True
        mnuHospitalDetails.Visible = True
    End If
    mnuStore.Caption = UserStore
    lblDateTime.Caption = Format("Date : " & Format(Date, "dd MMMM yyyy") & "   Time : " & Format(Time, "H:M AMPM"))
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    i = MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit?")
    If i = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub mnuAbout_Click()
    frmTem.Show
    frmTem.ZOrder 0
End Sub

Private Sub mnuAdjustmentCatogeries_Click()
    frmAdjustmentCategories.Show
    frmAdjustmentCategories.ZOrder 0
    frmAdjustmentCategories.Top = 0
    frmAdjustmentCategories.Left = 0
End Sub

Private Sub mnuAdmit_Click()
    frmAdmit.Show
    frmAdmit.ZOrder 0
    frmAdmit.Top = 0
    frmAdmit.Left = 0
End Sub

Private Sub mnuAllItemIssue_Click()
    frmAllItemIssue.Show
    frmAllItemIssue.ZOrder 0
    frmAllItemIssue.Top = 0
    frmAllItemIssue.Left = 0
End Sub

Private Sub mnuApproveOrders_Click()
    frmApproveOrderSelection.Show
    frmApproveOrderSelection.ZOrder 0
    frmApproveOrderSelection.Top = 0
    frmApproveOrderSelection.Left = 0
End Sub

Private Sub mnuAutomaticOrdering_Click()
    frmAutoOrderingNew.Show
    frmAutoOrderingNew.ZOrder 0
    frmAutoOrderingNew.Top = 0
    frmAutoOrderingNew.Left = 0
End Sub

Private Sub mnuBackOfficeCategoryStock_Click()
    frmCategoryBatchStock.Show
    frmCategoryBatchStock.ZOrder 0
End Sub

Private Sub mnuBackOfficeOtherReportsStockAdjustment_Click()
    frmReportStockAdjustment.Show
    frmReportStockAdjustment.ZOrder 0
    frmReportStockAdjustment.Top = 0
    frmReportStockAdjustment.Left = 0
End Sub

Private Sub mnuBackOfficePurchaseBillSettling_Click()
    frmPruchaseBillSettling.Show
    frmPruchaseBillSettling.ZOrder 0
End Sub

Private Sub mnuBackup_Click()
    frmBackUp.Show
    frmBackUp.ZOrder 0
    frmBackUp.Top = 0
    frmBackUp.Left = 0
End Sub

Private Sub mnuBatch_Click()
    frmEditBatch.Show
    frmEditBatch.ZOrder 0
    frmEditBatch.Top = 0
    frmEditBatch.Left = 0
End Sub

Private Sub mnuBatchStockAfterVerification_Click()
    frmBatchStockAfterVarification.Show
    frmBatchStockAfterVarification.ZOrder 0
End Sub

Private Sub mnuBatchStockBeforeVerification_Click()
    frmBatchStockBeforeVarification.Show
    frmBatchStockBeforeVarification.ZOrder 0
End Sub

Private Sub mnuBatchStocks_Click()
    frmBatchStock.Show
    frmBatchStock.ZOrder 0
    frmBatchStock.Top = 0
    frmBatchStock.Left = 0
End Sub

Private Sub mnuConsume_Click()
    frmConsumption.Show
    frmConsumption.ZOrder 0
    frmConsumption.Top = 0
    frmConsumption.Left = 0
End Sub

Private Sub mnuConsumptionCatogeries_Click()
    frmConsumptionCatogeries.Show
    frmConsumptionCatogeries.ZOrder 0
    frmConsumptionCatogeries.Top = 0
    frmConsumptionCatogeries.Left = 0
End Sub

Private Sub mnuConsumptionReport_Click()
    frmConsumptionReport.Show
    frmConsumptionReport.ZOrder 0
    frmConsumptionReport.Top = 0
    frmConsumptionReport.Left = 0
End Sub

Private Sub mnuCurrentStock_Click()
    frmCurrentStock.Show
    frmCurrentStock.ZOrder 0
    frmCurrentStock.Top = 0
    frmCurrentStock.Left = 0
End Sub

Private Sub mnuDepartments_Click()
    frmDepartments.Show
    frmDepartments.ZOrder 0
    frmDepartments.Top = 0
    frmDepartments.Left = 0
End Sub

Private Sub mnuDiscard_Click()
    frmDiscard.Show
    frmDiscard.ZOrder 0
    frmDiscard.Top = 0
    frmDiscard.Left = 0
End Sub

Private Sub mnuDiscardCatogeries_Click()
    frmDiscardCategories.SetFocus
    frmDiscardCategories.ZOrder 0
    frmDiscardCategories.Top = 0
    frmDiscardCategories.Left = 0
End Sub

Private Sub mnuDistributorItems_Click()
    frmDistributorItems.Show
    frmDistributorItems.ZOrder 0
    frmDistributorItems.Top = 0
    frmDistributorItems.Left = 0
End Sub

Private Sub mnudistributors_Click()
    frmDistributers.Show
    frmDistributers.ZOrder 0
    frmDistributers.Top = 0
    frmDistributers.Left = 0
End Sub

Private Sub mnuDItems_Click()
    frmDistributorItems.Show
    frmDistributorItems.ZOrder 0
    frmDistributorItems.Top = 0
    frmDistributorItems.Left = 0
End Sub

Private Sub mnuDistributorStock_Click()
    frmSupplierBatchStock.Show
    frmSupplierBatchStock.ZOrder 0
End Sub

Private Sub mnuDoseUnits_Click()
    frmDoseUnit.Show
    frmDoseUnit.ZOrder 0
    frmDoseUnit.Top = 0
    frmDoseUnit.Left = 0
End Sub

Private Sub mnuDurations_Click()
    frmDuration.Show
    frmDuration.ZOrder 0
    frmDuration.Top = 0
    frmDuration.Left = 0
End Sub

Private Sub mnuEditBHT_Click()
    frmEditBHT.Show
    frmEditBHT.ZOrder 0
    frmEditBHT.Top = 0
    frmEditBHT.Left = 0
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExpenceCategory_Click()
    frmExpenceCategories.Show
    frmExpenceCategories.ZOrder 0
    frmExpenceCategories.Top = 0
    frmExpenceCategories.Left = 0
End Sub

Private Sub mnuExpiaringStocks_Click()
    frmBatchStockExpired.Show
    frmBatchStockExpired.ZOrder 0
End Sub

Private Sub mnuFastAndSlowMovingItems_Click()
    frmAllItemMoving.Show
    frmAllItemMoving.ZOrder 0
    frmAllItemMoving.Top = 0
    frmAllItemMoving.Left = 0
End Sub

Private Sub mnuFrequencies_Click()
    frmFrequency.Show
    frmFrequency.ZOrder 0
    frmFrequency.Top = 0
    frmFrequency.Left = 0
End Sub

Private Sub mnuGenericNames_Click()
    frmGenericNames.Show
    frmGenericNames.ZOrder 0
    frmGenericNames.Top = 0
    frmGenericNames.Left = 0
End Sub

Private Sub mnuGenericNameSale_Click()
    frmGenericSale.Show
    frmGenericSale.ZOrder 0
    frmGenericSale.Top = 0
    frmGenericSale.Left = 0
End Sub

Private Sub mnuGoodReceive_Click()
    frmGoodReceiveSelection.Show
    frmGoodReceiveSelection.ZOrder 0
    frmGoodReceiveSelection.Top = 0
    frmGoodReceiveSelection.Left = 0
End Sub

Private Sub mnuHealthSchemeSuppliers_Click()
    frmHealthSchemeSuppliers.Show
    frmHealthSchemeSuppliers.ZOrder 0
    frmHealthSchemeSuppliers.Top = 0
    frmHealthSchemeSuppliers.Left = 0
End Sub

Private Sub mnuHospitalDetails_Click()
    frmHospital.Show
    frmHospital.ZOrder 0
    frmHospital.Top = 0
    frmHospital.Left = 0
End Sub

Private Sub mnuHospitalIssue_Click()
    frmHospitalSale.Show
    frmHospitalSale.ZOrder 0
End Sub

Private Sub mnuIDetails_Click()
    frmItemDetails.Show
    frmItemDetails.ZOrder 0
    frmItemDetails.Top = 0
    frmItemDetails.Left = 0
End Sub

Private Sub mnuImporters_Click()
    frmImporters.Show
    frmImporters.ZOrder 0
    frmImporters.Top = 0
    frmImporters.Left = 0
End Sub

Private Sub mnuIncomeCategory_Click()
    frmIncomeCategories.Show
    frmIncomeCategories.ZOrder 0
    frmIncomeCategories.Top = 0
    frmIncomeCategories.Left = 0
End Sub

Private Sub mnuIssueUnits_Click()
    frmIssueUnits.Show
    frmIssueUnits.ZOrder 0
    frmIssueUnits.Top = 0
    frmIssueUnits.Left = 0
End Sub

Private Sub mnuISuppliers_Click()
    frmReportItemSuppliers.Show
    frmReportItemSuppliers.ZOrder 0
    frmReportItemSuppliers.Top = 0
    frmReportItemSuppliers.Left = 0
End Sub

Private Sub mnuItemCatogeries_Click()
    frmItemCatogeries.Show
    frmItemCatogeries.ZOrder 0
    frmItemCatogeries.Top = 0
    frmItemCatogeries.Left = 0
End Sub


Private Sub mnuItemIssueToBHT_Click()
    frmAllItemBHTIssue.Show
    frmAllItemBHTIssue.ZOrder 0
    frmAllItemBHTIssue.Top = 0
    frmAllItemBHTIssue.Left = 0
End Sub

Private Sub mnuItemIssueToCustomers_Click()
    frmAllItemCustomerIssue.Show
    frmAllItemCustomerIssue.ZOrder 0
    frmAllItemCustomerIssue.Top = 0
    frmAllItemCustomerIssue.Left = 0
End Sub

Private Sub mnuItemIssueToStaff_Click()
    frmAllItemStaffIssue.Show
    frmAllItemStaffIssue.ZOrder 0
    frmAllItemStaffIssue.Top = 0
    frmAllItemStaffIssue.Left = 0
End Sub

Private Sub mnuItemIssueToUnits_Click()
    frmAllItemUnitIssue.Show
    frmAllItemUnitIssue.ZOrder 0
    frmAllItemUnitIssue.Top = 0
    frmAllItemUnitIssue.Left = 0
End Sub

Private Sub mnuItemMaster_Click()
    frmItemMaster.Show
    frmItemMaster.ZOrder 0
    frmItemMaster.Top = 0
    frmItemMaster.Left = 0
End Sub

Private Sub mnuItemSaleGraph_Click()
    frmReportsItemSale.Show
    frmReportsItemSale.ZOrder 0
End Sub

Private Sub mnuItemsDetails_Click()
    frmItemsDetails.Show
    frmItemsDetails.ZOrder 0
    frmItemsDetails.Top = 0
    frmItemsDetails.Left = 0
End Sub

Private Sub mnuItemSummery_Click()
    frmItemSummery.Show
    frmItemSummery.ZOrder 0
End Sub

Private Sub mnuItemSuppliers_Click()
    frmReportItemSuppliers.Show
    frmReportItemSuppliers.ZOrder 0
    frmReportItemSuppliers.Top = 0
    frmReportItemSuppliers.Left = 0
End Sub

Private Sub mnuManufactures_Click()
    frmManufacturers.Show
    frmManufacturers.ZOrder 0
    frmManufacturers.Top = 0
    frmManufacturers.Left = 0
End Sub

Private Sub mnuMessages_Click()
    frmPMessage.Show
    frmPMessage.ZOrder 0
    frmPMessage.Top = 0
    frmPMessage.Left = 0
End Sub

Private Sub mnuMyDayEndSaleSummery_Click()
    frmMyDayEndSummery.Show
    frmMyDayEndSummery.ZOrder 0
    frmMyDayEndSummery.Top = 0
    frmMyDayEndSummery.Left = 0
End Sub

Private Sub mnuMyShiftEndSaleSummery_Click()
    frmMyShiftEndSummery.Show
    frmMyShiftEndSummery.ZOrder 0
    frmMyShiftEndSummery.Top = 0
    frmMyShiftEndSummery.Left = 0
End Sub

Private Sub mnuNonMovingItems_Click()
    frmAllItemNonMoving.Show
    frmAllItemNonMoving.ZOrder 0
    frmAllItemNonMoving.Top = 0
    frmAllItemNonMoving.Left = 0
End Sub

Private Sub mnuPackUnits_Click()
    frmPackUnits.Show
    frmPackUnits.ZOrder 0
    frmPackUnits.Top = 0
    frmPackUnits.Left = 0
End Sub

Private Sub mnuPatientDetails_Click()
    frmPatientsDetails.Show
    frmPatientsDetails.ZOrder 0
    frmPatientsDetails.Top = 0
    frmPatientsDetails.Left = 0
End Sub

Private Sub mnuPrescreptionSale_Click()
    frmPSale.Show
    frmPSale.ZOrder 0
    frmPSale.Top = 0
    frmPSale.Left = 0
End Sub

Private Sub mnuPriceAdjustments_Click()
    frmEditPrices.Show
    frmEditPrices.ZOrder 0
    frmEditPrices.Top = 0
    frmEditPrices.Left = 0
End Sub

Private Sub mnuPrintingPreferances_Click()
    frmPrintingPreferances.Show
    frmPrintingPreferances.ZOrder 0
    frmPrintingPreferances.Top = 0
    frmPrintingPreferances.Left = 0
End Sub

Private Sub mnuProgramPreferances_Click()
    frmProgramPreferance.Show
    frmProgramPreferance.ZOrder 0
    frmProgramPreferance.Top = 0
    frmProgramPreferance.Left = 0
End Sub

Private Sub mnuPurchase_Click()
    On Error Resume Next
    frmPurchaseNew.Show
    frmPurchaseNew.ZOrder 0
    frmPurchaseNew.Top = 0
    frmPurchaseNew.Left = 0
End Sub

Private Sub mnuPurchaseBillSettlements_Click()
    frmPurchaseBillSettlements.Show
    frmPurchaseBillSettlements.ZOrder 0
End Sub

Private Sub mnuPurchaseCancellations_Click()
    frmPurchaseCancellationSelection.Show
    frmPurchaseCancellationSelection.ZOrder 0
    frmPurchaseCancellationSelection.Top = 0
    frmPurchaseCancellationSelection.Left = 0
End Sub

Private Sub mnuPurchasePriceAdjustment_Click()
    frmPurchasePriceChange.Show
    frmPurchasePriceChange.ZOrder 0
    frmPurchasePriceChange.Top = 0
    frmPurchasePriceChange.Left = 0
End Sub

Private Sub mnuPurchaseReprints_Click()
    frmPastPurchase.Show
    frmPastPurchase.ZOrder 0
    frmPastPurchase.Top = 0
    frmPastPurchase.Left = 0
End Sub

Private Sub mnuPurchaseReturns_Click()
    frmPurchaseReturnSelection.Show
    frmPurchaseReturnSelection.ZOrder 0
    frmPurchaseReturnSelection.Top = 0
    frmPurchaseReturnSelection.Left = 0
End Sub

Private Sub mnuReceive_Click()
    frmReceiveItems.Show
    frmReceiveItems.ZOrder 0
    frmReceiveItems.Top = 0
    frmReceiveItems.Left = 0
End Sub

Private Sub mnuReportItemDetails_Click()
    frmItemDetails.Show
    frmItemDetails.ZOrder 0
    frmItemDetails.Top = 0
    frmItemDetails.Left = 0
End Sub

Private Sub mnuReportItemSuppliers_Click()
    frmItemSuppliers.Show
    frmItemSuppliers.ZOrder 0
    frmItemSuppliers.Top = 0
    frmItemSuppliers.Left = 0
End Sub

Private Sub mnuReportPurchaseBills_Click()
    frmReportPurchaseBills.Show
    frmReportPurchaseBills.ZOrder 0
End Sub

Private Sub mnuRestore_Click()
    frmRestore.Show
    frmRestore.ZOrder 0
    frmRestore.Top = 0
    frmRestore.Left = 0
End Sub

Private Sub mnuReturn_Click()
    frmBillSearchForReturn.Show
    frmBillSearchForReturn.ZOrder 0
    frmBillSearchForReturn.Top = 0
    frmBillSearchForReturn.Left = 0
End Sub

Private Sub mnuRooms_Click()
    frmRoom.Show
    frmRoom.ZOrder 0
    frmRoom.Top = 0
    frmRoom.Left = 0
End Sub

Private Sub mnuSale_Click()
'    frmSale.Show
'    frmSale.ZOrder 0
End Sub

Private Sub mnuSale1_Click()
    On Error Resume Next
    frmHospitalSale1.Show
    frmHospitalSale1.ZOrder 0
    frmHospitalSale1.Top = 1
    frmHospitalSale1.Left = 0
End Sub

Private Sub mnuSaleCategoryReports_Click()
    frmCategorySaleReport.Show
    frmCategorySaleReport.ZOrder 0
    frmCategorySaleReport.Top = 0
    frmCategorySaleReport.Left = 0
End Sub

Private Sub mnuSaleCategoryViceItemIssue_Click()
    frmAllItemCategoryIssue.Show
    frmAllItemCategoryIssue.ZOrder 0
        frmAllItemCategoryIssue.Top = 0
    frmAllItemCategoryIssue.Left = 0
End Sub

Private Sub mnuSaleCatogeries_Click()
    frmSaleCatogeries.Show
    frmSaleCatogeries.ZOrder 0
    frmSaleCatogeries.Top = 0
    frmSaleCatogeries.Left = 0
End Sub


Private Sub mnuSaleDayEndSummery_Click()
    frmDayEndSummery.Show
    frmDayEndSummery.ZOrder 0
        frmDayEndSummery.Top = 0
    frmDayEndSummery.Left = 0
End Sub

Private Sub mnuSalesCancellation_Click()
    frmBillSearchForCancellations.Show
    frmBillSearchForCancellations.ZOrder 0
    frmBillSearchForCancellations.Top = 0
    frmBillSearchForCancellations.Left = 0
End Sub

Private Sub mnuSaleShiftEndSummery_Click()
    frmShiftEndSummery.Show
    frmShiftEndSummery.ZOrder 0
    frmShiftEndSummery.Top = 0
    frmShiftEndSummery.Left = 0
End Sub

Private Sub mnuSalespriceAjustments_Click()
    frmSalesPriceChange.Show
    frmSalesPriceChange.ZOrder 0
    frmSalesPriceChange.Top = 0
    frmSalesPriceChange.Left = 0
End Sub

Private Sub mnuStaff_Click()
    frmStaff.Show
    frmStaff.ZOrder 0
    frmStaff.Top = 0
    frmStaff.Left = 0
End Sub

Private Sub mnuStaffSale_Click()
    frmstaffSale.Show
    frmstaffSale.ZOrder 0
    frmstaffSale.Top = 0
    frmstaffSale.Left = 0
End Sub

Private Sub mnuStockAjustments_Click()
    frmStockAdjustment.Show
    frmStockAdjustment.ZOrder 0
        frmStockAdjustment.Top = 0
    frmStockAdjustment.Left = 0
End Sub

Private Sub mnuStoreOrders_Click()
    frmViewOrders.Show
    frmViewOrders.ZOrder 0
End Sub

Private Sub mnuStrengthUnits_Click()
    frmStrengthUnits.Show
    frmStrengthUnits.ZOrder 0
        frmStrengthUnits.Top = 0
    frmStrengthUnits.Left = 0
End Sub

Private Sub mnuTotalSaleReport_Click()
    frmSaleReports.Show
    frmSaleReports.ZOrder 0
    frmSaleReports.Top = 0
    frmSaleReports.Left = 0
End Sub

Private Sub mnuTradeNames_Click()
    frmTradeNames.Show
    frmTradeNames.ZOrder 0
        frmTradeNames.Top = 0
    frmTradeNames.Left = 0
End Sub

Private Sub mnuTransfer_Click()
    frmTransferReport.Show
    frmTransferReport.ZOrder 0
        frmTransferReport.Top = 0
    frmTransferReport.Left = 0
End Sub

Private Sub mnuTransferCatogeries_Click()
    frmTransferCatogeries.Show
    frmTransferCatogeries.ZOrder 0
        frmTransferCatogeries.Top = 0
    frmTransferCatogeries.Left = 0
End Sub

Private Sub mnuTransfers_Click()
    frmTransfer.Show
    frmTransfer.ZOrder 0
        frmTransfer.Top = 0
    frmTransfer.Left = 0
End Sub

Private Sub Timer1_Timer()
    lblDateTime.Caption = Format("Date : " & Format(Date, "dd MMMM yyyy") & "   Time : " & Format(Time, "H:M AMPM"))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnuHospitalIssue_Click
        Case 2: mnuSalesCancellation_Click
        Case 3: mnuReturn_Click
        Case 4: mnuPurchase_Click
        Case 5: mnuMyShiftEndSaleSummery_Click
        Case 6: mnuMyDayEndSaleSummery_Click
        Case 7: mnuExit_Click
    End Select
End Sub
