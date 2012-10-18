VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Lakmedipro e-Store"
   ClientHeight    =   8820
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10935
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
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
         End
         Begin VB.Menu mnuItemSuppliers 
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
      Begin VB.Menu mnuTransactions 
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
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnuPatients 
         Caption         =   "Patients"
         Begin VB.Menu mnuAdmit 
            Caption         =   "Admit"
         End
         Begin VB.Menu mnuChangeRoom 
            Caption         =   "Change Room"
         End
         Begin VB.Menu mnuEditBHT 
            Caption         =   "Edit BHT Details"
         End
         Begin VB.Menu mnuDischarge 
            Caption         =   "Discharge"
         End
      End
      Begin VB.Menu mnuRooms 
         Caption         =   "Rooms"
      End
      Begin VB.Menu mnuDepartments 
         Caption         =   "Departments"
      End
      Begin VB.Menu mnudistributors 
         Caption         =   "Distributors"
      End
      Begin VB.Menu mnuImporters 
         Caption         =   "Importers"
      End
      Begin VB.Menu mnuManufactures 
         Caption         =   "Manufactures"
      End
   End
   Begin VB.Menu mnuStore 
      Caption         =   "Store"
      Begin VB.Menu mnuSale 
         Caption         =   "Sale"
      End
      Begin VB.Menu mnuOPDSale 
         Caption         =   "OPD Sale"
      End
      Begin VB.Menu mnuGenericNameSale 
         Caption         =   "Generic Name Sale"
      End
      Begin VB.Menu mnuHealthSchemeSale 
         Caption         =   "Health Scheme Sale"
      End
      Begin VB.Menu mnuFill 
         Caption         =   "Refill"
         Begin VB.Menu mnuAutomaticOrdering 
            Caption         =   "Automatic Ordering"
         End
         Begin VB.Menu mnuGoodReceive 
            Caption         =   "Good Receive"
         End
         Begin VB.Menu mnuPurchase 
            Caption         =   "Purchase"
         End
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "Return"
      End
      Begin VB.Menu mnuSalesCancellation 
         Caption         =   "Cancelllation"
      End
      Begin VB.Menu mnuConsume 
         Caption         =   "Consume"
      End
      Begin VB.Menu mnuTransfers 
         Caption         =   "Transfers"
      End
      Begin VB.Menu mnuReceive 
         Caption         =   "Receive"
      End
      Begin VB.Menu mnuDiscard 
         Caption         =   "Discard"
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
      End
      Begin VB.Menu mnuIncome 
         Caption         =   "Income"
      End
      Begin VB.Menu mnuExpence 
         Caption         =   "Expence"
      End
   End
   Begin VB.Menu mnuBackOffice 
      Caption         =   "Back Office"
      Begin VB.Menu mnuApproveOrders 
         Caption         =   "Approve Orders"
      End
      Begin VB.Menu mnuShiftEndSummeries 
         Caption         =   "Shift End Summeries"
         Begin VB.Menu mnuShiftEndCashSummery 
            Caption         =   "Shift End Cash Summery"
         End
         Begin VB.Menu mnuShiftEndCreditSummery 
            Caption         =   "Shift End Credit Summery"
         End
         Begin VB.Menu mnuShiftEndChequeSummery 
            Caption         =   "Shift End Cheque Summery"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuShiftEndCreditCardSummery 
            Caption         =   "Shift End Credit Card Summery"
         End
         Begin VB.Menu mnuShiftEndWaucherSummery 
            Caption         =   "Shift End Waucher Summery"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuDayEndSummeries 
         Caption         =   "Day End Summeries"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuSaleReports 
            Caption         =   "Sale Reports"
            Begin VB.Menu mnuTotalSaleReport 
               Caption         =   "Total Sale Report"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCashSale 
               Caption         =   "Cash Sale"
            End
            Begin VB.Menu mnuCreditSale 
               Caption         =   "Credit Sale"
            End
            Begin VB.Menu mnuChequeSale 
               Caption         =   "Cheque Sale"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuCreditCardSale 
               Caption         =   "Credit Card Sale"
            End
            Begin VB.Menu mnuOutPatientSale 
               Caption         =   "Out Patient Sale"
            End
            Begin VB.Menu mnuInpatientSale 
               Caption         =   "Inpatient Sale"
            End
            Begin VB.Menu mnuStaffSale 
               Caption         =   "Staff Sale"
            End
         End
         Begin VB.Menu mnuSaleCategoryReports 
            Caption         =   "Sale Category Reports"
         End
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
         End
         Begin VB.Menu mnuFastMovingItems 
            Caption         =   "Fast Moving Items"
         End
         Begin VB.Menu mnuSlowMovingItems 
            Caption         =   "Slow Moving Items"
         End
         Begin VB.Menu mnuNonMovingItems 
            Caption         =   "Non Moving Items"
         End
      End
      Begin VB.Menu mnuDeleteRecords 
         Caption         =   "Delete Records"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Begin VB.Menu mnuMyShiftEndSummery 
         Caption         =   "My Shift-End Summery"
      End
      Begin VB.Menu mnuTimeWiseSummery 
         Caption         =   "My Timewise Summery"
      End
      Begin VB.Menu mnuMyDayEndSummery 
         Caption         =   "My Day-End Summery"
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
   Begin VB.Menu mnuTem 
      Caption         =   "tem"
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    If UserAuthority = 6 Then
        mnuBackOffice.Visible = False
        mnuStaff.Enabled = False
    Else
        mnuBackOffice.Enabled = True
        mnuStaff.Enabled = True
    End If
    mnuStore.Caption = UserStore
    mnuReturn.Enabled = False
End Sub

Private Sub mnuAdjustmentCatogeries_Click()
    frmAdjustmentCategories.Show
End Sub

Private Sub mnuAdmit_Click()
    frmAdmit.Show
End Sub

Private Sub mnuApproveOrders_Click()
    frmApproveOrderSelection.Show
End Sub

Private Sub mnuAutomaticOrdering_Click()
    frmAutoOrdering.Show
End Sub

Private Sub mnuBackup_Click()
    frmBackUp.Show
End Sub

Private Sub mnuBatch_Click()
    frmEditBatch.Show
End Sub

Private Sub mnuBHT_Click()
    frmBHT.Show
End Sub

Private Sub mnuCashSale_Click()
    frmCashSale.Show
End Sub

Private Sub mnuChequeSale_Click()
    frmChequeSale.Show
End Sub

Private Sub mnuConsume_Click()
    frmConsumption.Show
End Sub

Private Sub mnuConsumptionCatogeries_Click()
    frmConsumptionCatogeries.Show
End Sub

Private Sub mnuConsumptionReport_Click()
    frmConsumptionReport.Show
End Sub

Private Sub mnuCreditCardSale_Click()
    frmCreditCardSale.Show
End Sub

Private Sub mnuCreditSale_Click()
    frmCreditSale.Show
End Sub

Private Sub mnuCurrentStock_Click()
    frmCurrentStock.Show
End Sub

Private Sub mnuDeleteRecords_Click()
    frmDeletePastRecords.Show
    frmDeletePastRecords.ZOrder 0
End Sub

Private Sub mnuDepartments_Click()
    frmDepartments.Show
End Sub

Private Sub mnuDiscard_Click()
    frmDiscard.Show
End Sub

Private Sub mnuDiscardCatogeries_Click()
    frmDiscardCategories.SetFocus
End Sub

Private Sub mnuDischarge_Click()
    frmDischarge.Show
End Sub

Private Sub mnudistributors_Click()
    frmDistributers.Show
End Sub

Private Sub mnuDurations_Click()
    frmDuration.Show
End Sub

Private Sub mnuEditBHT_Click()
    frmEditBHT.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExpenceCategory_Click()
    frmExpenceCategories.Show
End Sub

Private Sub mnuFrequencies_Click()
    frmFrequency.Show
End Sub

Private Sub mnuGenericNames_Click()
    frmGenericNames.Show
End Sub

Private Sub mnuGenericNameSale_Click()
    frmGenericNames.Show
    frmGenericSale.ZOrder 0
End Sub

Private Sub mnuGoodReceive_Click()
    frmGoodReceiveSelection.Show
End Sub

Private Sub mnuHealthSchemeSale_Click()
    frmHighSale.Show
End Sub

Private Sub mnuHospitalDetails_Click()
    frmHospital.Show
End Sub

Private Sub mnuImporters_Click()
    frmImporters.Show
End Sub

Private Sub mnuIncomeCategory_Click()
    frmIncomeCategories.Show
End Sub

Private Sub mnuIssueUnits_Click()
    frmIssueUnits.Show
End Sub

Private Sub mnuItemCatogeries_Click()
    frmItemCatogeries.Show
End Sub

Private Sub mnuItemMaster_Click()
    frmItemMaster.Show
End Sub

Private Sub mnuItemSuppliers_Click()
    frmItemSuppliers.Show
End Sub

Private Sub mnuManufactures_Click()
    frmManufacturers.Show
End Sub

Private Sub mnuMessages_Click()
    frmPMessage.Show
End Sub

Private Sub mnuOPDSale_Click()
    frmOPDSale.Show
End Sub

Private Sub mnuPackUnits_Click()
    frmPackUnits.Show
End Sub

Private Sub mnuPatientDetails_Click()
    frmPatientsDetails.Show
End Sub

Private Sub mnuPatients_Click()
'    frmPatients.Show
End Sub

Private Sub mnuPrescreptionSale_Click()
    frmPrescreptionSale.Show
End Sub

Private Sub mnuPrintingPreferances_Click()
    frmPrintingPreferances.Show
End Sub

Private Sub mnuProgramPreferances_Click()
    frmProgramPreferance.Show
End Sub

Private Sub mnuPurchase_Click()
    frmPurchase.Show
End Sub

Private Sub mnuPurchasePriceAdjustment_Click()
    frmPurchasePriceChange.Show
End Sub

Private Sub mnuReceive_Click()
    frmReceiveItems.Show
End Sub

Private Sub mnuRestore_Click()
    frmRestore.Show
End Sub

Private Sub mnuReturn_Click()
    frmBillSearchForReturn.Show
End Sub

Private Sub mnuRooms_Click()
    frmRoom.Show
End Sub

Private Sub mnuSale_Click()
    frmSale.Show
End Sub

Private Sub mnuSaleCategoryReportsCash_Click()
    frmCashSaleCategory.Show
End Sub

Private Sub mnuSaleCategoryReportsCredit_Click()
    frmCreditSaleCategory.Show
End Sub

Private Sub mnuSaleCategoryReports_Click()
    frmCategorySaleReport.Show
End Sub

Private Sub mnuSaleCatogeries_Click()
    frmSaleCatogeries.Show
End Sub


Private Sub mnuSalesCancellation_Click()
    frmBillSearchForCancellations.Show
End Sub

Private Sub mnuSalespriceAjustments_Click()
    frmSalesPriceChange.Show
End Sub


Private Sub mnuShiftEndCashSummery_Click()
    frmShiftEndCashSale.Show
End Sub

Private Sub mnuShiftEndCreditCardSummery_Click()
    frmShiftEndCreditCardSale.Show
End Sub

Private Sub mnuShiftEndCreditSummery_Click()
    frmShiftEndCreditSale.Show
End Sub

Private Sub mnuStaff_Click()
    frmStaff.Show
End Sub

Private Sub mnuStaffSale_Click()
    frmstaffSale.Show
End Sub

Private Sub mnuStockAjustments_Click()
    frmStockAdjustment.Show
End Sub

Private Sub mnuStrengthUnits_Click()
    frmStrengthUnits.Show
End Sub

Private Sub mnuTimeWiseSummery_Click()
    frmCategorySaleReportWithTimeMatara.Show
    frmCategorySaleReportWithTimeMatara.ZOrder 0
End Sub

Private Sub mnuTotalSaleReport_Click()
    frmSaleReports.Show
End Sub

Private Sub mnuTradeNames_Click()
    frmTradeNames.Show
End Sub

Private Sub mnuTransfer_Click()
    frmTransferReport.Show
End Sub

Private Sub mnuTransferCatogeries_Click()
    frmTransferCatogeries.Show
End Sub

Private Sub mnuTransfers_Click()
    frmTransfer.Show
End Sub
