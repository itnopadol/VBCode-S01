if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ARNotReceiveQueueItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ARNotReceiveQueueItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_AccessProgram]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_AccessProgram]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_AvgItemIssue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_AvgItemIssue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_BPlusDepartment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_BPlusDepartment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CalcDriveInMergeTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CalcDriveInMergeTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CancelDriveInDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CancelDriveInDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CancelReqPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CancelReqPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckAuthorizePromotion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckAuthorizePromotion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckInsertItemCodeChangePrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckInsertItemCodeChangePrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckItemCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckItemCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckItemInRecProduct2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckItemInRecProduct2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckLevelPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckLevelPrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckPayBillByDateTime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckPayBillByDateTime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckPayGoods]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckPayGoods]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckQuePickCenter]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckQuePickCenter]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckQuePickCenterZone]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckQuePickCenterZone]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckQueueMonitorMoreThan]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckQueueMonitorMoreThan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckShelfPrintPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckShelfPrintPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CheckStockItemCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CheckStockItemCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CloseOpenItemMinuteLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CloseOpenItemMinuteLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_CouponData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_CouponData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteChangePriceDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteChangePriceDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteDataPrintLabel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteDataPrintLabel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteDriveInMergeTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteDriveInMergeTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteHoldingBill]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteHoldingBill]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteQueueItemSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteQueueItemSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DeleteRequestConfirm]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DeleteRequestConfirm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_DriveInCheckOutPos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_DriveInCheckOutPos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_GenerateItemChangePriceNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_GenerateItemChangePriceNumber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_GetMaxNoHoldingBill]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_GetMaxNoHoldingBill]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertAuthorizePromotion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertAuthorizePromotion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertBasketUpdateItemPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertBasketUpdateItemPrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertBasketUpdateItemPriceDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertBasketUpdateItemPriceDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertCouponDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertCouponDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertCouponMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertCouponMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDataPrintServer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDataPrintServer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDataQueueManagement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDataQueueManagement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDataQueueManagement1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDataQueueManagement1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDriveInCheckOut]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDriveInCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDriveInCheckOutSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDriveInCheckOutSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDriveInMergeTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDriveInMergeTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDriveInSlip]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDriveInSlip]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertDriveInSlipSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertDriveInSlipSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertHoldingBillDriveIn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertHoldingBillDriveIn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertHoldingBillDriveInSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertHoldingBillDriveInSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertItemReqPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertItemReqPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertLabelTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertLabelTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertLogPrintRunningRes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertLogPrintRunningRes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertNPPrintQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertNPPrintQueue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertOpenItemMinuteLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertOpenItemMinuteLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertOrderPickHoldBill]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertOrderPickHoldBill]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertOrderPickHoldBillSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertOrderPickHoldBillSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPayGoods]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPayGoods]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPayGoodsReserve]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPayGoodsReserve]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPickingDataLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPickingDataLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPickingRequestMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPickingRequestMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPickingRequestSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPickingRequestSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertPrintLabel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertPrintLabel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQuePickCenterDriveInSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQuePickCenterDriveInSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQuePickCenterMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQuePickCenterMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQuePickCenterMasterDriveIn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQuePickCenterMasterDriveIn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQuePickCenterSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQuePickCenterSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQueueManagementSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQueueManagementSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQueueManagementSub1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQueueManagementSub1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQueueManagementSub2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQueueManagementSub2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertQueueSpeech]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertQueueSpeech]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertRequestQueueItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertRequestQueueItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertScanItemShelfCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertScanItemShelfCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertSelectItemPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertSelectItemPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertSelectItemPickingMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertSelectItemPickingMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertSelectItemPickingMaster1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertSelectItemPickingMaster1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InsertUpdatePrintLabeLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InsertUpdatePrintLabeLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InvoiceGroupWareHouse]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InvoiceGroupWareHouse]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_InvoiceGroupWareHouseRes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_InvoiceGroupWareHouseRes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ItemChagePriceLevel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ItemChagePriceLevel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_LabelItemPriceLevel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_LabelItemPriceLevel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_LabelItemVendor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_LabelItemVendor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_LabelPriceList_BarCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_LabelPriceList_BarCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_LabelPriceList_ItemCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_LabelPriceList_ItemCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_LabelPriceList_ItemName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_LabelPriceList_ItemName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_NewInsertDataQueueManagement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_NewInsertDataQueueManagement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_NewInsertDataQueueManagement_New]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_NewInsertDataQueueManagement_New]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_NewInsertDataQueueManagement_Test]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_NewInsertDataQueueManagement_Test]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_OpenMinuteItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_OpenMinuteItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_PaymentMoneySubUpdate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_PaymentMoneySubUpdate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_PaymentMoneyUpdate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_PaymentMoneyUpdate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_PickItemNotCreateBill]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_PickItemNotCreateBill]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_PrintLabel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_PrintLabel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_PrintReserveQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_PrintReserveQueue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueueCheckPickingItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueueCheckPickingItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueueHaveRemainPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueueHaveRemainPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueueItemNotStopTime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueueItemNotStopTime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueueManagementByDocDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueueManagementByDocDate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueueManagementByNow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueueManagementByNow]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueuePickingLaterTop5]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueuePickingLaterTop5]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueuePickingNotFully]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueuePickingNotFully]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QueuePickingUnUsual]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QueuePickingUnUsual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QuotaionInsertDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QuotaionInsertDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QuotaionInsertHeader]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QuotaionInsertHeader]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_Quotation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_Quotation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QuotationNewDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QuotationNewDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QuotationSearchItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QuotationSearchItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_QuotationSelectPriceList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_QuotationSelectPriceList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ReportQueueSumPickQTYPerPicker]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ReportQueueSumPickQTYPerPicker]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ReportQueueSumPickQTYPerPickerByDocDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ReportQueueSumPickQTYPerPickerByDocDate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ReserveItemQtyDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ReserveItemQtyDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SaleItemHistory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SaleItemHistory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SeaechUserLogIn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SeaechUserLogIn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchArCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchArCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchArCodeLike]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchArCodeLike]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchBarCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchBarCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchChangePrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchChangePrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchChangePriceDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchChangePriceDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchCheckCountSOPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchCheckCountSOPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchCheckOutHolding]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchCheckOutHolding]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchCheckPickStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchCheckPickStatus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchCheckShelfSaleOrderData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchCheckShelfSaleOrderData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDataQueueDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDataQueueDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDataQueueDetails1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDataQueueDetails1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDataQueueDetails2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDataQueueDetails2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDocQueuePrint]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDocQueuePrint]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDocnoUpdatePrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDocnoUpdatePrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDriveInCheckOut]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDriveInCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchDriveInDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchDriveInDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchGroupPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchGroupPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchHodingBill]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchHodingBill]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchHodingBillDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchHodingBillDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchHoldingDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchHoldingDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchItemChangePrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchItemChangePrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchItemDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchItemDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchItemDetails_Market]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchItemDetails_Market]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchItemPriceDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchItemPriceDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchItemSite]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchItemSite]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchLabelPriceList_UnitCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchLabelPriceList_UnitCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchListDocument]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchListDocument]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchNewDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchNewDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPOS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPOS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPayGoodsPrintReserve]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPayGoodsPrintReserve]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPickingLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPickingLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPickingReq]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPickingReq]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPickingRequest]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPickingRequest]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPrintPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPrintPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPrintQTY]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPrintQTY]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPulsePicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPulsePicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPulsePickingExist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPulsePickingExist]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchPulsePickingLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchPulsePickingLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueCenterBegin]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueCenterBegin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueCenterDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueCenterDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueCenterFinish]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueCenterFinish]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueCenterPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueCenterPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueCheckOut]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQuePayItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQuePayItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueDoc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueDoc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueFinish]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueFinish]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueItemDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueItemDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueItemDetails1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueItemDetails1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueItemDetails2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueItemDetails2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueItemDetails3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueItemDetails3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueItemPickingDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueItemPickingDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueLine]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueLine]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueueLogs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueueLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQueuePrint]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQueuePrint]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQuotation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQuotation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchQuotationDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchQuotationDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchReqPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchReqPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchRequestConfirm]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchRequestConfirm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchRequestQueueItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchRequestQueueItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchSaleOrder]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchSaleOrder]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchSaleOrderData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchSaleOrderData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchSaleOrderGroupShelf]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchSaleOrderGroupShelf]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchSendQueuePicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchSendQueuePicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchStatusMinuteItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchStatusMinuteItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SearchTransferFromDeposit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SearchTransferFromDeposit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SelectDocnoChangePrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SelectDocnoChangePrice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SelectItemChangePricePrintLabel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SelectItemChangePricePrintLabel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SelectItemReceivePrintLabel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SelectItemReceivePrintLabel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_SelectReportName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_SelectReportName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_ShowQueMonitor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_ShowQueMonitor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_StockRequestApprove]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_StockRequestApprove]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateCarLicense]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateCarLicense]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateCheckQtyQue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateCheckQtyQue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateCountOfPrintPicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateCountOfPrintPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateCustItemReceiptStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateCustItemReceiptStatus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateDriveInCancelCheckOut]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateDriveInCancelCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateDriveInCheckerItemCheckOut]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateDriveInCheckerItemCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateDriveInMergeTempConfirm]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateDriveInMergeTempConfirm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateHoldBillQtyQue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateHoldBillQtyQue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateIsCancelQuotation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateIsCancelQuotation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateMydescriptionQueueManagement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateMydescriptionQueueManagement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateMydescriptionQueueManagement1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateMydescriptionQueueManagement1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateMydescriptionQueueManagement2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateMydescriptionQueueManagement2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateMydescriptionQueueManagement3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateMydescriptionQueueManagement3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateNewDocNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateNewDocNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePayGoods]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePayGoods]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePayItemQtyQue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePayItemQtyQue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePickQueCenterSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePickQueCenterSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePrintStatusQueueManagement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePrintStatusQueueManagement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePrintStatusQueueManagement1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePrintStatusQueueManagement1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdatePrintStatusQueueManagement2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdatePrintStatusQueueManagement2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQuePickCenterReason]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQuePickCenterReason]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueReceived]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueReceived]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueStatusDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueStatusDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueCustRec]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueCustRec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueCustRec_Test]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueCustRec_Test]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueMyDescription]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueMyDescription]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueMyDescription1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueMyDescription1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueuePrintStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueuePrintStatus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueReceived]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueReceived]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueReceivedStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueReceivedStatus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueReceivedStatus1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueReceivedStatus1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateQueueReceivedStatus2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateQueueReceivedStatus2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateRequestPickingQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateRequestPickingQueue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateSendQueuePicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateSendQueuePicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_NP_UpdateStatusCustItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_NP_UpdateStatusCustItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_CheckIsCancelInvoice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_CheckIsCancelInvoice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_CheckItemReceiptUpdateCancel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_CheckItemReceiptUpdateCancel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_CheckShelfGroup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_CheckShelfGroup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_InsertCustItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_InsertCustItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_InsertCustItemReceiptSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_InsertCustItemReceiptSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_InsertLineItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_InsertLineItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_SearchCheckCustItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_SearchCheckCustItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_SearchCustItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_SearchCustItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_SearchCustItemReceiptChecking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_SearchCustItemReceiptChecking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_SearchWHCodeCustReceiptItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_SearchWHCodeCustReceiptItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_UpdateCancelCustItemReceipt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_UpdateCancelCustItemReceipt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_QUE_UpdateIsReceivedItem]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_QUE_UpdateIsReceivedItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_NP_GenerateRunNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_NP_GenerateRunNumber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_NP_RunningNumberDocs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_NP_RunningNumberDocs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_NP_RunningNumberDocs_Res]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_NP_RunningNumberDocs_Res]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_CheckInvoicePicking]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_CheckInvoicePicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_InsertPosCoupon]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_InsertPosCoupon]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_InsertPosCouponTemp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_InsertPosCouponTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchDriveInMaster]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchDriveInMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchItemPickUp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchItemPickUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchItemPickUpCancel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchItemPickUpCancel]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchPickUp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchPickUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchReqPickingDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchReqPickingDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchReqPickingInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchReqPickingInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchReqPickingInformationLastSend]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchReqPickingInformationLastSend]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_SearchReqPickingItemZone]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_SearchReqPickingItemZone]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_np_billauto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_np_billauto]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* คิวที่จัดสินค้าเรียบร้อย แต่ลูกค้าไม่ได้มารับสินค้า ณ วันที่จัดของ*/

CREATE	procedure dbo.USP_NP_ARNotReceiveQueueItem
@vZoneID as int,
@vBegDate as nvarchar(20),
@vStopDate as nvarchar(20)
as

set	dateformat dmy

if @vZoneID = 0 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,'') as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code 
where isreceived = 0 and status = 2 and zoneid in ('01') and docdate between @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end
if @vZoneID = 1 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,saleman) as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code 
where isreceived = 0 and status = 2 and zoneid in ('02','03')and docdate between @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_AccessProgram 
@vUserID as nvarchar(20)
as 
select 	top 100 percent a.departmentcode,a.levelid,a.prgid,a.pageid,pagestatus,isnull(mydescription,'') as mydescription
from 	npmaster.dbo.TB_NP_AuthorityProgram a 
where	departmentcode = @vUserID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_AvgItemIssue

as

set dateformat dmy

if exists(select name from npmaster.dbo.sysobjects where name = 'TB_NP_AvgItemIssue')
drop table npmaster.dbo.TB_NP_AvgItemIssue

select	*,	
		case 
		when (diff1/30) = 0 then 0
		else qty/(diff1/30) 
		end as avgissue30,
		qty/diff1 as avgissueday,
		(qty/diff1)*30 as avgissuecalc30
into	npmaster.dbo.TB_NP_AvgItemIssue
from
(
select	*,datediff(day,createdatetime,now) as diff1
from
(
select	itemcode,sum(qty) as qty,unitcode,
		cast(isnull(isnull((select top 1 createdatetime from dbo.bcitem where code = result.itemcode order by createdatetime),(select top 1 a.createdatetime from dbo.bcapinvoice a inner join dbo.bcapinvoicesub b on  a.docno = b.docno and a.docdate = b.docdate where itemcode = result.itemcode order by a.createdatetime)),(select top 1 docdate from dbo.BCStkIssueSub where itemcode = result.itemcode order by createdatetime)) as datetime) as createdatetime,
		cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) as now
from
(
select	'เบิก' as typq,itemcode,sum(qty) as qty,unitcode 
from	dbo.BCSTKISSUE a
		left join dbo.BCStkIssueSub b on a.docno = b.docno and a.docdate = b.docdate
where	a.iscancel =0 
group	by itemcode,unitcode
union
select	'รับคืน' as type,itemcode,-1*sum(qty) as qty,unitcode 
from	dbo.BCSTKISSUERET a 
		left join dbo.BCStkIssRetSub b on a.docno = b.docno and a.docdate = b.docdate 
where	left(a.docno,2) in ('id') and a.iscancel = 0
group	by itemcode,unitcode
) as	result
group	by itemcode,unitcode
) as	result1
) as	result2
order	by itemcode


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_BPlusDepartment
as
select   department,dept_thaidesc from npmaster.dbo.TB_NP_Department 
order	by department

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CalcDriveInMergeTemp
@vDocNo nvarchar(20)
as

set		dateformat dmy

select	result.*,isnull(a.name1,'') as itemname,result.qty*(price-discountamount) as amount,(select top 1 sum(qty*(price-discountamount)) as net from npmaster.dbo.TB_NP_DriveInMergeTemp where docno = @vDocNo) as netamount,
	isnull(b.rate1,1) as rate1,isnull(rate2,1) as rate2
from 
(
select	DocNo,DocDate,ItemCode,sum(QTY) as qty,UnitCode,Price,BarCode,whcode,shelfcode,(select top 1 discountamount from npmaster.dbo.TB_NP_DriveInMergeTemp where docno = @vDocNo and a.itemcode = itemcode and a.unitcode = unitcode order by docno desc) as discountamount,
	(select top 1 isnull(refno,'') as refno from npmaster.dbo.TB_NP_DriveInMergeTemp where docno = @vDocNo and a.itemcode = itemcode and a.unitcode = unitcode order by refno desc) as refno
from	npmaster.dbo.TB_NP_DriveInMergeTemp a
where	docno = @vDocNo
group	by DocNo,DocDate,ItemCode,UnitCode,Price,BarCode,whcode,shelfcode
) as	result
	left join dbo.bcitem a on result.itemcode = a.code
	left join dbo.bcstkpacking b on result.itemcode = b.itemcode and result.unitcode = b.unitcode
order	by result.itemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CancelDriveInDocNo
@vDocNo as nvarchar(50)

as

set		dateformat dmy

declare	@vExist as int

set		@vExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_DriveInSlipMaster where docno = @vDocNo and iscancel = 0)

if		@vExist <> 0 
begin
update	npmaster.dbo.TB_NP_DriveInSlipMaster  
set		iscancel = 1
where	docno = @vDocNo and iscancel = 0
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CancelReqPicking
@vDocNo as nvarchar(50)

as

set		dateformat dmy

declare	@vExist as int

set		@vExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_PickingRequestMaster where docno = @vDocNo and iscancel = 0)

if		@vExist <> 0 
begin
update	npmaster.dbo.TB_NP_PickingRequestMaster  
set		iscancel = 1
where	docno = @vDocNo and iscancel = 0
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_CheckAuthorizePromotion
@vDepartmentCode as nvarchar(20),
@vLevelID as int
as
select 	a.prgid,a.pageid,a.pagename,a.pagedescription,
	isnull(departmentcode,'') as departmentcode,isnull(levelid,0) as levelid,isnull(pagestatus,0) as pagestatus 
from 	npmaster.dbo.TB_NP_PageProgram a
	left join npmaster.dbo.TB_NP_AuthorityProgram b on a.prgid = b.prgid and a.pageid = b.pageid and departmentcode = @vDepartmentCode and levelid = @vLevelID
where 	a.prgid = '02' 
order	by a.pageid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckInsertItemCodeChangePrice
@vScheduleDate as nvarchar(20),
@vItemCode as nvarchar(25),
@vPriceLevel as smallint,
@vsaleType as smallint,
@vTransSportType as smallint,
@vUnitCode as nvarchar(30)

as

set	dateformat dmy
select 	isnull(count(itemcode),0) as vCountItem
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceSub a
	left join npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster b on a.docno = b.docno and a.docdate = b.docdate
where 	scheduledate = @vScheduleDate and 
	itemcode = @vItemCode and 
	pricelevel = @vPriceLevel and 
	saletype = @vsaleType and 
	transsporttype = @vTransSportType and 
	unitcode = @vUnitCode and 
	isupdate = 0
group	by itemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_CheckItemCode
@vItemCode as nvarchar(20)
as

set	dateformat dmy
set	language us_english
select	code,name1,defsaleunitcode from dbo.BCItem where activestatus = 1 and Code = @vItemCode 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckItemInRecProduct2
@vItemCode as nvarchar(20)
as
select 	isnull(count(productcode),0) as vCount 
from
(
select	productcode
from 	dbo.bcrecproduct2 
where	productcode = @vItemCode
union
select	itemcode as productcode
from	npmaster.dbo.np_scanbarcode_logs
where	itemcode = @vItemCode
) as 	result
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckLevelPrice
@vItemCode as nvarchar(30),
@vUnitCode as nvarchar(20)
as

set	dateformat dmy
select 	itemcode,count(itemcode) as vCountLine,unitcode,saletype,transporttype 
from 	dbo.bcpricelist 
where 	itemcode = @vItemCode and remark <> 'promotion' and unitcode = @vUnitCode
group 	by itemcode,unitcode,saletype,transporttype
order	by itemcode,saletype,transporttype

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckPayBillByDateTime
as
set	dateformat dmy

select 	a.docno,a.docdate,month(a.docdate) as DocMonth,year(a.docdate) as DocYear,a.arcode,a.sumofinvoice,sumofdebitnote,sumofcreditnote,totalamount,creditday,paybillamount,
	b.invoiceno,b.invoicedate,invbalance,payamount,paybalance 
from 	dbo.BCPAYBILL a
	left join dbo.BCPAYBILLsub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
where 	a.iscancel = 0 and a.docdate between cast(rtrim(day((getdate()-60)))+'/'+rtrim(month((getdate()-60)))+'/'+rtrim(year((getdate()-60))) as datetime)and cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1)))as datetime)
order 	by a.docdate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckPayGoods
as
set	dateformat dmy
select  a.invoiceno,a.whcode,paynumber,b.docno,paydatetime,isnull(c.name1,'') as arname
from 	npmaster.dbo.np_paygoods a 
	left join dbo.bcarinvoicesub b on a.invoiceno = b.docno and a.whcode = b.whcode
	left join dbo.bcar c on b.arcode = c.code
where 	checked = 0 and b.docno is not null and year(paydatetime) = year(getdate()) and 
	month(paydatetime) = month(getdate()) and day(paydatetime) = day(getdate()) 
group 	by a.invoiceno,a.whcode,paynumber,b.docno,paydatetime,c.name1
order	by a.paydatetime desc



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckQuePickCenter
@vDocno as nvarchar(20),
@vQueDocDate as nvarchar(20)
as

set		dateformat dmy

select	isnull(max(QueTime),0) as Max1 
from	npmaster.dbo.TB_NP_QuePickCenterMaster 
where	docno = @vDocno and quedocdate = @vQueDocDate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckQuePickCenterZone
@vQueID as int,
@vQueDocDate as nvarchar(20)
as

set	dateformat dmy

select	isnull(quezone,'') as quezone 
from	npmaster.dbo.TB_NP_QuePickCenterMaster 
where	queid = @vQueID and quedocdate = @vQueDocDate  
order	by quedate desc 






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckQueueMonitorMoreThan
@vCondition as int,
@vZoneID as int,
@vTime as nvarchar(20),
@vBegDate as nvarchar(20),
@vEndDate as nvarchar(20)

as

set	dateformat dmy

if 	@vCondition = 0 
begin
if 	@vZoneID = 0
begin
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty,e.itemcode,e.itemname,e.qty,e.pickqty,e.unitcode
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	inner join BCHistory.dbo.TB_NP_QueueManagementSubLogs e on a.docno = e.pickingno and a.docdate = e.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between @vBegDate and @vEndDate
and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > @vTime
order	by a.docdate,cast(a.docno as int)
end
if 	@vZoneID = 1
begin
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty,e.itemcode,e.itemname,e.qty,e.pickqty,e.unitcode
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	inner join BCHistory.dbo.TB_NP_QueueManagementSubLogs e on a.docno = e.pickingno and a.docdate = e.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('02','03') and status = 2 and a.docdate between @vBegDate and @vEndDate
and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > @vTime
order	by a.docdate,cast(a.docno as int)
end
end

if @vCondition = 1 
begin
if 	@vZoneID = 0
begin
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty,e.itemcode,e.itemname,e.qty,e.pickqty,e.unitcode
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	inner join BCHistory.dbo.TB_NP_QueueManagementSubLogs e on a.docno = e.pickingno and a.docdate = e.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between @vBegDate and @vEndDate
and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) < @vTime
order	by a.docdate,cast(a.docno as int)
end
if 	@vZoneID = 1
begin
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty,e.itemcode,e.itemname,e.qty,e.pickqty,e.unitcode
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	inner join BCHistory.dbo.TB_NP_QueueManagementSubLogs e on a.docno = e.pickingno and a.docdate = e.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('02','03') and status = 2 and a.docdate between @vBegDate and @vEndDate
and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) < @vTime
order	by a.docdate,cast(a.docno as int)
end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CheckShelfPrintPicking 
@vDocno as nvarchar(20),
@vTimeID as int 
as 

set	dateformat dmy
select 	docno,docdate,saleorderno,whcode,shelfgroup,isreceived,status,picker,saleman,timeid,refdocno
from 	npmaster.dbo.tb_np_queuemanagement
where 	docno = @vDocno and timeid = @vTimeID and isreceived = 0 and docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate()))as datetime)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_CheckStockItemCode
@vItemCode as nvarchar(20)
as

set	dateformat dmy
set	language us_english
select	ItemCode,sum(qty) as QTYOnHand,UnitCode from dbo.BCStkLocation where ItemCode = @vItemCode group by ItemCode,UnitCode

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CloseOpenItemMinuteLogs
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vUnitCode as nvarchar(20),
@vUserOpen as nvarchar(50)
as
Declare @vCloseDateTime as nvarchar(20)
set	@vCloseDateTime = getdate()

set	dateformat dmy
set	language us_english
update npmaster.dbo.TB_HMX_OpenItemMinuteLogs set closedatetime = @vCloseDateTime
where	docdate = @vDocDate and itemcode = @vItemCode and unitcode = @vUnitCode and useropen = @vUserOpen

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_CouponData
@vCPCode as nvarchar(20)
as

set		dateformat dmy
set		language us_english

select	a.cpcode,a.cpname,a.fromdate,a.todate,a.cpformat,a.cplenght,isnull(a.isused,0) as isused,isnull(a.mydescription,'') as mydescription,b.cpheader,b.cpvalue,b.cpqty,b.cpapprove,b.cpremain,b.linenumber
from	npmaster.dbo.tb_np_couponmaster a 
		inner join npmaster.dbo.tb_np_coupondetails b on a.cpcode = b.cpcode and a.cpformat = b.cpformat
where	a.cpcode = @vCPCode 
order	by a.cpcode ,b.linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_DeleteChangePriceDocNo
@vDocNo as nvarchar(30)
as

set	dateformat dmy

delete 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster where docno = @vDocNo
delete 	npmaster.dbo.TB_NP_BasketItemUpdatePriceSub  where docno = @vDocNo


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create 	procedure dbo.USP_NP_DeleteDataPrintLabel
@vUserID as nvarchar(20)
as
set dateformat dmy
Delete From dbo.NP_Label_Temp Where Useduser =@vUserID
Delete From dbo.TB_NP_PrintLabelTemp Where Useduser =@vUserID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_DeleteDriveInMergeTemp
@vDocNo nvarchar(20)
as

set		dateformat dmy

delete	npmaster.dbo.TB_NP_DriveInMergeTemp
where	docno = @vDocNo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_DeleteHoldingBill
@vType as int,
@vDocNo as nvarchar(30)

as

set		dateformat dmy
declare	@vCheckExist as int

set		@vCheckExist = (select isnull(count(docno),0) as vcount from bcnpdisa.dbo.bpsholdingbill where docno = @vDocno)


if	@vType = 1
begin
	if		@vCheckExist > 0
	begin
	update	npmaster.dbo.TB_NP_QuePickCenterMaster
	set		isconfirm = 0,mergeno = null,holdbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_PickingRequestMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_PickingRequestSub d on c.docno = d.docno and b.itemcode = d.itemcode 	
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
			c.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	
	update	npmaster.dbo.TB_NP_QuePickCenterSub
	set		mergeno = null,holdbillno = null,checkqty = 0 ,invqty = 0
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_PickingRequestMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_PickingRequestSub d on c.docno = d.docno and b.itemcode = d.itemcode 	
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
			c.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	update	npmaster.dbo.TB_NP_DriveInMergeTemp
	set		isconfirm = 0,posbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_PickingRequestMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_PickingRequestSub d on c.docno = d.docno and b.itemcode = d.itemcode 	
			inner join npmaster.dbo.TB_NP_DriveInMergeTemp e on b.itemcode = e.itemcode and c.docno = e.refno and a.docno = e.posbillno
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
			c.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	delete	bcnpdisa.dbo.bpsholdingbill where docno = @vDocNo
	delete	bcnpdisa.dbo.bpsholdingbillsub where docno = @vDocNo
	end
end

if	@vType = 2
begin
	if		@vCheckExist > 0
	begin
	update	npmaster.dbo.TB_NP_QuePickCenterMaster
	set		isconfirm = 0,mergeno = null,holdbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_QueueRequestPickingMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_QueueRequestPicking d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	update	npmaster.dbo.TB_NP_QuePickCenterSub
	set		mergeno = null,holdbillno = null,checkqty = 0 ,invqty = 0
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_QueueRequestPickingMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_QueueRequestPicking d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	update	npmaster.dbo.TB_NP_DriveInMergeTemp
	set		isconfirm = 0,posbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_QueueRequestPickingMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_QueueRequestPicking d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_DriveInMergeTemp e on b.itemcode = e.itemcode and c.docno = e.refno and a.docno = e.posbillno
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	delete	bcnpdisa.dbo.bpsholdingbill where docno = @vDocNo
	delete	bcnpdisa.dbo.bpsholdingbillsub where docno = @vDocNo
	end
end

if	@vType = 3
begin
	if		@vCheckExist > 0
	begin
	update	npmaster.dbo.TB_NP_QuePickCenterMaster
	set		isconfirm = 0,mergeno = null,holdbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_DriveInSlipMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_DriveInSlipSub d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	update	npmaster.dbo.TB_NP_QuePickCenterSub
	set		mergeno = null,holdbillno = null,checkqty = 0 ,invqty = 0
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_DriveInSlipMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_DriveInSlipSub d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterSub e on a.docno = e.holdbillno and c.docno = e.refno and d.itemcode = e.itemcode 
			inner join npmaster.dbo.TB_NP_QuePickCenterMaster f on e.queid = f.queid and e.quedocdate = f.quedocdate
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	update	npmaster.dbo.TB_NP_DriveInMergeTemp
	set		isconfirm = 0,posbillno = null
	from	bcnpdisa.dbo.bpsholdingbill a 
			inner join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
			inner join npmaster.dbo.TB_NP_DriveInSlipMaster c on b.sorefno = c.docno 
			inner join npmaster.dbo.TB_NP_DriveInSlipSub d on c.docno = d.docno and b.itemcode = d.itemcode 
			inner join npmaster.dbo.TB_NP_DriveInMergeTemp e on b.itemcode = e.itemcode and c.docno = e.refno and a.docno = e.posbillno
	where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	
	
	delete	bcnpdisa.dbo.bpsholdingbill where docno = @vDocNo
	delete	bcnpdisa.dbo.bpsholdingbillsub where docno = @vDocNo
	end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_DeleteQueueItemSub
@vDocNo as int,
@vDocDate as nvarchar(20),
@vTimeID as int
as

set		dateformat dmy

delete	npmaster.dbo.tb_np_queuemanagementsub
where	pickingno = @vDocNo and docdate = @vDocDate and timeid = @vTimeID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_DeleteRequestConfirm
@vDocno as nvarchar(20)
as

set	dateformat dmy

declare	@vExist as int

set	@vExist = (select count(*) as vCount from dbo.bcreqconfirm where docno = @vDocno and billstatus = 0)
if 	@vExist <> 0
begin

update	dbo.bcstkrequest 
set	billstatus = 0
from	dbo.bcstkrequest a
	inner join (select distinct stkrequestno as docno from dbo.bcreqconfirmsub where docno = @vDocno) b on a.docno = b.docno 

update	dbo.bcstkrequest 
set	isconfirm = 0
from	dbo.bcstkrequest a
	inner join (select distinct stkrequestno as docno from dbo.bcreqconfirmsub where docno = @vDocno) b on a.docno = b.docno 


update	dbo.bcstkrequestsub
set	remainqty = qty
from	dbo.bcstkrequestsub a
	inner join (select distinct stkrequestno as docno from dbo.bcreqconfirmsub where docno = @vDocno) b on a.docno = b.docno 

delete 	dbo.bcreqconfirmsub where docno = @vDocno
 
delete	dbo.bcreqconfirm where docno = @vDocno

end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_DriveInCheckOutPos
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vPosNo as nvarchar(20),
@vUserID as nvarchar(30),
@vItemCode as nvarchar(20),
@vInvQTY as money,
@vAmount as money
as

set		dateformat dmy

update	npmaster.dbo.tb_np_driveinslipmaster
set	billposno = @vPosNo,isconfirm = 1,confirmcode = @vUserID,confirmdatetime = getdate()
where	docno = @vDocNo and docdate = @vDocDate


update	npmaster.dbo.tb_np_driveinslipsub
set	invqty = @vInvQTY,amount = @vAmount
where	docno = @vDocNo and docdate = @vDocDate and itemcode = @vItemCode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_GenerateItemChangePriceNumber
as
set	dateformat dmy
select 	year(getdate())+543 as Year1,month(getdate()) as Month1,isnull(cast(right(max(docno),3)as int)+1,1) as MaxNumber
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster
where	year(docdate) = year(getdate())and month(docdate) = month(getdate())

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_GetMaxNoHoldingBill
@vMachineNo as nvarchar(10),
@vDocDate as nvarchar(20)
as

set		dateformat dmy

select		isnull(left(maxno,8),
		@vMachineNo+right(rtrim(year(getdate())+543),2)+case len(month(getdate())) when 1 then rtrim('0'+rtrim(month(getdate()))) when 2 then rtrim(month(getdate())) end +case len(day(getdate())) when 1 then rtrim('0'+rtrim(day(getdate()))) when 2 then rtrim(day(getdate())) end) as header,cast(isnull(right(maxno,4),0)as int)+1 as maxnumber
		
from
		(
		select	max(docno) as maxno from bcnpdisa.dbo.BPSHoldingBill where machineno = @vMachineNo and docdate = @vDocDate
		) as	a
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertAuthorizePromotion
@vDepartmentCode as nvarchar(20),
@vLevelID as int,
@vPrgID as nvarchar(10),
@vPageID as nvarchar(10),
@vPageStatus as int
as

set	dateformat dmy
declare @vExist as int 

set	 @vExist = (select isnull(count(pageid),0) as vCount from npmaster.dbo.TB_NP_AuthorityProgram where departmentcode = @vDepartmentCode and levelid = @vLevelID and pageid = PageID and PrgID = @vPrgID)

if 	@vExist = 0 
begin
	insert into npmaster.dbo.TB_NP_AuthorityProgram(DepartmentCode,LevelID,PrgID,PageID,PageStatus)
	values (@vDepartmentCode,@vLevelID,@vPrgID,@vPageID,@vPageStatus)
end

if 	@vExist > 0
begin
	update npmaster.dbo.TB_NP_AuthorityProgram 
	set PageStatus = @vPageStatus 
	where DepartmentCode = @vDepartmentCode and 
	LevelID = @vLevelID and 
	PrgID = @vPrgID and 
	PageID = @vPageID
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertBasketUpdateItemPrice
@vType as smallint,
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vScheduleDate as nvarchar(20),
@vCreatorCode as nvarchar(50)

as

declare @vIsConfirm as smallint

set	dateformat dmy
set	@vIsConfirm = 0

if 	@vType = 0
begin
insert	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster(DocNo,DocDate,ScheduleDate,CreatorCode,CreateDate,IsConfirm)
values	(@vDocNo,@vDocDate,@vScheduleDate,@vCreatorCode,getdate(),@vIsConfirm)
end

if 	@vType = 1
begin
update	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster
set	ScheduleDate = @vScheduleDate,EditorCode=@vCreatorCode,EditDate = getdate()
where	docno = @vDocNo

delete	npmaster.dbo.TB_NP_BasketItemUpdatePriceSub 
where	docno = @vDocNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertBasketUpdateItemPriceDetails
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(25),
@vItemName as nvarchar(250),
@vUnitCode as nvarchar(25),
@vPriceLevel as smallint,
@vSaleType as smallint,
@vTransSportType as smallint,
@vNewPrice as money,
@vOldPrice as money,
@vLineNumber as int

as

declare	@vIsUpDate as smallint

set	dateformat dmy
set	@vIsUpDate = 0


insert	npmaster.dbo.TB_NP_BasketItemUpdatePriceSub(DocNo,DocDate,ItemCode,ItemName,UnitCode,PriceLevel,SaleType,TransSportType,NewPrice,OldPrice,LineNumber)
values	(@vDocNo,@vDocDate,@vItemCode,@vItemName,@vUnitCode,@vPriceLevel,@vSaleType,@vTransSportType,@vNewPrice,@vOldPrice,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertCouponDetails
@vCPCode as nvarchar(20),
@vCPHeader as nvarchar(10),
@vCPFormat as int,
@vCPValue as money,
@vCPQty as int,
@vCPApprove as int,
@vCPRemain as int,
@vLineNumber as int

as

set		dateformat dmy
set		language us_english

declare	@vCheckIsUsed as smallint

set		@vCheckIsUsed = (select isused from npmaster.dbo.TB_NP_CouponMaster where cpcode = @vCPCode)

if		@vCheckIsUsed = 0
begin
insert	into npmaster.dbo.TB_NP_CouponDetails(CPCode,CPHeader,CPFormat,CPValue,CPQty,CPApprove,CPRemain,LineNumber)
		values(@vCPCode,@vCPHeader,@vCPFormat,@vCPValue,@vCPQty,@vCPApprove,@vCPQty,@vLineNumber)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertCouponMaster
@vCPCode as nvarchar(20),
@vCPName as nvarchar(200),
@vFromDate as nvarchar(20),
@vToDate as nvarchar(20),
@vCPMerge as int,
@vCPFormat as int,
@vCPLenght as int,
@vMyDescription as nvarchar(200),
@vUserID as nvarchar(50)

as

set		dateformat dmy
set		language us_english

declare	@vCheckCPCode as int
declare	@vCheckIsUsed as smallint
declare	@vIsUsed as smallint

set		@vIsUsed = 0

set		@vCheckCPCode = (select count(cpcode) as vCount from npmaster.dbo.TB_NP_CouponMaster where cpcode = @vCPCode)
set		@vCheckIsUsed = (select isnull(isused,0) as isused from npmaster.dbo.TB_NP_CouponMaster where cpcode = @vCPCode)
if		@vCheckCPCode = 0
begin
insert	into npmaster.dbo.TB_NP_CouponMaster(CPCode,CPName,FromDate,ToDate,CPMerge,CPFormat,CPLenght,IsUsed,MyDescription,CreatorCode,CreateDateTime)
	values(@vCPCode,@vCPName,@vFromDate,@vToDate,@vCPMerge,@vCPFormat,@vCPLenght,@vIsUsed,@vMyDescription,@vUserID,getdate())
delete	npmaster.dbo.TB_NP_CouponDetails where cpcode = @vCPCode
end
else
begin
if		@vCheckIsUsed = 0
begin
		update	npmaster.dbo.TB_NP_CouponMaster
		set	CPName = @vCPName,FromDate = @vFromDate,ToDate = @vToDate,CPMerge=@vCPMerge,CPFormat = @vCPFormat,CPLenght = @vCPLenght,MyDescription = @vMyDescription,LastEditorCode = @vUserID,LastEditDateTime = getdate()
		where	CPCode = @vCPCode

delete	npmaster.dbo.TB_NP_CouponDetails where cpcode = @vCPCode
end
else
begin
		update	npmaster.dbo.TB_NP_CouponMaster
		set	ToDate = @vToDate,MyDescription = @vMyDescription,LastEditorCode = @vUserID,LastEditDateTime = getdate()
		where	CPCode = @vCPCode
end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDataPrintServer
@vJobID as nvarchar(2),
@vDocNo as nvarchar(20),
@vReportID as int,
@vReportType as nvarchar(10), 
@vPara1 as nvarchar(20),
@vPara2 as nvarchar(20),
@vPara3 as nvarchar(20),
@vPara4 as nvarchar(20),
@vPara5 as nvarchar(20),
@vPara6 as nvarchar(20),
@vPara7 as nvarchar(20),
@vPara8 as nvarchar(20),
@vPrintStatus as int,
@vUserPrint as nvarchar(20)

as

declare @vDatePrint as datetime
set 	dateformat dmy
set	@vDatePrint = getdate()

insert	into npmaster.dbo.TB_NP_CheckQueuePrint(JobID,DocNo,ReportID,ReportType,Para1,Para2,Para3,Para4,Para5,Para6,Para7,Para8,PrintStatus,UserPrint,DatePrint)
values 	(@vJobID,@vDocNo,@vReportID,@vReportType,@vPara1,@vPara2,@vPara3,@vPara4,@vPara5,@vPara6,@vPara7,@vPara8,@vPrintStatus,@vUserPrint,@vDatePrint)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE	procedure dbo.USP_NP_InsertDataQueueManagement
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vDocType as int,
@vARCode as nvarchar(20),
@vSaleMan as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vRefDocNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(10),
@vZoneID as nvarchar(2),
@vTimeID as int,
@vStatus as int

as

declare	@vStatusDesc as nvarchar(30)
declare	@vPickingStatus as int
declare	@vIsReceived as int
declare @vQueueDateTime as nvarchar(20)

set 	dateformat dmy
set	LANGUAGE us_english


SET LOCK_TIMEOUT 20000--  lock object  ได้ 10 วินาที

set @vStatusDesc = 'รอจัด'
set @vPickingStatus = 0
set @vIsReceived = 0
set @vQueueDateTime = getdate()	

insert	into npmaster.dbo.TB_NP_QueueManagement
	(Docno,Docdate,DocType,ARCode,Status,StatusDesc,QueueDateTime,PickingStatus,IsReceived,SaleMan,SaleOrderNo,RefDocNo,WHCode,ShelfGroup,ZoneID,TimeID,StartDateTime,Picker)
values 	(@vDocno,@vDocdate,@vDocType,@vARCode,@vStatus,@vStatusDesc,@vQueueDateTime,@vPickingStatus,@vIsReceived,@vSaleMan,@vSaleOrderNo,@vRefDocNo,@vWHCode,@vShelfGroup,@vZoneID,@vTimeID,getdate(),'OutLet')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE	procedure dbo.USP_NP_InsertDataQueueManagement1
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vDocType as int,
@vARCode as nvarchar(20),
@vSaleMan as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vRefDocNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(2),
@vZoneID as nvarchar(10),
@vTimeID as int,
@vStatus as int

as

declare	@vStatusDesc as nvarchar(30)
declare	@vPickingStatus as int
declare	@vIsReceived as int
declare @vQueueDateTime as nvarchar(20)

set 	dateformat dmy
set	LANGUAGE us_english


SET LOCK_TIMEOUT 20000--  lock object  ได้ 10 วินาที

set @vStatusDesc = 'รอจัด'
set @vPickingStatus = 0
set @vIsReceived = 0
set @vQueueDateTime = getdate()	

insert	into npmaster.dbo.TB_NP_QueueManagement_Test
	(Docno,Docdate,DocType,ARCode,Status,StatusDesc,QueueDateTime,PickingStatus,IsReceived,SaleMan,SaleOrderNo,RefDocNo,WHCode,ShelfGroup,ZoneID,TimeID,StartDateTime,Picker)
values 	(@vDocno,@vDocdate,@vDocType,@vARCode,@vStatus,@vStatusDesc,@vQueueDateTime,@vPickingStatus,@vIsReceived,@vSaleMan,@vSaleOrderNo,@vRefDocNo,@vWHCode,@vShelfGroup,@vZoneID,@vTimeID,getdate(),'OutLet')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDriveInCheckOut
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vChecker as nvarchar(30),
@vNetDebtAmount as money,
@vUserID as nvarchar(30)

as

set		dateformat dmy

declare	@vExist as int
set		@vExist = (select isnull(count(docno),0) as vcount from npmaster.dbo.TB_NP_DriveInCheckOut where docno = @vDocNo)

if		@vExist = 0
begin
insert	into npmaster.dbo.TB_NP_DriveInCheckOut(DocNo,DocDate,Checker,PosNo,NetDebtAmount,IsCancel,IsConfirm,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vChecker,'',@vNetDebtAmount,0,0,@vUserID,getdate())
end

if		@vExist > 0
begin
update	npmaster.dbo.TB_NP_DriveInCheckOut
set	Checker=@vChecker,NetDebtAmount=@vNetDebtAmount
where	docno = @vDocNo

delete	npmaster.dbo.TB_NP_DriveInCheckOutSub
where	docno = @vDocNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDriveInCheckOutSub
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(30),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(20),
@vQTY as money,
@vBillQTY as money,
@vUnitCode as nvarchar(20),
@vPrice as money,
@vAmount as money,
@vBarCode as nvarchar(30),
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_DriveInCheckOutSub(DocNo,DocDate,ItemCode,WHCode,ShelfCode,QTY,BillQTY,UnitCode,Price,Amount,BarCode,IsCancel,LineNumber)
values	(@vDocNo,@vDocDate,@vItemCode,@vWHCode,@vShelfCode,@vQTY,@vBillQTY,@vUnitCode,@vPrice,@vAmount,@vBarCode,0,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDriveInMergeTemp
@vDocNo nvarchar(20),
@vDocDate nvarchar(20),
@vItemCode nvarchar(30),
@vWHCode nvarchar(10),
@vShelfCode nvarchar(20),
@vQTY as money,
@vUnitCode nvarchar(20),
@vPrice as money,
@vDisCountAmount as money,
@vAmount as money,
@vBarCode as nvarchar(20),
@vRefNo as nvarchar(20),
@vQueID as int,
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_DriveInMergeTemp(DocNo,DocDate,ItemCode,WHCode,ShelfCode,QTY,UnitCode,Price,DisCountAmount,Amount,BarCode,RefNo,QueID,LineNumber)
values	(@vDocNo,@vDocDate,@vItemCode,@vWHCode,@vShelfCode,@vQTY,@vUnitCode,@vPrice,@vDisCountAmount,@vAmount,@vBarCode,@vRefNo,@vQueID,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDriveInSlip
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vArCode as nvarchar(20),
@vMemberID as nvarchar(30),
@vSaleCode as nvarchar(30),
@vRefNo as nvarchar(50),
@vPickZone as nvarchar(2),
@vBeforeTaxAmount as money,
@vTaxAmount as money,
@vTotalNetAmount as money,
@vCreatorCode as nvarchar(30)

as

declare	@vCheckExist as int

set		dateformat dmy

set	@vCheckExist = (select count(docno) from npmaster.dbo.TB_NP_DriveInSlipMaster where docno = @vDocno)

if	@vCheckExist = 0 
begin
insert	npmaster.dbo.TB_NP_DriveInSlipMaster(DocNo,DocDate,ARCode,MemberID,SaleCode,RefNo,PickZone,BeforeTaxAmount,TaxAmount,TotalNetAmount,IsCancel,IsMerge,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vARCode,@vMemberID,@vSaleCode,@vRefNo,@vPickZone,@vBeforeTaxAmount,@vTaxAmount,@vTotalNetAmount,0,0,@vCreatorCode,getdate())
end

if	@vCheckExist > 0 
begin
update	npmaster.dbo.TB_NP_DriveInSlipMaster
set	arcode = @vARCode,memberid=@vMemberID,salecode = @vSaleCode,refNo = @vRefNo,BeforeTaxAmount=@vBeforeTaxAmount,TaxAmount=@vTaxAmount,TotalNetAmount = @vTotalNetAmount,LastEditorCode = @vCreatorCode,LastEditDateTime = getdate()
where	docno = @vDocno
end

delete	npmaster.dbo.TB_NP_DriveInSlipSub where docno = @vDocno
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertDriveInSlipSub
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(250),
@vWHCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vShelfID as nvarchar(20),
@vZoneID as nvarchar(20),
@vQTY as money,
@vUnitCode as nvarchar(20),
@vPrice as money,
@vDisCountWord as nvarchar(50),
@vDisCountAmount as money,
@vAmount as money,
@vBarCode as nvarchar(30),
@vLineNumber as int
as

set		dateformat dmy

insert	npmaster.dbo.TB_NP_DriveInSlipSub(DocNo,DocDate,ItemCode,ItemName,WHCode,ShelfCode,ShelfID,ZoneID,QTY,UnitCode,Price,DisCountWord,DisCountAmount,Amount,IsCancel,BarCode,LineNumber)
values	(@vDocNo,@vDocDate,@vItemCode,@vItemName,@vWHCode,@vShelfCode,@vShelfID,@vZoneID,@vQTY,@vUnitCode,@vPrice,@vDisCountWord,@vDisCountAmount,@vAmount,0,@vBarCode,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertHoldingBillDriveIn
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vExpireCredit as smallint,
@vArCode as nvarchar(20),
@vCashierCode as nvarchar(20),
@vMachineNo as nvarchar(20),
@vMachineCode as nvarchar(20),
@vSaleCode as nvarchar(20),
@vTaxRate as money,
@vSumOfItemAmount as money,
@vAfterDiscount as money,
@vBeforeTaxAmount as money,
@vTaxAmount as money,
@vTotalAmount as money,
@vNetDebtAmount as money,
@vCreatorCode as nvarchar(20),
@vSHIFTCODE as nvarchar(20),
@vMyDescription as nvarchar(100)
as

set		dateformat dmy

declare	@vCheckExist as int
declare	@vBillTime as nvarchar(20)

set		@vCheckExist = (select count(docno) from bcnpdisa.dbo.BPSHoldingBill where docno = @vDocNo)
set		@vBillTime = (select rtrim(datepart(hh,getdate()))+':'+rtrim(datepart(n,getdate()))+':'+rtrim(datepart(ss,getdate())) as time1)

if		@vCheckExist = 0 
begin
		insert	into bcnpdisa.dbo.BPSHoldingBill(DocNo,DocDate,ExpireCredit,ArCode,CashierCode,MachineNo,MachineCode,BillTime,SaleCode,TaxRate,SumOfItemAmount,AfterDiscount,BeforeTaxAmount,TaxAmount,TotalAmount,NetDebtAmount,CreatorCode,CreateDateTime,SHIFTCODE,MyDescription)
		values	(@vDocNo,@vDocDate,@vExpireCredit,@vArCode,@vCashierCode,@vMachineNo,@vMachineCode,@vBillTime,@vSaleCode,@vTaxRate,@vSumOfItemAmount,@vAfterDiscount,@vBeforeTaxAmount,@vTaxAmount,@vTotalAmount,@vNetDebtAmount,@vCreatorCode,getdate(),@vSHIFTCODE,@vMyDescription)
end


if		@vCheckExist <> 0 
begin
		update	bcnpdisa.dbo.BPSHoldingBill
		set	CashierCode=@vCashierCode,MachineNo=@vMachineNo,MachineCode=@vMachineCode,BillTime=@vBillTime,SaleCode=@vSaleCode,
			TaxRate=@vTaxRate,SumOfItemAmount=@vSumOfItemAmount,AfterDiscount=@vAfterDiscount,BeforeTaxAmount=@vBeforeTaxAmount,
			TaxAmount=@vTaxAmount,TotalAmount=@vTotalAmount,NetDebtAmount=@vNetDebtAmount,LastEditorCode=@vCreatorCode,LastEditDateT =getdate()
		where	docno = @vDocNo and docdate = @vDocDate

		delete	bcnpdisa.dbo.BPSHoldingBillsub where docno = @vDocNo and docdate = @vDocDate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertHoldingBillDriveInSub
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vWHCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vQTY as money,
@vPrice as money,
@vDiscountAmount as money,
@vAmount as money,
@vNetAmount as money,
@vUnitCode as nvarchar(20),
@vStockType as smallint,
@vLineNumber as int,
@vBarCode as nvarchar(20),
@vCashierCode as nvarchar(20),
@vPosStatus as smallint,
@vSORefNo as nvarchar(20)

as

set		dateformat dmy

declare		@vBillTime as nvarchar(20)

set		@vBillTime = (select rtrim(datepart(hh,getdate()))+':'+rtrim(datepart(n,getdate()))+':'+rtrim(datepart(ss,getdate())) as time1)

insert	into bcnpdisa.dbo.BPSHoldingBillSUB(DocNo,ItemCode,DocDate,WHCode,ShelfCode,Qty,Price,DiscountAmount,Amount,NetAmount,UnitCode,StockType,LineNumber,BarCode,BillTime,CashierCode,PosStatus,SORefNo)
values	(@vDocNo,@vItemCode,@vDocDate,@vWHCode,@vShelfCode,@vQty,@vPrice,@vDiscountAmount,@vAmount,@vNetAmount,@vUnitCode,@vStockType,@vLineNumber,@vBarCode,@vBillTime,@vCashierCode,@vPosStatus,@vSORefNo)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertItemReqPicking
@docno as nvarchar(20),
@docdate as nvarchar(20),
@itemcode as nvarchar(20),
@itemname as nvarchar(200),
@reqqty as money,
@unitcode as nvarchar(20),
@whcode as nvarchar(10),
@shelfcode as nvarchar(10),
@linenumber as int
as

set		dateformat dmy
set		language us_english

declare	@checkExist as int
set		@checkExist = (select isnull(count(itemcode),0) as vcount from npmaster.dbo.TB_NP_QueueRequestPicking where docno = @docno and docdate=@docdate and itemcode = @itemcode)

if		@checkExist = 0
begin
		insert	into npmaster.dbo.TB_NP_QueueRequestPicking(docno,docdate,itemcode,itemname,reqqty,unitcode,whcode,shelfcode,linenumber)
		values(@docno,@docdate,@itemcode,@itemname,@reqqty,@unitcode,@whcode,@shelfcode,@linenumber)
end

if		@checkExist <> 0
begin
		update	npmaster.dbo.TB_NP_QueueRequestPicking
		set		reqqty = @reqqty
		where	docno = @docno and docdate = @docdate and itemcode = @itemcode
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE Procedure dbo.USP_NP_InsertLabelTemp
@vItemCode as nvarchar(20),
@vBarCode as nvarchar(20),
@VName1 as nvarchar(200),
@vName2 as nvarchar(200),
@vQTY as int,
@vPriceLevel as money,
@vPrice as money,
@vUnitCode as nvarchar(20),
@vUsedUser as nvarchar(20),
@vCategory_ID as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(10),
@vVendorID as nvarchar(20),
@vRemark as nvarchar(10),
@vSPrice as money,
@vONHand as nvarchar(20),
@vQTYAllocate as nvarchar(20),
@vType as int ,
@vRemainOutQTY as nvarchar(20),
@vRemainInQTY as nvarchar(20),
@vSopNum as nvarchar(20),
@vSopDoc as nvarchar(20)
AS
set dateformat dmy
Insert 	into dbo.NP_LABEL_TEMP
	(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode,UsedUser,Category_ID,WHCode,ShelfCode,
	VENDR_ID,remark,SPrice,ONHAND,QTYALLOCATE,Type,RemainOutQTY,RemainInQTY,SopNum,SopDoc) 
values	(@vItemCode,@vBarcode,@vNAME1,@vNAME2,@vQTY,@vPriceLevel,@vPrice,@vUnitCode,@vUsedUser,@vCategory_ID,@vWHCode,@vShelfCode,
	@vVendorID,@vRemark,@vSPrice,@vONHAND,@vQTYALLOCATE,@vType,@vRemainOutQTY,@vRemainInQTY,@vSopNum,@vSopDoc)



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertLogPrintRunningRes 
@DocNo as varchar(20),
@UserPrint as varchar(80),
@GroupCode as int
as
insert into npmaster.dbo.TB_NP_RunningNumberReserveDocs(Docno,DocDate,UserPrint,GroupCode)
values(@Docno,getdate(),@UserPrint,@GroupCode)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertNPPrintQueue
@vJobID as nvarchar(10),
@vZoneID as nvarchar(10),
@vModuleID as nvarchar(10),
@vDocNo as nvarchar(10),
@vPrintStatus as int,
@vUserPrint as nvarchar(20)
as
set dateformat dmy

SET LOCK_TIMEOUT 10000

insert	into npmaster.dbo.TB_NP_CheckQueuePrint(JobID,ZoneID,ModuleID,DocNo,PrintStatus,UserPrint,DatePrint)
	values(@vJobID,@vZoneID,@vModuleID,@vDocNo,@vPrintStatus,@vUserPrint,getdate())
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertOpenItemMinuteLogs
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(200),
@vSaleQTY as int,
@vOnHandQTY as int,
@vUnitCode as nvarchar(20),
@vUserRequest as nvarchar(50),
@vUserOpen as nvarchar(50),
@vReasonDescription as nvarchar(150)
as
Declare @vOpenDateTime as nvarchar(20)
set	@vOpenDateTime = getdate()

set	dateformat dmy
set	language us_english
insert	into npmaster.dbo.TB_HMX_OpenItemMinuteLogs(DocDate,OpenDateTime,ItemCode,ItemName,SaleQTY,OnHandQTY,UnitCode,UserRequest,UserOpen,ReasonDescription)
values	(@vDocDate,@vOpenDateTime,@vItemCode,@vItemName,@vSaleQTY,@vOnHandQTY,@vUnitCode,@vUserRequest,@vUserOpen,@vReasonDescription)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertOrderPickHoldBill
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(20),
@vDatePicking as nvarchar(20),
@vBillType as smallint,
@vSOStatus as smallint,
@vIsCancel as smallint,
@vSaleCode as nvarchar(20),
@vCarLicense as nvarchar(20),
@vIsConditionSend as smallint,
@vSOCountNumber as int,
@vShelfGroup as nvarchar(10),
@vDueDate as nvarchar(20),
@vPickStatus as smallint,
@vSumOfItemAmount as money,
@vTaxAmount as money,
@vNetAmount as money,
@vUserID as nvarchar(50)

as

declare	@vCheckExist as int
set		dateformat dmy
set		@vCheckExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_QueueRequestPickingMaster where docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup)

if		@vCheckExist = 0 
begin
insert	into	npmaster.dbo.TB_NP_QueueRequestPickingMaster(DocNo,DocDate,ARCode,DatePicking,BillType,SOStatus,IsCancel,SaleCode,CarLicense,IsConditionSend,SOCountNumber,shelfgroup,DueDate,PickStatus,SumOfItemAmount,TaxAmount,NetAmount,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vARCode,@vDatePicking,@vBillType,@vSOStatus,@vIsCancel,@vSaleCode,@vCarLicense,@vIsConditionSend,@vSOCountNumber,@vShelfGroup,@vDueDate,@vPickStatus,@vSumOfItemAmount,@vTaxAmount,@vNetAmount,@vUserID,getdate())
end

/*
if		@vCheckExist > 0 
begin
update	npmaster.dbo.TB_NP_QueueRequestPickingMaster
set		CarLicense = @vCarLicense
where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup 

delete	npmaster.dbo.TB_NP_QueueRequestPicking where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber 
end
*/


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertOrderPickHoldBillSub
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vDatePicking as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(200),
@vReqQTY as money,
@vUnitCode as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(15),
@vZoneID as nvarchar(10),
@vIsCancel as smallint,
@vSelectItemDateTime as nvarchar(20),
@vSOCountNumber as int,
@vPrice as money,
@vDiscountAmount as money,
@vItemAmount as money,
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_QueueRequestPicking(DocNo,DocDate,DatePicking,ItemCode,ItemName,ReqQTY,UnitCode,WHCode,ShelfCode,ZoneID,IsCancel,SelectItemDateTime,SOCountNumber,Price,DiscountAmount,ItemAmount,LineNumber)
values(@vDocNo,@vDocDate,@vDatePicking,@vItemCode,@vItemName,@vReqQTY,@vUnitCode,@vWHCode,@vShelfCode,@vZoneID,@vIsCancel,@vSelectItemDateTime,@vSOCountNumber,@vPrice,@vDiscountAmount,@vItemAmount,@vLineNumber)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertPayGoods
@vInvoiceNo as nvarchar(20),
@vPaynumber as nvarchar(20),
@vWHCode as nvarchar(10),
@vUserPrint as nvarchar(20)

as
declare @vPayDatetime as nvarchar(20),
	@vLastPrintCount as int,
	@vChecked as int 

set	dateformat dmy
set	@vPayDatetime = getdate()
set	@vLastPrintCount = 1
set	@vChecked = 0 

insert 	npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked)
values (@vInvoiceNo,@vPaynumber,@vPayDatetime,@vWHCode,@vUserPrint,@vLastPrintCount,@vChecked)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertPayGoodsReserve
@vInvoiceNo as nvarchar(20),
@vPayNumber as nvarchar(20),
@vWHCode as nvarchar(10),
@vPrintDate as nvarchar(20),
@vReasonDesc as nvarchar(150),
@vReserveCode as nvarchar(20)
as
declare	@vPrintDateTime as nvarchar(20)
set	dateformat dmy
set	@vPrintDateTime = getdate()

insert into npmaster.dbo.TB_NP_PrintPayGoodsReserve(InvoiceNo,PayNumber,WHCode,PrintDate,PrintDateTime,ReasonDesc,ReserveCode) 
values(@vInvoiceNo,@vPayNumber,@vWHCode,@vPrintDate,@vPrintDateTime,@vReasonDesc,@vReserveCode)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertPickingDataLogs
@vSaleOrderNo as nvarchar(20),
@vPickingNo as nvarchar(20),
@vWHCode as nvarchar(5),
@vShelfGroup as nvarchar(2),
@vUserPrint  as nvarchar(20),
@vSaleCode1 as nvarchar(20)
as
set	dateformat dmy
insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount)
values(@vSaleOrderNo,@vPickingNo,getdate(),@vWHCode,@vShelfGroup,@vUserPrint,@vSaleCode1,1)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertPickingRequestMaster
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(30),
@vSaleCode as nvarchar(30),
@vRefNo as nvarchar(30),
@vMemberID as nvarchar(20),
@vBeforeTaxAmount as money,
@vTaxAmount as money,
@vNetDebtAmount as money,
@vIsConditionSend as int,
@vReqTime as nvarchar(20),
@vMyDescription as nvarchar(300),
@vUserID as nvarchar(30)

as

set		dateformat dmy

declare	@vExist as int

set	@vExist = (select isnull(count(docno),0) as vcount from npmaster.dbo.TB_NP_PickingRequestMaster where docno = @vDocNo)

if	@vExist = 0 
begin
insert	into npmaster.dbo.TB_NP_PickingRequestMaster(DocNo,DocDate,ARCode,SaleCode,RefNo,MemberID,BeforeTaxAmount,TaxAmount,NetDebtAmount,IsConditionSend,ReqTime,MyDescription,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vARCode,@vSaleCode,@vRefNo,@vMemberID,@vBeforeTaxAmount,@vTaxAmount,@vNetDebtAmount,@vIsConditionSend,@vReqTime,@vMyDescription,@vUserID,getdate())
end

if	@vExist > 0
begin
update	npmaster.dbo.TB_NP_PickingRequestMaster
set		ARCode = @vARCode,SaleCode=@vSaleCode,RefNo=@vRefNo,MemberID=@vMemberID,BeforeTaxAmount=@vBeforeTaxAmount,TaxAmount=@vTaxAmount,NetDebtAmount=@vNetDebtAmount,IsConditionSend=@vIsConditionSend,ReqTime=@vReqTime,MyDescription=@vMyDescription,LastEditorCode = @vUserID,LastEditDateTime = getdate()
where	docno = @vDocNo

delete	npmaster.dbo.TB_NP_PickingRequestSub
where	docno = @vDocNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertPickingRequestSub
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vQTY as money,
@vUnitcode as nvarchar(20),
@vPrice as money,
@vDisCountWord as nvarchar(20),
@vDisCountAmount as money,
@vNetAmount as money,
@vWHCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vShelfID as nvarchar(20),
@vZoneID as nvarchar(20),
@vBarCode as nvarchar(20),
@vLineNumber as int

as

set	dateformat dmy

insert	into npmaster.dbo.TB_NP_PickingRequestSub(DocNo,DocDate,ItemCode,QTY,Unitcode,Price,DisCountWord,DisCountAmount,NetAmount,WHCode,ShelfCode,ShelfID,ZoneID,BarCode,LineNumber)
values	(@vDocNo,@vDocDate,@vItemCode,@vQTY,@vUnitcode,@vPrice,@vDisCountWord,@vDisCountAmount,@vNetAmount,@vWHCode,@vShelfCode,@vShelfID,@vZoneID,@vBarCode,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

Create 	procedure dbo.USP_NP_InsertPrintLabel
@vItemCode as nvarchar(20),
@vBarCode as nvarchar(20),
@VName1 as nvarchar(200),
@vName2 as nvarchar(200),
@vQTY as int,
@vPriceLevel as money,
@vPrice as money,
@vUnitCode as nvarchar(20),
@vUsedUser as nvarchar(20),
@vCategory_ID as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(10),
@vVendorID as nvarchar(20),
@vRemark as nvarchar(10),
@vSPrice as money,
@vONHand as nvarchar(20),
@vRemainOutQTY as nvarchar(20),
@vRemainInQTY as nvarchar(20),
@vSopNum as nvarchar(20),
@vSopDoc as nvarchar(20)
as
set dateformat dmy
Insert 	Into dbo.TB_NP_PrintLabelTemp
	(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode,
	 VENDR_ID, remark, SPrice,OnHand,RemainOutQTY,RemainInQTY,SOPNUM,SOPDOC) 
values	(@vItemCode,@vBarcode,@vNAME1,@vNAME2,@vQTY,@vPriceLevel,@vPrice,@vUnitCode,@vUsedUser,@vCategory_ID,@vWHCode,@vShelfCode,
	@vVendorID,@vRemark,@vSPrice,@vONHAND,@vRemainOutQTY,@vRemainInQTY,@vSopNum,@vSopDoc)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQuePickCenterDriveInSub
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vWHCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vShelfID as nvarchar(20),
@vZoneID as nvarchar(20),
@vQTY as money,
@vPickQTY as money,
@vInvQTY as money,
@vUnitcode as nvarchar(20),
@vBarCode as nvarchar(20),
@vRefNo as nvarchar(30),
@vQueTime as int,
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_QuePickCenterSub(QueID,QueDocDate,ItemCode,WHCode,ShelfCode,ShelfID,ZoneID,QTY,PickQTY,OnCarQTY,InvQTY,Unitcode,BarCode,RefNo,QueTime,LineNumber)
values	(@vQueID,@vQueDocDate,@vItemCode,@vWHCode,@vShelfCode,@vShelfID,@vZoneID,@vQTY,@vQTY,@vQTY,0,@vUnitcode,@vBarCode,@vRefNo,@vQueTime,@vLineNumber)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQuePickCenterMaster
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(30),
@vSaleCode as nvarchar(30),
@vRefNo as nvarchar(30),
@vMemberID as nvarchar(20),
@vSourceID as int,
@vQueZone as nvarchar(10),
@vIsConditionSend as int,
@vQueReqTime as nvarchar(20),
@vQueTime as int

as

set		dateformat dmy

--QueID,QueDocDate,DocNo,DocDate,ARCode,SaleCode,RefNo,SourceID,CashierCode,HoldBillNo,IsConfiirm,Checker,CheckOutDateTime,QueDate,QuePicker,QueStart,QueStop,QueStatus,QueReqTime,QueReason,QueTime,IsCancel
declare	@vExist as int

set	@vExist = (select isnull(count(queid),0) as vcount from npmaster.dbo.TB_NP_QuePickCenterMaster where queid = @vQueID and QueDocDate =@vQueDocDate)

if	@vExist = 0 
begin
insert	into npmaster.dbo.TB_NP_QuePickCenterMaster(QueID,QueDocDate,DocNo,DocDate,ARCode,SaleCode,RefNo,MemberID,SourceID,QueZone,QueDate,IsConditionSend,QueReqTime,QueStatus,QueDescription,QueTime)
values	(@vQueID,@vQueDocDate,@vDocNo,@vDocDate,@vARCode,@vSaleCode,@vRefNo,@vMemberID,@vSourceID,@vQueZone,getdate(),@vIsConditionSend,@vQueReqTime,0,'รอจัด',@vQueTime)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQuePickCenterMasterDriveIn
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vDocNo as nvarchar(30),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(30),
@vSaleCode as nvarchar(30),
@vRefNo as nvarchar(30),
@vMemberID as nvarchar(20),
@vSourceID as int,
@vQueZone as nvarchar(10),
@vIsConditionSend as int,
@vQueReqTime as nvarchar(20),
@vQueTime as int

as

set		dateformat dmy

declare	@vExist as int

set	@vExist = (select isnull(count(queid),0) as vcount from npmaster.dbo.TB_NP_QuePickCenterMaster where queid = @vQueID and QueDocDate =@vQueDocDate)

if	@vExist = 0 
begin
insert	into npmaster.dbo.TB_NP_QuePickCenterMaster(QueID,QueDocDate,DocNo,DocDate,ARCode,SaleCode,RefNo,MemberID,SourceID,QueZone,QueDate,IsConditionSend,QueReqTime,QueStatus,QuePickStatus,QueRecStatus,QueReceived,QueDescription,QueTime)
values	(@vQueID,@vQueDocDate,@vDocNo,@vDocDate,@vARCode,@vSaleCode,@vRefNo,@vMemberID,@vSourceID,@vQueZone,getdate(),@vIsConditionSend,@vQueReqTime,2,1,1,1,'รับของแล้ว',@vQueTime)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQuePickCenterSub
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vWHCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vShelfID as nvarchar(20),
@vZoneID as nvarchar(20),
@vQTY as money,
@vPickQTY as money,
@vInvQTY as money,
@vUnitcode as nvarchar(20),
@vBarCode as nvarchar(20),
@vRefNo as nvarchar(30),
@vQueTime as int,
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_QuePickCenterSub(QueID,QueDocDate,ItemCode,WHCode,ShelfCode,ShelfID,ZoneID,QTY,PickQTY,InvQTY,Unitcode,BarCode,RefNo,QueTime,LineNumber)
values	(@vQueID,@vQueDocDate,@vItemCode,@vWHCode,@vShelfCode,@vShelfID,@vZoneID,@vQTY,0,0,@vUnitcode,@vBarCode,@vRefNo,@vQueTime,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQueueManagementSub
@vPickingNo as nvarchar(20),
@vItemCode as nvarchar(25),
@vItemName as nvarchar(150),
@vQTY as int,
@vPickQTY as int,
@vUnitCode as nvarchar(20),
@vPickItemStatus as nvarchar(4),
@vLineNumber as int,
@vTimeID as int,
@vDocDate as nvarchar(20) ,
@vRefNo as nvarchar(20)
as
set	dateformat dmy
insert	into npmaster.dbo.TB_NP_QueueManagementSub(PickingNo,ItemCode,ItemName,QTY,PickQTY,UnitCode,PickItemStatus,LineNumber,TimeID,docdate,RefNo)
values	(@vPickingNo,@vItemCode,@vItemName,@vQTY,@vPickQTY,@vUnitCode,@vPickItemStatus,@vLineNumber,@vTimeID,@vDocDate,@vRefNo)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQueueManagementSub1
@vPickingNo as nvarchar(20),
@vItemCode as nvarchar(25),
@vItemName as nvarchar(150),
@vWHCode as nvarchar(10),
@vQTY as int,
@vPickQTY as int,
@vUnitCode as nvarchar(20),
@vPickItemStatus as nvarchar(4),
@vLineNumber as int,
@vTimeID as int,
@vDocDate as nvarchar(20) ,
@vRefNo as nvarchar(20)
as
set	dateformat dmy
insert	into npmaster.dbo.TB_NP_QueueManagementSub(PickingNo,ItemCode,ItemName,WHCode,QTY,PickQTY,UnitCode,PickItemStatus,LineNumber,TimeID,docdate,RefNo)
values	(@vPickingNo,@vItemCode,@vItemName,@vWHCode,@vQTY,@vPickQTY,@vUnitCode,@vPickItemStatus,@vLineNumber,@vTimeID,@vDocDate,@vRefNo)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertQueueManagementSub2
@vPickingNo as nvarchar(20),
@vItemCode as nvarchar(25),
@vItemName as nvarchar(150),
@vWHCode as nvarchar(10),
@vQTY as int,
@vPickQTY as int,
@vUnitCode as nvarchar(20),
@vPickItemStatus as nvarchar(4),
@vLineNumber as int,
@vTimeID as int,
@vDocDate as nvarchar(20) ,
@vRefNo as nvarchar(20)
as
set	dateformat dmy
insert	into npmaster.dbo.TB_NP_QueueManagementSub_Test(PickingNo,ItemCode,ItemName,WHCode,QTY,PickQTY,UnitCode,PickItemStatus,LineNumber,TimeID,docdate,RefNo)
values	(@vPickingNo,@vItemCode,@vItemName,@vWHCode,@vQTY,@vPickQTY,@vUnitCode,@vPickItemStatus,@vLineNumber,@vTimeID,@vDocDate,@vRefNo)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertQueueSpeech
@vDocno	as int,
@vStatus as int,
@vPickingStatus as int,
@vZoneID as nvarchar(10)
as

set	dateformat dmy
insert 	into npmaster.dbo.tb_np_queuespeech(docno,status,pickingstatus,zoneid,speechdatetime)
values(@vDocno,@vStatus,@vPickingStatus,@vZoneID,getdate())

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertRequestQueueItem
@SaleOrderNo as nvarchar(20),
@SaleOrderDate as nvarchar(20),
@DocDate as nvarchar(20),
@ARCode as nvarchar(20),
@SaleCode as nvarchar(20),
@RequestDate as nvarchar(20),
@RequestTime as nvarchar(20),
@RequestStatus as smallint,
@RequestCountItem as money,
@RequestCountQTY as money,
@RequestFromPerson as nvarchar(50),
@RequestAt as smallint
as

declare  @PrintStatus as smallint
declare  @vCheckQueueExist as smallint
set	dateformat dmy
set	language us_english

set	@PrintStatus = 0
set	@vCheckQueueExist = (select isnull(count(saleorderno),0) as vCount  from npmaster.dbo.TB_NP_PickingQueueRequest where saleorderno = @SaleOrderNo and PrintStatus=0)

if 	@vCheckQueueExist = 0 
begin
insert 	into npmaster.dbo.TB_NP_PickingQueueRequest(SaleOrderNo,SaleOrderDate,DocDate,ARCode,SaleCode,RequestDate,RequestTime,RequestStatus,RequestCountItem,RequestCountQTY,PrintStatus,RequestFromPerson,RequestAt)
values(@SaleOrderNo,@SaleOrderDate,@DocDate,@ARCode,@SaleCode,@RequestDate,@RequestTime,@RequestStatus,@RequestCountItem,@RequestCountQTY,@PrintStatus,@RequestFromPerson,@RequestAt)
end

if 	@vCheckQueueExist = 1 
begin
update	npmaster.dbo.TB_NP_PickingQueueRequest
set	RequestDateHistory = RequestDate,RequestTimeHistory = RequestTime,RequestDate = @RequestDate,RequestTime = @RequestTime,LastEditRequestFrom=@RequestFromPerson,RequestAt = @RequestAt 
where 	saleorderno = @SaleOrderNo and PrintStatus = 0
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertScanItemShelfCode
@vItemCode as nvarchar(30),
@vBarCode as nvarchar(30),
@vItemName as nvarchar(300),
@vUnitCode as nvarchar(20),
@vWHCode as nvarchar(20),
@vZoneCode as nvarchar(20),
@vShelfCode as nvarchar(20),
@vUserScan as nvarchar(50),
@vModeScan as nvarchar(50)
as

declare	@vScanDateTime  as datetime
set	@vScanDateTime = getdate()

set	dateformat dmy

insert 	npmaster.dbo.NP_ScanBarCode_Logs (itemcode,barcode,itemname,unitcode,whcode,zonecode,shelfcode,scandatetime,UserScan,ModeScan)
values(@vItemCode,@vBarCode,@vItemName,@vUnitCode,@vWHCode,@vZoneCode,@vShelfCode,@vScanDateTime,@vUserScan,@vModeScan)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertSelectItemPicking
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vDatePicking as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(200),
@vReqQTY as money,
@vUnitCode as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(15),
@vZoneID as nvarchar(10),
@vIsCancel as smallint,
@vSelectItemDateTime as nvarchar(20),
@vSOCountNumber as int,
@vLineNumber as int

as

set		dateformat dmy

insert	into npmaster.dbo.TB_NP_QueueRequestPicking(DocNo,DocDate,DatePicking,ItemCode,ItemName,ReqQTY,UnitCode,WHCode,ShelfCode,ZoneID,IsCancel,SelectItemDateTime,SOCountNumber,LineNumber)
values(@vDocNo,@vDocDate,@vDatePicking,@vItemCode,@vItemName,@vReqQTY,@vUnitCode,@vWHCode,@vShelfCode,@vZoneID,@vIsCancel,@vSelectItemDateTime,@vSOCountNumber,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertSelectItemPickingMaster
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(20),
@vDatePicking as nvarchar(20),
@vBillType as smallint,
@vSOStatus as smallint,
@vIsCancel as smallint,
@vSaleCode as nvarchar(20),
@vCarLicense as nvarchar(20),
@vIsConditionSend as smallint,
@vSOCountNumber as int,
@vShelfGroup as nvarchar(10),
@vDueDate as nvarchar(20),
@vUserID as nvarchar(50)

as

declare	@vCheckExist as int
set		dateformat dmy
set		@vCheckExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_QueueRequestPickingMaster where docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup)

if		@vCheckExist = 0 
begin
insert	into	npmaster.dbo.TB_NP_QueueRequestPickingMaster(DocNo,DocDate,ARCode,DatePicking,BillType,SOStatus,IsCancel,SaleCode,CarLicense,IsConditionSend,SOCountNumber,shelfgroup,DueDate,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vARCode,@vDatePicking,@vBillType,@vSOStatus,@vIsCancel,@vSaleCode,@vCarLicense,@vIsConditionSend,@vSOCountNumber,@vShelfGroup,@vDueDate,@vUserID,getdate())
end


if		@vCheckExist > 0 
begin
update	npmaster.dbo.TB_NP_QueueRequestPickingMaster
set		CarLicense = @vCarLicense
where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup 

delete	npmaster.dbo.TB_NP_QueueRequestPicking where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber 
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InsertSelectItemPickingMaster1
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vARCode as nvarchar(20),
@vDatePicking as nvarchar(20),
@vBillType as smallint,
@vSOStatus as smallint,
@vIsCancel as smallint,
@vSaleCode as nvarchar(20),
@vCarLicense as nvarchar(20),
@vIsConditionSend as smallint,
@vSOCountNumber as int,
@vShelfGroup as nvarchar(10),
@vDueDate as nvarchar(20),
@vPickStatus as smallint,
@vUserID as nvarchar(50)

as

declare	@vCheckExist as int
set		dateformat dmy
set		@vCheckExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_QueueRequestPickingMaster where docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup)

if		@vCheckExist = 0 
begin
insert	into	npmaster.dbo.TB_NP_QueueRequestPickingMaster(DocNo,DocDate,ARCode,DatePicking,BillType,SOStatus,IsCancel,SaleCode,CarLicense,IsConditionSend,SOCountNumber,shelfgroup,DueDate,PickStatus,CreatorCode,CreateDateTime)
values	(@vDocNo,@vDocDate,@vARCode,@vDatePicking,@vBillType,@vSOStatus,@vIsCancel,@vSaleCode,@vCarLicense,@vIsConditionSend,@vSOCountNumber,@vShelfGroup,@vDueDate,@vPickStatus,@vUserID,getdate())
end


if		@vCheckExist > 0 
begin
update	npmaster.dbo.TB_NP_QueueRequestPickingMaster
set		CarLicense = @vCarLicense
where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber and shelfgroup = @vShelfGroup 

delete	npmaster.dbo.TB_NP_QueueRequestPicking where	docno = @vDocNo and docdate = @vDocDate and SOCountNumber = @vSOCountNumber 
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_InsertUpdatePrintLabeLogs
@vItemCode as nvarchar(30),
@vUnitCode as nvarchar(30),
@vBarCode as nvarchar(30),
@vLastUserPrinted as nvarchar(30),
@vLastFormPrint as nvarchar(200)

as

declare	@vCheckExist as int

set	dateformat dmy
set	@vCheckExist = (select isnull(count(itemcode),0) as vCount from npmaster.dbo.TB_NP_PrintLableLogs where itemcode = @vItemCode and unitcode = @vUnitCode)

if		@vCheckExist = 0
begin
insert	into npmaster.dbo.TB_NP_PrintLableLogs(ItemCode,UnitCode,BarCode,LastUserPrinted,LastPrintedDateTime,LastFormPrint)
values	(@vItemCode,@vUnitCode,@vBarCode,@vLastUserPrinted,getdate(),@vLastFormPrint)
end

if		@vCheckExist > 0
begin
update	npmaster.dbo.TB_NP_PrintLableLogs
set		BarCode=@vBarCode,LastUserPrinted=@vLastUserPrinted,LastPrintedDateTime=getdate(),LastFormPrint=@vLastFormPrint
where	itemcode = @vItemCode and unitcode =@vUnitCode
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InvoiceGroupWareHouse
@vDocNo as nvarchar(20)
as
select  a.docno,a.whcode,isnull(paynumber,'') as paynumber 
from 	bcarinvoicesub a 
	left join npmaster.dbo.np_paygoods b on a.docno = b.invoiceno and a.whcode = b.whcode
where 	a.docno = @vDocNo and paynumber = '' and checked = 0
group 	by a.docno,a.whcode,paynumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_InvoiceGroupWareHouseRes
@vDocNo as nvarchar(20),
@vZoneID as int
as
select  docno,whcode,invoiceno 
from 	npmaster.dbo.TB_QUE_CustItemReceipt 
where 	docno = @vDocno and iscancel = 0 and status = 1 and printstatus = 1 and zoneid = @vZoneID 
group 	by docno,whcode,invoiceno

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_ItemChagePriceLevel
as

set	dateformat dmy
select 	a.docno,scheduledate,itemcode,itemname,unitcode,newprice,oldprice,
	case pricelevel
	when 1 then 'ราคาที่1'
	when 2 then 'ราคาที่2'
	end as pricelevel,
	case saletype
	when 0 then
		case transsporttype 
		when 0 then
		'สดรับเอง'	
		when 1 then 
		'สดส่งให้'	
		end
	when 1 then
		case transsporttype 
		when 0 then
		'เชื่อรับเอง'	
		when 1 then 
		'เชื่อส่งให้'
		end	
	end as 	type
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster a
	left join npmaster.dbo.TB_NP_BasketItemUpdatePriceSub b on a.docno = b.docno and a.docdate = b.docdate
where 	isconfirm = 1 and isupdate = 0 and day(scheduledate) = day(getdate())and 
	month(scheduledate) = month(getdate())and 
	year(scheduledate) = year(getdate())
order	by a.docno,linenumber

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE Procedure dbo.USP_NP_LabelItemPriceLevel
@vItemCode as nvarchar(20),
@vUnitCode as nvarchar(20)
AS
set dateformat dmy
SELECT  TOP 100 PERCENT 
	a.ItemCode, a.UnitCode,isnull(b.UnitCode,'') as UnitCodePriceErect, 
	SalePrice1,isnull(b.PriceErect,0) as PriceErect,isnull(c.saleprice,0) as salepromotion
FROM   dbo.BPSPriceList a 
	left outer join dbo.BCPriceErect b on a.itemcode = b.itemcode
	left outer join dbo.bpspromoprice c on a.itemcode = c.itemcode
where 	a.itemcode = @vItemCode and a.unitcode = @vUnitCode
GROUP BY a.ItemCode, a.UnitCode,b.UnitCode,SalePrice1,b.PriceErect,c.saleprice
ORDER BY a.ItemCode




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE Procedure dbo.USP_NP_LabelItemVendor
@vItemCode as nvarchar(20)
AS
set dateformat dmy
SELECT DISTINCT TOP 100 PERCENT 
	isnull(isub.ApCode,'') AS VenderCode, isnull(ap.Name1,'') AS VenderName, 
	isnull(isub.ItemCode,'') AS ItemCode, isnull(item.Name1,'') AS ItemName, 
	isnull(BCBarCodeMaster.Barcode,'') AS Barcode, isnull(isub.UnitCode,'') AS UnitCode
FROM   dbo.BCAPINVOICESUB isub INNER JOIN
            dbo.BCAP ap ON isub.ApCode = ap.Code INNER JOIN
            dbo.BCITEM item ON isub.ItemCode = item.Code left outer JOIN  
	dbo.BCBarCodeMaster ON isub.ItemCode = dbo.BCBarCodeMaster.ItemCode
where 	isub.iscancel = 0 and isub.itemcode = @vItemCode




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_LabelPriceList_BarCode
@vBarCode as nvarchar(20)
/*@vItemName as nvarchar(50)*/
as
set	dateformat dmy

select	distinct a.Barcode, a.ItemCode,   b.Name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'') as ShelfCode,
	isnull(c.whcode,g.whcode) as WHCode,isnull(d.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion
FROM	dbo.BCBarCodeMaster a
	left join dbo.BCITEM b 		on a.ItemCode = b.Code
	left join dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join dbo.bcpriceerect e 	on a.itemcode = e.itemcode
	left join dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join dbo.bcitemwarehouse g on a.itemcode = g.itemcode  --and g.whcode in ('014','020')
WHERE     (a.ActiveStatus = 1) and a.barcode = @vBarCode and d.fromqty = 1
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,g.whcode,d.unitcode,c.whcode,
		d.SalePrice1,e.PriceErect,f.saleprice,d.fromqty
/*
if @vItemCode = '' and @vBarCode = '' and @vItemName <> ''

begin
SELECT	a.Barcode, a.ItemCode,   b.Name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'') as ShelfCode,
	isnull(c.whcode,'') as WHCode,isnull(d.unitcode,'') as UnitCode,d.SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion
FROM	dbo.BCBarCodeMaster a
	left JOIN	dbo.BCITEM b 		ON a.ItemCode = b.Code
	left join  	dbo.bcrecproduct c 	on a.itemcode = c.productcode
	left join    	dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join 	dbo.bcpriceerect e 	on a.itemcode = e.itemcode
	left join 	dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
WHERE     (a.ActiveStatus = 1) and b.name1 like '%'+@vItemName+'%'
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,c.whcode,d.unitcode,
		d.SalePrice1,e.PriceErect,f.saleprice
end*/
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_LabelPriceList_ItemCode
@vItemCode as nvarchar(20)
/*@vBarCode as nvarchar(20),
@vItemName as nvarchar(50)*/
as
set	dateformat dmy

/*if @vItemCode <> '' and @vBarCode = '' and @vItemName = ''
begin*/
select	distinct a.Barcode, a.ItemCode,   b.Name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'') as ShelfCode,
	isnull(c.whcode,g.whcode) as WHCode,isnull(a.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion
FROM	dbo.BCBarCodeMaster a
	left join	dbo.BCITEM b 		on a.ItemCode = b.Code
	--left join  	dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join  	npmaster.dbo.NP_ScanBarCode_Logs c 	on a.itemcode = c.itemcode
	left join 	dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join 	dbo.bcpriceerect e 	on a.itemcode = e.itemcode and a.unitcode = e.unitcode
	left join	dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join 	dbo.bcitemwarehouse g 	on a.itemcode = g.itemcode  --and g.whcode in ('014','020')
WHERE     (a.ActiveStatus = 1) and a.itemcode = @vItemCode and d.fromqty = 1 
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,c.whcode,a.unitcode,g.whcode,
		d.SalePrice1,e.PriceErect,f.saleprice,d.fromqty
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_LabelPriceList_ItemName
@vItemName as nvarchar(30)
as
set	dateformat dmy

select	distinct isnull(a.Barcode,'') as barcode, a.ItemCode,   b.Name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'') as ShelfCode,
	isnull(c.whcode,g.whcode) as WHCode,isnull(d.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion
FROM	dbo.BCBarCodeMaster a
	left JOIN	dbo.BCITEM b 		ON a.ItemCode = b.Code
	left join  	dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join    	dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join 	dbo.bcpriceerect e 	on a.itemcode = e.itemcode and a.unitcode = e.unitcode
	left join 	dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join dbo.bcitemwarehouse g on a.itemcode = g.itemcode  --and g.whcode in ('014','020')
WHERE     (a.ActiveStatus = 1) and b.name1 like '%'+@vItemName+'%' and d.fromqty = 1
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,g.whcode,d.unitcode,c.whcode,
		d.SalePrice1,e.PriceErect,f.saleprice,d.fromqty
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_NewInsertDataQueueManagement
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vDocType as int,
@vARCode as nvarchar(20),
@vSaleMan as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vRefDocNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(10),
@vZoneID as nvarchar(2),
@vTimeID as int,
@vStatus as int,
@vRequestTime as nvarchar(20),
@vCustomerZone as int

as

declare	@vStatusDesc as nvarchar(30)
declare	@vPickingStatus as int
declare	@vIsReceived as int
declare @vQueueDateTime as nvarchar(20)

set 	dateformat dmy
set	LANGUAGE us_english

set @vStatusDesc = 'รอจัด'
set @vPickingStatus = 0
set @vIsReceived = 0
set @vQueueDateTime = getdate()	

insert	into npmaster.dbo.TB_NP_QueueManagement
	(Docno,Docdate,DocType,ARCode,Status,StatusDesc,QueueDateTime,PickingStatus,IsReceived,SaleMan,SaleOrderNo,RefDocNo,WHCode,ShelfGroup,ZoneID,TimeID,StartDateTime,Picker,RequestTime,CustomerZone)
values 	(@vDocno,@vDocdate,@vDocType,@vARCode,@vStatus,@vStatusDesc,@vQueueDateTime,@vPickingStatus,@vIsReceived,@vSaleMan,@vSaleOrderNo,@vRefDocNo,@vWHCode,@vShelfGroup,@vZoneID,@vTimeID,getdate(),'OutLet',@vRequestTime,@vCustomerZone)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_NewInsertDataQueueManagement_New
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vDocType as int,
@vARCode as nvarchar(20),
@vSaleMan as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vRefDocNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(10),
@vZoneID as nvarchar(2),
@vTimeID as int,
@vStatus as int,
@vRequestTime as nvarchar(20),
@vCustomerZone as int,
@vTimePick as int

as

declare	@vStatusDesc as nvarchar(30)
declare	@vPickingStatus as int
declare	@vIsReceived as int
declare @vQueueDateTime as nvarchar(20)

set 	dateformat dmy
set	LANGUAGE us_english

set @vStatusDesc = 'รอจัด'
set @vPickingStatus = 0
set @vIsReceived = 0
set @vQueueDateTime = getdate()	

insert	into npmaster.dbo.TB_NP_QueueManagement
		(Docno,Docdate,DocType,ARCode,Status,StatusDesc,QueueDateTime,PickingStatus,IsReceived,SaleMan,SaleOrderNo,RefDocNo,WHCode,ShelfGroup,ZoneID,TimeID,StartDateTime,Picker,RequestTime,CustomerZone,TimePick)
values 	(@vDocno,@vDocdate,@vDocType,@vARCode,@vStatus,@vStatusDesc,@vQueueDateTime,@vPickingStatus,@vIsReceived,@vSaleMan,@vSaleOrderNo,@vRefDocNo,@vWHCode,@vShelfGroup,@vZoneID,@vTimeID,getdate(),'OutLet',@vRequestTime,@vCustomerZone,@vTimePick)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_NewInsertDataQueueManagement_Test
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vDocType as int,
@vARCode as nvarchar(20),
@vSaleMan as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vRefDocNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(10),
@vZoneID as nvarchar(2),
@vTimeID as int,
@vStatus as int,
@vRequestTime as nvarchar(20),
@vCustomerZone as int

as

declare	@vStatusDesc as nvarchar(30)
declare	@vPickingStatus as int
declare	@vIsReceived as int
declare @vQueueDateTime as nvarchar(20)

set 	dateformat dmy
set	LANGUAGE us_english

set @vStatusDesc = 'รอจัด'
set @vPickingStatus = 0
set @vIsReceived = 0
set @vQueueDateTime = getdate()	

insert	into npmaster.dbo.TB_NP_QueueManagement_Test
	(Docno,Docdate,DocType,ARCode,Status,StatusDesc,QueueDateTime,PickingStatus,IsReceived,SaleMan,SaleOrderNo,RefDocNo,WHCode,ShelfGroup,ZoneID,TimeID,StartDateTime,Picker,RequestTime,CustomerZone)
values 	(@vDocno,@vDocdate,@vDocType,@vARCode,@vStatus,@vStatusDesc,@vQueueDateTime,@vPickingStatus,@vIsReceived,@vSaleMan,@vSaleOrderNo,@vRefDocNo,@vWHCode,@vShelfGroup,@vZoneID,@vTimeID,getdate(),'OutLet',@vRequestTime,@vCustomerZone)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_OpenMinuteItem
@vStatus as int
as
if @vStatus = 1--เปิดติดลบ
begin
update dbo.bpsconfig set  checkstock =1  
end
if @vStatus = 2 --ปิดติดลบ
begin
update dbo.bpsconfig set  checkstock =0 
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE dbo.USP_NP_PaymentMoneySubUpdate 
@Return_Status as int,
@IsCompleteSave as int,
@vDocno as varchar(20),
@vDocdate as datetime,
@vItemCode as varchar(20),
@vItemName as varchar(50),
@vQTY as money,
@vUnitCode as varchar(20),
@vAmount as money,
@vIsCancel as int,
@vIsCompleteSave as int,
@vLineNumber as int
AS

set dateformat dmy
declare @ErrorNumber as int
declare @Return_StatusFinal as int
declare @ErrorDesc as varchar(30)

if @Return_Status = '1' --Error
begin
	SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
	ROLLBACK TRAN
	--SET IMPLICIT_TRANSACTIONS off --สับขาหลอก
	set @ErrorDesc = 'Error : บันทึกไม่สำเร็จ'
	return 1
end

if @vLineNumber = 0
begin
	delete  NPMaster.dbo.TB_NP_PaymentMoneySub where Docno = @vDocno
end

set @ErrorNumber = @@error
set @ErrorDesc = 'Error   :'+ cast(@ErrorNumber as varchar(25))
 
if @ErrorNumber <> 0
begin
	set @Return_StatusFinal = 1
	set @IsCompleteSave = 1
	goto  Commit_Rollback
end
else
begin
	set @Return_StatusFinal = 0
end

insert 	NPMaster.dbo.TB_NP_PaymentMoneySub(Docno,Docdate,ItemCode,ItemName,QTY,UnitCode,Amount,IsCancel,IsCompleteSave,LineNumber)
values	(@vDocno,@vDocdate,@vItemCode,@vItemName,@vQTY,@vUnitCode,@vAmount,@vIsCancel,@vIsCompleteSave,@vLineNumber)

if @vIsCompleteSave = 1
begin
	update NPMaster.dbo.TB_NP_PaymentMoneySub set IsCompleteSave = 1 where Docno = @vDocno
end

 if @@error <> 0
begin
	set @Return_StatusFinal = 1
	set @IsCompleteSave = 1
	goto  Commit_Rollback
end

Commit_Rollback:
if @IsCompleteSave = 1
begin
	if @Return_StatusFinal = 0  --ไม่มี Error  และบันทึกทั้งหมดสำเร็จ
	begin
		SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
		COMMIT TRAN
		--SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
		--SET IMPLICIT_TRANSACTIONS off --สับขาหลอก
		return 0
	end
	else
	begin
		SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
		ROLLBACK TRAN
		--SET IMPLICIT_TRANSACTIONS off --สับขาหลอก
		return 1
	end
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE dbo.USP_NP_PaymentMoneyUpdate
@ID as int=null,
@vDocno as varchar(20),
@vDocdate as datetime,
@vNetAmount as money,
@vIsCancel as int,
@vIsCompleteSave as int,
@vMyDescription as varchar(150),
@vUserID as varchar(50),
@vCreatorCode as varchar(50),
@vEditDatetime as datetime
AS

set dateformat dmy
declare @vCreateDatetime as datetime
set @vCreateDatetime = getdate()

if @ID is null

	begin
	insert 	NPMaster.dbo.TB_NP_PaymentMoney(Docno,Docdate,NetAmount,IsCancel,IsCompleteSave,UserID,MyDescription,CreatorCode,CreateDatetime)
	values	(@vDocno,@vDocdate,@vNetAmount,@vIsCancel,@vIsCompleteSave,@vUserID,@vMyDescription,@vCreatorCode,@vCreateDatetime)		
	end
else
	begin
	update NPMaster.dbo.TB_NP_PaymentMoney
		 set DocDate=@vDocDate,NetAmount =@vNetAmount,UserID = @vUserID,MyDescription = @vMyDescription,EditDatetime = @vEditDatetime
	where Docno = @vDocno
	end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_PickItemNotCreateBill

as

set		dateformat dmy

select	a.saleorderno as 'ใบสั่งขาย/จอง',queueno as 'คิวที่',a.docdate as 'วันที่คิว',a.timepick as 'สถานะจัด/จอง',a.isreceived as 'สถานะรับของ',
		a.pickingstatus as 'สถานะการจัด',a.timeid as 'ครั้งที่พิมพ์',a.statusdesc as 'รายละเอียดการจัด',issend as 'การจัดส่ง',
		a.requesttime as 'ต้องการรับของเวลา',a.arcode as 'รหัสลูกค้า',arname as 'ชื่อลูกค้า',a.picker as 'ชื่อพนักงานจัด',a.saleman as 'รหัสพนักงาน',salename as 'ชื่อพนักงาน',
		a.itemcode as 'รหัสสินค้า',a.itemname as 'ชื่อสินค้า',a.whcode as 'คลัง',a.shelfgroup as 'ชั้นเก็บ',a.qty as 'จำนวนที่ต้องการ',a.pickqty as 'จัดได้',a.unitcode as 'หน่วย',a.linenumber as 'ลำดับ'
from 
(
select	a.saleorderno,cast(a.docno as int) as queueno,a.docdate,a.timepick,a.isreceived,a.pickingstatus,a.shelfgroup,a.timeid,
		a.statusdesc,isnull(a.requesttime,'') as requesttime,a.arcode,isnull(a.picker,'') as picker,a.saleman,
		b.docno,b.sorefno,isnull(c.itemcode,'') as itemcode,isnull(c.itemname,'') as itemname,isnull(c.whcode,'') as whcode,
		isnull(c.qty,0) as qty,isnull(c.pickqty,0) as pickqty,isnull(c.unitcode,'') as unitcode,isnull(c.linenumber,0) as linenumber,
		isnull(d.name1,'') as arname,isnull(e.name,'') as salename,
		case f.isconditionsend 
		when 0 then 'รับเอง'
		when 1 then 'ส่งให้' end as issend
from	npmaster.dbo.tb_np_queuemanagement a
		left join dbo.bcarinvoice b on a.saleorderno = b.sorefno
		left join npmaster.dbo.tb_np_queuemanagementsub c on a.docno = c.pickingno and a.docdate = c.docdate and a.timeid = c.timeid
		left join dbo.bcar d on a.arcode = d.code
		left join dbo.bcsale e on a.saleman = e.code
		left join dbo.bcsaleorder f on a.saleorderno = f.docno
where	day(a.docdate) = day(getdate()) and month(a.docdate) = month(getdate()) and year(a.docdate) = year(getdate())  and 
		a.shelfgroup not in ('PKA','PKB') and a.timepick = 0
) as	a
where	docno is null
order	by saleorderno,queueno,linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_PrintLabel
@vUserID as nvarchar(20)
as
set dateformat dmy
select 	ItemCode,barcode,b.Name1,isnull(b.Name2,'') as Name2,isnull(a.Name1,'') as ItemName,
	isnull(QTY,0)as QTY,isnull(PriceLevel,0) as PriceLevel,isnull(Price,0) as Price,
	isnull(UnitCode,'') as UnitCode,isnull(UsedUser,'') as UsedUser,isnull(Category_ID,'') as Category_ID,
	isnull(WHCode,'') as WHCode,isnull(ShelfCode,'') as ShelfCode,isnull(VENDR_ID,'') as VENDR_ID,
	isnull(a.remark,'') as remark,isnull(SPrice,0) as SPrice,isnull(SOPNUM,'') as SOPNUM,isnull(SOPQUAN,'') as SOPQUAN,
	isnull(SOPDOC,'') as SOPDOC,isnull(SOPQUAD,'') as SOPQUAD,isnull(SOPREQS,'') as SOPREQS,isnull(SOPSALE,'') as SOPSALE,
	isnull(SOPCUST,'') as SOPCUST,isnull(SOPSHPM,'') as SOPSHPM,isnull(ONHAND,'') as ONHAND,--isnull(QTYALLOCATE,'') as QTYALLOCATE,
	isnull(a.RemainOutQTY,'') as RemainOutQTY,isnull(a.RemainInQTY,'') as RemainInQTY,isnull(ID,0) as ID,isnull(Type,0) as Type,
	isnull(c.department,'') as department,isnull(d.coderef,'') as coderef,isnull(b.shortname,'') as shortname,isnull(b.weight,0) as weight,isnull(avifilename,'') as shortcode,
	isnull(saleprice1,0) as volume1,isnull(saleprice2,0) as volume2,isnull(saleprice3,0) as volume3,
	isnull(saleprice4,0) as volumeqty1,isnull(saleprice5,0) as volumeqty2,isnull(saleprice6,0) as volumeqty3,
	isnull(saleprice7,0) as point1,isnull(saleprice8,0) as point2,isnull(saleprice9,0) as  point3
from 	dbo.tb_np_printlabeltemp a
	left join dbo.bcitem b on a.itemcode = b.code		
	left join dbo.BCItemCategory c on b.categorycode = c.code
	left join npmaster.dbo.TB_IC_CatDepartment d on isnull(c.department,'') = d.departmentcode
where useduser = @vUserID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_PrintReserveQueue
@vDocNo as nvarchar(20)
as
set	dateformat dmy
select 	docno,docdate,doctype,saleorderno,refdocno,whcode,shelfgroup,zoneid 
from 	npmaster.dbo.tb_np_queuemanagement
where 	day(docdate) = day(getdate())and month(docdate) = month(getdate()) and year(docdate) = year(getdate()) and docno = @vDocNo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_QueueCheckPickingItem
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20)
as

set	dateformat dmy

select 	pickingno as docno,docdate,itemcode,itemname,qty,pickqty,unitcode
from 	bchistory.dbo.TB_NP_QueueManagementsublogs 
where 	pickingno = @vDocno and docdate = @vDocDate
order	by linenumber 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* คิวที่จัดสินค้าไม่ครบ*/

CREATE	procedure dbo.USP_NP_QueueHaveRemainPicking
@vZoneID as int,
@vBegDate as nvarchar(20),
@vStopDate as nvarchar(20)
as

set	dateformat dmy

if @vZoneID = 0 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,'') as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
where 	status = 2 and pickingstatus = 2 and zoneid in ('01')  and docdate between  @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end
if @vZoneID = 1 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,saleman) as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
where 	status = 2 and pickingstatus = 2 and zoneid in ('02','03')and stopdatetime is null and docdate between @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* คิวที่จัดสินค้าที่ไม่หยุดเวลาจัดสินค้า*/

CREATE	procedure dbo.USP_NP_QueueItemNotStopTime
@vZoneID as int,
@vBegDate as nvarchar(20),
@vStopDate as nvarchar(20)
as

set	dateformat dmy

if @vZoneID = 0 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,'') as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code	
where 	status <> 2 and zoneid in ('01') and stopdatetime is null and docdate between @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end
if @vZoneID = 1 
begin
select 	saleorderno,docno,docdate,arcode,picker,saleman,whcode,shelfgroup,name1 as arname,isnull(c.name,saleman) as salename
from 	bchistory.dbo.TB_NP_QueueManagementlogs a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
where 	status <> 2 and zoneid in ('02','03')and stopdatetime is null and docdate between @vBegDate and @vStopDate
order	by docdate,cast(docno as int)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*ข้อมูลการจัดสินค้าตามคิว*/
CREATE	procedure dbo.USP_NP_QueueManagementByDocDate
@vBegDate as nvarchar(20),
@vEndDate as nvarchar(20)
as
set	dateformat dmy
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid ,convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a 
	left join (select pickingno as Docno,isnull(sum(pickqty),0) as SumQty
		from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
		where 	docdate between @vBegDate and @vEndDate
		group 	by pickingno)as b on a.docno = b.docno
where 	a.docdate between @vBegDate and @vEndDate
)as	Result
order	by docno
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*ข้อมูลการจัดสินค้าตามคิว*/
CREATE	procedure dbo.USP_NP_QueueManagementByNow
as
set	dateformat dmy
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid ,convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime
from 	NPMaster.dbo.TB_NP_QueueManagement a 
	left join (select pickingno as Docno,isnull(sum(pickqty),0) as SumQty
		from 	NPMaster.dbo.TB_NP_QueueManagementSub
		group 	by pickingno)as b on a.docno = b.docno
)as	Result
order	by docno
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_QueuePickingLaterTop5
@vZoneID as int
as

set	dateformat dmy

if @vZoneID = 0 
begin
select	* from
(
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1))) as datetime)and cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1))) as datetime)and cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00' 
)as a1
union
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2))) as datetime)and cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2))) as datetime)and cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00' 
)as a2
union
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3))) as datetime)and cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3))) as datetime)and cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00' 
)as a3
union
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4))) as datetime)and cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4))) as datetime)and cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00' 
)as a4
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5))) as datetime)and cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5))) as datetime)and cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00' 
)as a5
union
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6))) as datetime)and cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6))) as datetime)and cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a6
union
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7))) as datetime)and cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7))) as datetime)and cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a7
)as result
order	by docdate
end

if @vZoneID = 1 
begin
select	* from
(
select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1))) as datetime)and cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1))) as datetime)and cast(rtrim(day((getdate()-1)))+'/'+rtrim(month((getdate()-1)))+'/'+rtrim(year((getdate()-1)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a1
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2))) as datetime)and cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2))) as datetime)and cast(rtrim(day((getdate()-2)))+'/'+rtrim(month((getdate()-2)))+'/'+rtrim(year((getdate()-2)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a2
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3))) as datetime)and cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3))) as datetime)and cast(rtrim(day((getdate()-3)))+'/'+rtrim(month((getdate()-3)))+'/'+rtrim(year((getdate()-3)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a3
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4))) as datetime)and cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4))) as datetime)and cast(rtrim(day((getdate()-4)))+'/'+rtrim(month((getdate()-4)))+'/'+rtrim(year((getdate()-4)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a4
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5))) as datetime)and cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5))) as datetime)and cast(rtrim(day((getdate()-5)))+'/'+rtrim(month((getdate()-5)))+'/'+rtrim(year((getdate()-5)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a5
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6))) as datetime)and cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6))) as datetime)and cast(rtrim(day((getdate()-6)))+'/'+rtrim(month((getdate()-6)))+'/'+rtrim(year((getdate()-6)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a6
union
select	* from
(
select 	
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7))) as datetime)and cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7)))as datetime)
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in  ('02','03')  and status = 2 and a.docdate between cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7))) as datetime)and cast(rtrim(day((getdate()-7)))+'/'+rtrim(month((getdate()-7)))+'/'+rtrim(year((getdate()-7)))as datetime)
	 and convert(nvarchar(30),(stopdatetime-startdatetime),8 ) > '00:15:00'
)as a7
)as result
order	by docdate 
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_QueuePickingNotFully
@vZoneID as int,
@vBegDate as nvarchar(20),
@vEndDate as nvarchar(20)
as

set	dateformat dmy

if @vZoneID = 0 
begin

select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,itemcode,itemname,qty,pickqty,unitcode,
	case isreceived when 0 then 'ลูกค้ายกเลิกการซื้อขาย' else 'ลูกค้ารับของตามจำนวน' end as isreceived
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join BCHistory.dbo.TB_NP_QueueManagementSubLogs b 	on a.docno = b.pickingno and a.docdate = b.docdate and qty <> pickqty
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and pickingstatus = 2 and a.docdate between @vBegDate and @vEndDate
)as a1
order	by docdate
end

if @vZoneID = 1 
begin

select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,itemcode,itemname,qty,pickqty,unitcode,
	case isreceived when 0 then 'ลูกค้ายกเลิกการซื้อขาย' else 'ลูกค้ารับของตามจำนวน' end as isreceived
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join BCHistory.dbo.TB_NP_QueueManagementSubLogs b 	on a.docno = b.pickingno and a.docdate = b.docdate and qty <> pickqty
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('02','03') and pickingstatus = 2  and a.docdate between @vBegDate and @vEndDate
)as a1
order	by docdate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_QueuePickingUnUsual
@vZoneID as int,
@vBegDate as nvarchar(20),
@vEndDate as nvarchar(20)
as

set	dateformat dmy

if @vZoneID = 0 
begin

select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('01') and status = 2 and a.docdate between @vBegDate and @vEndDate
	and convert(nvarchar(30),(stopdatetime-startdatetime),8 )< '00:00:30'
)as a1
order	by docdate
end

if @vZoneID = 1 
begin

select	* from
(
select 	 
	a.Docno,a.Docdate,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartDateTime,a.StopDateTime,isnull(a.Picker,'') as Picker,a.SaleOrderNo,a.WHCode,
	c.name1 as arname,isnull(d.name,'') as salename,
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) as PickingTime,SumQty
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a
	left join (	select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
			from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
			where 	docdate between @vBegDate and @vEndDate
			group 	by pickingno,docdate
		) b 	on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	zoneid in ('02','03') and status = 2 and a.docdate between @vBegDate and @vEndDate
	and convert(nvarchar(30),(stopdatetime-startdatetime),8 )< '00:00:30'
)as a1
order	by docdate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_QuotaionInsertDetails
@vReturnStatus as smallint,
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(200),
@vQTY as int,
@vPrice as money,
@vUnitCode as nvarchar(20),
@vDisCountAmount as money,
@vSumDiscountAmount as money,
@vAmount as money,
@vIsCancel as smallint,
@vLineNumber as int
as
set dateformat dmy
declare @vReturnStatusFinal as smallint 

if @vReturnStatus = 1 
begin
	set implicit_transactions on
	rollback tran
	return 1
end

insert 	npmaster.dbo.TB_NP_QuotationSub (DocNo,DocDate,ItemCode,ItemName,QTY,Price,UnitCode,DisCountAmount,SumDiscountAmount,Amount,IsCancel,LineNumber)
select 	@vDocNo,@vDocDate,@vItemCode,@vItemName,@vQTY,@vPrice,@vUnitCode,@vDisCountAmount,@vSumDiscountAmount,@vAmount,@vIsCancel,@vLineNumber

if @@Error <> 0
begin
	set @vReturnStatusFinal = 1
	goto  Commit_Rollback
end
else
	set @vReturnStatusFinal = 0

Commit_Rollback:

if @vReturnStatusFinal = 0  --ไม่มี Error  และบันทึกทั้งหมดสำเร็จ
	begin
		SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
		COMMIT TRAN
		Return 0
	end
else
	begin
		SET IMPLICIT_TRANSACTIONS on --สับขาหลอก
		ROLLBACK TRAN
		Return 1
	end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_QuotaionInsertHeader
@vIsOpen as smallint,
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vBillType as smallint,
@vArCode as nvarchar(20),
@vCreditDay as smallint,
@vValidate as nvarchar(20),
@vIsConditionSend as smallint, 
@vSaleCode as nvarchar(20),
@vTaxRate as money,
@vIsCancel as smallint,
@vSumOfItemAmount as money,
@vTaxAmount as money,
@vDiscountAmount as money,
@vTotalAmount as money, 
@vNetAmount as money,
@vMyDescription as nvarchar(200),
@vLogIN as nvarchar(20)
as
set dateformat dmy
declare @vCountDocno as smallint
declare @vCreatorCode as nvarchar(20)
declare @vCreateDateTime as nvarchar(20)
declare @vLastEditorCode as nvarchar(20)
declare @vLastEditDateT as nvarchar(20)

set @vCountDocno = (select count(docno) as CountDoc from npmaster.dbo.TB_NP_Quotation where docno = @vDocNo)

set implicit_transactions on

if @vIsOpen = 0 and @vCountDocno = 0
	begin
	select  @vCreatorCode = @vLogIN,@vCreateDateTime = getdate()
	select @vLastEditorCode = null,@vLastEditDateT = null
	end
else
	begin
	select @vCreatorCode = CreatorCode,@vCreateDateTime = CreateDateTime from npmaster.dbo.TB_NP_Quotation where docno = @vDocNo
	select @vLastEditorCode = @vLogIN,@vLastEditDateT = getdate()
	delete npmaster.dbo.TB_NP_Quotation where docno = @vDocNo
	delete npmaster.dbo.TB_NP_QuotationSub where docno = @vDocNo
	end
	insert  npmaster.dbo.TB_NP_Quotation (DocNo,DocDate,BillType,ArCode,CreditDay,Validate,IsConditionSend, 
	            SaleCode,TaxRate,IsCancel,SumOfItemAmount,TaxAmount,DiscountAmount,TotalAmount, 
		NetAmount,Mydescription,CreatorCode,CreateDateTime,LastEditorCode,LastEditDateT)
	select @vDocNo,@vDocDate,@vBillType,@vArCode,@vCreditDay,@vValidate,@vIsConditionSend, 
	            @vSaleCode,@vTaxRate,@vIsCancel,@vSumOfItemAmount,@vTaxAmount,@vDiscountAmount,@vTotalAmount, 
		@vNetAmount,@vMyDescription,@vCreatorCode,@vCreateDateTime,@vLastEditorCode,@vLastEditDateT

if @@error <> 0 
	return 1
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_Quotation
@DocNo varchar(20)
as

IF @DocNo = '' 
BEGIN
	RAISERROR ('ไม่ได้กำหนดเงื่อนไขเลขที่เอกสาร',16,1)
	return 0
END

set dateformat dmy 

SELECT  top 100 percent 
	a.DocNo,a.DocDate,a.BillType,a.ArCode,a.CreditDay,a.validity as Validate,getdate()+a.validity as Validate1,
	a.SaleCode,a.TaxRate,a.MyDescription1 as MyDescription,a.SumOfItemAmount,a.DiscountAmount,a.TaxAmount , 
	a.TotalAmount,a.NetAmount,a.CreatorCode,a.CreateDateTime, 
	b.ItemCode,b.ItemName,b.Qty,b.Price,b.DiscountAmount as DiscountAmountsub,b.Amount,
	b.UnitCode,b.LineNumber,c.name1 as ARName,isnull(b.mydescription,'') as mydescriptionsub,
	c.BillAddress,isnull(c.Telephone,'') as Telephone,isnull(c.Fax,'') as Fax,isnull(d.name,'') as salename,
	isnull(e.telephone,'-') as salephone,isnull(e.mobile,'-') as salemobile,isnull(e.email,'-') as saleemail,BeforeTaxAmount
FROM   --npmaster.dbo.tb_np_quotation a 
	dbo.BCQuotation a
	--inner join npmaster.dbo.TB_NP_QuotationSub b ON a.DocNo = b.DocNo 
	inner join dbo.BCQuotationSub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
	LEFT OUTER JOIN dbo.BCAR c ON a.ARCode = c.Code
	Left outer join dbo.bcsale d on a.salecode = d.code
	left outer join npmaster.dbo.tb_np_salecontact e on a.salecode = e.code
WHERE     (a.IsCancel = 0) and a.docno = @Docno
order by a.docdate,a.docno
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE dbo.USP_NP_QuotationNewDocNo 
@vTypeDocno as smallint
AS
declare @Head as varchar(3)
declare @DocYear varchar(2)
declare @NewDocNo as varchar(20)
declare @NotDup as int 
declare @count as int
declare @CountPR as int
declare @DocMonth as varchar(2)
declare @DocDay as varchar(2)

if 	@vTypeDocno = 0 
	begin
	set @Head = 'QNV'
	end
else
	begin
	set @Head = 'QNC'
	end

set @NotDup = 1 --มี Record ซ้ำ
set @count = 1

set @DocYear = substring(dbo.FT_CG_ThaiYear(getdate()),3,2)
set @DocMonth = dbo.FT_CG_MonthForRunNumber(getdate())
--set @DocDay = dbo.FT_CG_DayForRunNumber(getdate())

while @NotDup = 1
begin
	--กำหนดรหัสเอกสารใหม่
	set @NewDocNo = 

	(select ltrim(rtrim(@Head)) + ltrim(rtrim(@DocYear)) + ltrim(rtrim(@DocMonth))+'-'+
		case when len(cast(Running as varchar(4))) = 1 then '000' + rtrim(cast(Running as varchar(1))) else
			case when len(cast(Running as varchar(4))) = 2 then '00' + rtrim(cast(Running as varchar(2))) else
				case when len(cast(Running as varchar(4))) = 3 then '0' + rtrim(cast(Running as varchar(3))) else
						case when len(cast(Running as varchar(4))) = 4 then  rtrim(cast(Running as varchar(4))) 
					end
				end
			end
		end
		as NewDocNo
	from 
		(select 	cast(isnull(substring(max(docno),11,4),0) as int)+1 as Running
		from 	NPMaster.dbo.tb_np_quotation where left(docno,3) = @Head ) as a)

	set @CountPR = (select count(*)  from NPMaster.dbo.tb_np_quotation  where DocNo = @NewDocNo)
	if @CountPR = 0
		set  @NotDup = 0

	set @Count = @Count + 1
end
select @NewDocNo as NewDocNo
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_QuotationSearchItem
@ItemCode as nvarchar(20)
as

set	dateformat dmy


select 	distinct code,isnull(name1,'') as name1,isnull(stockqty,0) as stockqty,isnull(defstkunitcode,'') as stkdefunitcode,isnull(unitcode,'') as unitcode,isnull(defsaleunitcode,'') as defsaleunitcode
from 	dbo.bcitem a
	inner join bcpricelist b on a.code = b.itemcode 
where	code = @ItemCode and activestatus = 1
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_QuotationSelectPriceList
@ItemCode as nvarchar(20),
@UnitCode as nvarchar(20),
@SaleType as smallint,
@TranSportType as smallint,
@PriceLevel as smallint

as

set	dateformat dmy

if	@PriceLevel = 1
begin
	select 	code,isnull(name1,'') as name1,isnull(unitcode,'') as unitcode,isnull(b.saleprice1,0)  as price
	from 	dbo.bcitem a
		left join dbo.bcpricelist b on a.code = b.itemcode 
	where	code = @ItemCode and activestatus = 1 and stopdate > cast(getdate()as datetime) and
		saletype = @SaleType and transporttype = @TranSportType and unitcode = @UnitCode
end

if	@PriceLevel = 2
begin
	select 	code,isnull(name1,'') as name1,isnull(unitcode,'') as unitcode,isnull(b.saleprice2,0)  as price
	from 	dbo.bcitem a
		left join dbo.bcpricelist b on a.code = b.itemcode 
	where	code = @ItemCode and activestatus = 1 and stopdate > cast(getdate()as datetime) and
		saletype = @SaleType and transporttype = @TranSportType and unitcode = @UnitCode
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_ReportQueueSumPickQTYPerPicker
@vType as int
as

set	dateformat dmy

if 	@vType = 0 
begin
select	picker,sum(sumqty) as totalqty,sum(DiffPicking)as totaltime,
	case 
	when sum(sumqty) > 0 then sum(DiffPicking)/sum(sumqty)
	when sum(sumqty) <= 0 then 0 end as totalitemeverage
from
(
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	a.docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid 
from 	NPMaster.dbo.TB_NP_QueueManagement a 
	left join (select 	pickingno as docno,docdate,isnull(sum(pickqty),0) as SumQty 
			from 	NPMaster.dbo.TB_NP_QueueManagementSub
			group	by pickingno,docdate)as b on a.docno = b.docno
where 	 zoneid in ('01')
)as	Result
)as 	Result1
group	by picker
end

if 	@vType = 1 
begin
select	picker,sum(sumqty) as totalqty,sum(DiffPicking)as totaltime,
	case 
	when sum(sumqty) > 0 then sum(DiffPicking)/sum(sumqty)
	when sum(sumqty) <= 0 then 0 end as totalitemeverage
from
(
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	a.docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid 
from 	NPMaster.dbo.TB_NP_QueueManagement a 
	left join (select 	pickingno as docno,docdate,isnull(sum(pickqty),0) as SumQty 
			from 	NPMaster.dbo.TB_NP_QueueManagementSub
			group	by pickingno,docdate)as b on a.docno = b.docno
where 	 zoneid in ('02','03')
)as	Result
)as 	Result1
group	by picker
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_ReportQueueSumPickQTYPerPickerByDocDate
@vType as int,
@vBegDate as nvarchar(20),
@vEndDate as nvarchar(20)
as

set	dateformat dmy

if 	@vType = 0 
begin
select	docdate,picker,sum(sumqty) as totalqty,sum(DiffPicking)as totaltime,
	case 
	when sum(sumqty) > 0 then sum(DiffPicking)/sum(sumqty)
	when sum(sumqty) <= 0 then 0 end as totalitemeverage
from
(
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	a.docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid 
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a 
	left join (select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
		from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
		where 	docdate between @vBegDate and @vEndDate
		group 	by pickingno,docdate)as b on a.docno = b.docno and a.docdate = b.docdate
where 	a.docdate between @vBegDate and @vEndDate and zoneid in ('01')
)as	Result
)as 	Result1
group	by picker,docdate
end

if 	@vType = 1 
begin
select	docdate,picker,sum(sumqty) as totalqty,sum(DiffPicking)as totaltime,	
	case 
	when sum(sumqty) > 0 then sum(DiffPicking)/sum(sumqty)
	when sum(sumqty) <= 0 then 0 end as totalitemeverage
from
(
select	case 
	when sumqty >0 then (diffpicking/sumqty)
	when sumqty <=0 then 0
	end  as pickitemeverage,*
from
(
select 	cast(a.docno as int) as docno,isnull(b.sumqty,0) as SumQTY,isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,
	a.docdate,startdatetime,stopdatetime,arcode,doctype,status,picker,isreceived,saleman,
	saleorderno,whcode,shelfgroup,zoneid,timeid 
from 	BCHistory.dbo.TB_NP_QueueManagementLogs a 
	left join (select pickingno as Docno,docdate,isnull(sum(pickqty),0) as SumQty
		from 	BCHistory.dbo.TB_NP_QueueManagementSubLogs
		where 	docdate between @vBegDate and @vEndDate
		group 	by pickingno,docdate)as b on a.docno = b.docno and a.docdate = b.docdate
where 	a.docdate between @vBegDate and @vEndDate and zoneid in ('02','03')
)as	Result
)as 	Result1
group	by picker,docdate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_ReserveItemQtyDetails
@vSearch  as nvarchar(100)
as

set		dateformat dmy

if		@vSearch = '' 
begin
		select	distinct a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 
		order	by a.itemcode
end

if		@vSearch <> ''
begin
		select	distinct *
		from
		(
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and a.itemcode like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and f.name1 like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and b.shelfcode like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and c.arcode like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and d.name1 like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and c.salecode like '%'+@vSearch+'%'
union
		select	a.itemcode as stkitemcode,a.whcode as stkwhcode,a.reserveqty,a.unitcode as stkunit,
				b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.remainqty,b.unitcode,
				c.docno,c.docdate,c.arcode,c.duedate,c.salecode,isnull(c.mydescription,'') as mydescription,
				isnull(deliveryday,0) as deliveryday,isnull(c.deliverydate,'') as deliverydate,
				isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.name1,'') as itemnamemaster
				,dbo.FT_CHK_GL_TextToMoney(a.reserveqty) as reserveqtyText
		from	dbo.bcstkwarehouse a
				left join (	select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,sum(qty) as qty,sum(remainqty)as remainqty,unitcode 
							from dbo.bcsaleordersub 
							where remainqty <> 0 
							group by docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.whcode = b.whcode and a.itemcode = b.itemcode
				left join dbo.bcsaleorder c on b.docno = c.docno and b.docdate = c.docdate and b.arcode = c.arcode
				left join dbo.bcar d on c.arcode = d.code
				left join dbo.bcsale e on c.salecode = e.code
				left join dbo.bcitem f on b.itemcode = f.code
		where	a.reserveqty <> 0 and c.iscancel = 0 and c.sostatus = 1 and e.name like '%'+@vSearch+'%'
		) as
		result
		order	by itemcode
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SaleItemHistory
@vArCode as nvarchar(20),
@vSaleCode as nvarchar(20),
@vItemCode as nvarchar(20)
as

set 	dateformat dmy
if @vARCode <> '' and @vSaleCode = '' and  @vItemCode = ''
	begin
	select 	a.docno,a.docdate,a.itemcode,isnull(b.name1,'')  as itemname,a.arcode,a.salecode,
		a.whcode,a.qty,unitcode,a.price,a.discountamount,
		a.amount,a.netamount,isnull(a.sorefno,'') sorefno,case posstatus 
		when 2 then 'ยิงขาย'
		else 'ออกบิล' end as PosStatus,
		c.name1,
		d.name
	from 	dbo.bcarinvoicesub a 
		inner join dbo.bcitem b on a.itemcode = b.code
		inner join dbo.bcar c on a.arcode = c.code 
		inner join dbo.bcsale d on a.salecode = d.code
	where 	a.iscancel = 0 and a.arcode = @vARCode
	order	by a.docdate,a.itemcode
	end
else if @vARCode <> '' and  @vSaleCode <> '' and @vItemCode = ''
	begin
	select 	a.docno,a.docdate,a.itemcode,isnull(b.name1,'')  as itemname,a.arcode,a.salecode,
		a.whcode,a.qty,unitcode,a.price,a.discountamount,
		a.amount,a.netamount,isnull(a.sorefno,'') sorefno,case posstatus 
		when 2 then 'ยิงขาย'
		else 'ออกบิล' end as PosStatus,
		c.name1,
		d.name
	from 	dbo.bcarinvoicesub a 
		inner join dbo.bcitem b on a.itemcode = b.code
		inner join dbo.bcar c on a.arcode = c.code 
		inner join dbo.bcsale d on a.salecode = d.code
	where 	a.iscancel = 0 and a.arcode = @vARCode and a.salecode = @vSaleCode
	order	by a.docdate,a.itemcode
	end
else if @vARCode <> '' and @vSaleCode <> '' and @vItemCode <> ''
	begin
	select 	a.docno,a.docdate,a.itemcode,isnull(b.name1,'')  as itemname,a.arcode,a.salecode,
		a.whcode,a.qty,unitcode,a.price,a.discountamount,
		a.amount,a.netamount,isnull(a.sorefno,'') sorefno,case posstatus 
		when 2 then 'ยิงขาย'
		else 'ออกบิล' end as PosStatus,
		c.name1,
		d.name
	from 	dbo.bcarinvoicesub a 
		inner join dbo.bcitem b on a.itemcode = b.code
		inner join dbo.bcar c on a.arcode = c.code 
		inner join dbo.bcsale d on a.salecode = d.code
	where 	a.iscancel = 0 and a.arcode = @vARCode and a.salecode = @vSaleCode and a.itemcode = @vItemCode
	order	by a.docdate,a.itemcode
	end
else if @vARCode <> '' and @vSaleCode = '' and @vItemCode <> ''
	begin
	select 	a.docno,a.docdate,a.itemcode,isnull(b.name1,'')  as itemname,a.arcode,a.salecode,
		a.whcode,a.qty,unitcode,a.price,a.discountamount,
		a.amount,a.netamount,isnull(a.sorefno,'') sorefno,case posstatus 
		when 2 then 'ยิงขาย'
		else 'ออกบิล' end as PosStatus,
		c.name1,
		d.name
	from 	dbo.bcarinvoicesub a 
		inner join dbo.bcitem b on a.itemcode = b.code
		inner join dbo.bcar c on a.arcode = c.code 
		inner join dbo.bcsale d on a.salecode = d.code
	where 	a.iscancel = 0 and a.arcode = @vARCode and a.itemcode = @vItemCode
	order	by a.docdate,a.itemcode
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SeaechUserLogIn
@vUserID as nvarchar(30)
as
select	comment+'-'+ name as salename,code
from 	npmaster.dbo.TB_NP_BCUserID
where	code = @vUserID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create procedure dbo.USP_NP_SearchArCode
@ARCode as varchar(20)
as
select code,name1 from bcnp.dbo.bcar where code = @ARCode

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_SearchArCodeLike
@ARCode as varchar(20)
as
select * from
(
	select code,name1 from bcnp.dbo.bcar where code like '%'+@ARCode+'%' 
	union
	select code,name1 from bcnp.dbo.bcar where name1 like '%'+@ARCode+'%' 
) as Result 
order	by name1

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchBarCode 
@vSearch as nvarchar(50)
as

set	dateformat dmy
select	*
from
(
select 	a.itemcode,a.barcode,isnull(b.name1,'')  as itemname,isnull(a.unitcode,'') as UnitCode,
		isnull(defsalewhcode,'') as defsalewhcode,isnull(defsaleshelf,'')as defsaleshelf,isnull(stockqty,0) as stockqty,isnull(remainoutqty,0) as remainoutqty,
		isnull((select top 1 saleprice1 from dbo.bcpricelist where saletype = 0 and transporttype = 0 and itemcode = b.code and unitcode = a.unitcode),0) as price,
		isnull(c.zoneid,'X') as zone
from 	dbo.bcbarcodemaster a
		left join dbo.bcitem b on a.itemcode = b.code
		left join npmaster.dbo.TB_NP_CategoryPickingZone c on b.categorycode = c.categorycode 
where 	a.barcode like '%'+@vSearch+'%' and a.activestatus = 1 and b.activestatus = 1

union

select 	a.itemcode,a.barcode,isnull(b.name1,'')  as itemname,isnull(a.unitcode,'') as UnitCode,
		isnull(defsalewhcode,'') as defsalewhcode,isnull(defsaleshelf,'')as defsaleshelf,isnull(stockqty,0) as stockqty,isnull(remainoutqty,0) as remainoutqty,
		isnull((select top 1 saleprice1 from dbo.bcpricelist where saletype = 0 and transporttype = 0 and itemcode = b.code and unitcode = a.unitcode),0) as price,
		isnull(c.zoneid,'X') as zone
from 	dbo.bcbarcodemaster a
		left join dbo.bcitem b on a.itemcode = b.code 
		left join npmaster.dbo.TB_NP_CategoryPickingZone c on b.categorycode = c.categorycode 
where 	b.code like '%'+@vSearch+'%' and a.activestatus = 1 and b.activestatus = 1

union

select 	a.itemcode,a.barcode,isnull(b.name1,'')  as itemname,isnull(a.unitcode,'') as UnitCode,
		isnull(defsalewhcode,'') as defsalewhcode,isnull(defsaleshelf,'')as defsaleshelf,isnull(stockqty,0) as stockqty,isnull(remainoutqty,0) as remainoutqty,
		isnull((select top 1 saleprice1 from dbo.bcpricelist where saletype = 0 and transporttype = 0 and itemcode = b.code and unitcode = a.unitcode),0) as price,
		isnull(c.zoneid,'X') as zone
from 	dbo.bcbarcodemaster a
		left join dbo.bcitem b on a.itemcode = b.code 
		left join npmaster.dbo.TB_NP_CategoryPickingZone c on b.categorycode = c.categorycode 
where 	b.name1 like '%'+@vSearch+'%' and a.activestatus = 1 and b.activestatus = 1
)	as	result
order 	by itemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchChangePrice
@vDocNo as nvarchar(30)
as

set	dateformat dmy
select 	a.docno,a.docdate,a.scheduledate,isnull(itemcode,'') as itemcode,
	isnull(itemname,'') as itemname,isnull(unitcode,'') as unitcode,isnull(pricelevel,0) as pricelevel,
	isnull(saletype,0) as saletype,isnull(transsporttype,0) as transsporttype,
	isnull(newprice,0) as newprice,isnull(oldprice ,0) oldprice
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster a
	left join npmaster.dbo.TB_NP_BasketItemUpdatePriceSub b on a.docno = b.docno and a.docdate = b.docdate
where 	a.docno = @vDocNo
order	by a.docdate,a.docno,linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchChangePriceDocNo
@vType as smallint,
@vSearch as nvarchar(30)
as

set	dateformat dmy

if	@vType = 0
begin
select 	* 
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster 
order	by Docdate,DocNo
end

if @vType = 1
begin
select	* from
(
select 	* 
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster 
where 	DocNo like '%'+@vSearch+'%'
union
select 	* 
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster 
where 	creatorcode like '%'+@vSearch+'%'
)	as Result
order	by docdate,docno
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchCheckCountSOPicking
@vDocNo as nvarchar(20)
as

declare	@vCheckPickingTimes as int
declare	@vSetTimes as int

set		dateformat dmy

set	@vCheckPickingTimes = (select isnull(max(SOCountNumber),0) as sorefno from npmaster.dbo.TB_NP_QueueRequestPickingMaster where docno = @vDocNo)
set	@vSetTimes = @vCheckPickingTimes+1

select @vSetTimes as vCount
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchCheckOutHolding
@vSearch as nvarchar(100)
as

set		dateformat dmy

if		@vSearch = ''
begin
select	DocNo,DocDate,isnull(CashierCode,'') as CashierCode,TotalAmount,NetDebtAmount,isnull(MyDescription ,'') as MyDescription
from	bcnpdisa.dbo.BPSHoldingBill 
where	docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by docno
end

if	@vSearch <> ''
begin
select	*
from
(
select	DocNo,DocDate,isnull(CashierCode,'') as CashierCode,TotalAmount,NetDebtAmount,isnull(MyDescription ,'') as MyDescription
from	bcnpdisa.dbo.BPSHoldingBill 
where	docno like '%'+@vSearch+'%' and docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
union
select	DocNo,DocDate,isnull(CashierCode,'') as CashierCode,TotalAmount,NetDebtAmount,isnull(MyDescription ,'') as MyDescription
from	bcnpdisa.dbo.BPSHoldingBill 
where	MyDescription like '%'+@vSearch+'%' and docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
) as	result
order	by docno
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchCheckPickStatus
@vReserveNo as nvarchar(20)
as

set		dateformat dmy

select	top 1 pickstatus,socountnumber
from
(	
select	docno,pickstatus,socountnumber
from	npmaster.dbo.tb_np_queuerequestpickingmaster 
where	docno = @vReserveNo 
union
select	docno,pickstatus,socountnumber
from	bchistory.dbo.tb_np_queuerequestpickingmaster_2008 
where	docno = @vReserveNo 
) as	a
order	by socountnumber desc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchCheckShelfSaleOrderData
@vDocNo as nvarchar(20)
as

set		dateformat dmy
set		language us_english

/*
select distinct a.docno,
	isnull(left(case 
	when b.whcode = '014' and y.itemtype is not null and b.shelfcode = 'BAK' then 'K'
	when b.whcode = '014' and y.itemtype is not null and b.shelfcode <> 'BAK' then 'M'
	when b.whcode = '014' and y.itemtype is  null  then 'H'
	when b.whcode = '020' then 'H'
	when b.whcode = '016' then 'Y'
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
	end,1),'D') as shelfgroup
from	dbo.bcsaleorder a 
	inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno =@vDocNo and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
	inner join dbo.bcar c on a.arcode = c.code
	left join dbo.bcitem g on b.itemcode = g.code
	left join (select distinct productcode,whcode,(select top 1 shelfcode from dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from dbo.bcrecproduct2) as aaa )z 
	on b.itemcode = z.productcode  and b.whcode = z.whcode
	left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype
where	a.iscancel = 0 and a.billstatus = 0 and a.docno = @vDocNo
*/

select	distinct a.docno,b.whcode,b.shelfcode as shelfgroup
from	dbo.bcsaleorder a
	inner join dbo.bcsaleordersub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode 
where	a.billstatus = 0 and a.docno = @vDocNo and a.iscancel = 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchDataQueueDetails
@vDocno as nvarchar(10),
@vZoneID as nvarchar(5)
as

set	dateformat dmy

declare	@vStatus as int
set	@vStatus = (select isnull((select distinct status from npmaster.dbo.tb_np_queuemanagement where docno = @vDocno),0)  as status)

if 	@vZoneID = 1
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('01')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('01')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('01')
	end
end

if 	@vZoneID = 2
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('02','03')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('02','03')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and zoneid in ('02','03')
	end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchDataQueueDetails1
@vDocno as nvarchar(10),
@vZoneID as nvarchar(5),
@vDocDate as nvarchar(20)
as

set	dateformat dmy

declare	@vStatus as int
set	@vStatus = (select isnull((select distinct status from npmaster.dbo.tb_np_queuemanagement where docno = @vDocno and docdate = @vDocDate),0)  as status)

if 	@vZoneID = 1
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
end

if 	@vZoneID = 2
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchDataQueueDetails2
@vDocno as nvarchar(10),
@vZoneID as nvarchar(5),
@vDocDate as nvarchar(20)
as

--TEST
set	dateformat dmy

declare	@vStatus as int

set	@vStatus = (select isnull((select distinct status from npmaster.dbo.tb_np_queuemanagement where docno = @vDocno and docdate = @vDocDate),0)  as status)

if 	@vZoneID = 1
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('01')
	end
end

if 	@vZoneID = 2
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and zoneid in ('02','03')
	end
end

if 	@vZoneID = 3
begin
	if 	@vStatus = 0
	begin
		select 	docno,docdate,arcode,'' as picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,0 as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and shelfgroup = 'AVL' --and zoneid in ('03') and shelfgroup <> 'PKB'
	end
	
	if 	@vStatus = 1
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,getdate()) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and shelfgroup = 'AVL' --and zoneid in ('03') and shelfgroup <> 'PKB'
	end
	
	if 	@vStatus = 2
	begin
		select 	docno,docdate,arcode,picker,whcode,saleorderno,isnull(b.name1,'') as arname,saleman,isnull(c.name,'') as salename,status,
			isnull(round(cast(datediff(second,startdatetime,stopdatetime) as decimal(10,2))/cast(60 as decimal(10,2)),2),0) as DiffPicking,isreceived,doctype,timeid,
			case
			when status = 0 then
			'00:00:00'
			when status = 1 then
			convert(nvarchar(30),(getdate()-startdatetime),8 ) 
			when status = 2 then
			convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
			end as PickingTime,isnull(a.mydescription,'') as mydescription
		from 	npmaster.dbo.tb_np_queuemanagement a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.saleman = c.code
		where 	docno = @vDocno and docdate = @vDocDate and shelfgroup = 'AVL' --and zoneid in ('03') and shelfgroup <> 'PKB'
	end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchDocQueuePrint
as

/*
set	dateformat dmy
select 	top 1 jobid,zoneid,moduleid,docno,a.reportid,a.reporttype,printstatus,isnull(userprint,'') as userprint,isnull(dateprint,getdate()) as dateprint,
	isnull(b.reportname,'') as reportname,case printstatus when 0 then 'ยังไม่ได้พิมพ์' else 'พิมพ์แล้ว' end as PrintText,
	case jobid 
	when '01' then 'พิมพ์ใบจัดสินค้า Picking Slip'
	when '02' then 'พิมพ์ใบจัดสินค้า จากใบโอน' end as PrintDocument
from 	npmaster.dbo.TB_NP_CheckQueuePrint a left join 
	dbo.bcreportname b on a.reportid = b.repid and a.reporttype = b.reptype
where 	printstatus = 0 and jobid in ('01','02')

*/

set	dateformat dmy

SET LOCK_TIMEOUT 10000 --ให้ lock object  ได้ 10 วิ

select 	top 1 jobid,a.zoneid,a.moduleid,a.docno,a.reportid,a.reporttype,printstatus,isnull(userprint,'') as userprint,isnull(dateprint,getdate()) as dateprint,
	case printstatus when 0 then 'ยังไม่ได้พิมพ์' else 'พิมพ์แล้ว' end as PrintText,
	case jobid 
	when '01' then 'พิมพ์ใบจัดสินค้า Picking Slip'
	when '02' then 'พิมพ์ใบจัดสินค้า จากใบโอน' end as PrintDocument,isnull(c.shelfgroup,'O') as shelfgroup
from 	npmaster.dbo.TB_NP_CheckQueuePrint a 
	inner join npmaster.dbo.TB_NP_Queuemanagement c on a.docno = c.docno and year(c.docdate) = year(getdate())and month(c.docdate) = month(getdate())and day(c.docdate) = day(getdate())
where 	printstatus = 0 and jobid in ('01') and year(getdate()) = year(dateprint) and month(getdate()) = month(dateprint) and day(getdate()) = day(dateprint)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchDocnoUpdatePrice
as

set	dateformat dmy
select 	docno,docdate,scheduledate,creatorcode
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster 
where 	isconfirm = 1  and 
	docno in (select distinct docno from npmaster.dbo.TB_NP_BasketItemUpdatePricesub where isupdate = 0) and 
	day(scheduledate) = day(getdate()) and month(scheduledate) = month(getdate())and year(scheduledate) = year(getdate())
order	by docno



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchDriveInCheckOut
@vDocNo as nvarchar(30)
as

set		dateformat dmy

select	a.DocNo,a.DocDate,a.Checker,a.PosNo,a.NetDebtAmount,a.IsCancel,a.IsConfirm,ItemCode,QTY,BillQTY,UnitCode,Price,Amount,BarCode,b.IsCancel as IsCancelLine,LineNumber
from	npmaster.dbo.TB_NP_DriveInCheckOut a
		left join npmaster.dbo.TB_NP_DriveInCheckOutSub b on a.docno = b.docno and a.docdate = b.docdate
where	a.docno = @vDocNo and a.iscancel = 0 and b.iscancel = 0
order	by b.linenumber

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchDriveInDetails
@vDocNo as nvarchar(30)
as

set		dateformat dmy

select	a.DocNo,a.DocDate,isnull(a.ARCode,'') as arcode,isnull(a.MemberID,'') as memberid,isnull(a.SaleCode,'') as salecode,
		isnull(a.RefNo,'') as refno,isnull(PickZone,'') as PickZone,isnull(BeforeTaxAmount,0)as BeforeTaxAmount,isnull(TaxAmount,0) as TaxAmount,
		isnull(TotalNetAmount,0) as TotalNetAmount,a.IsCancel,a.IsMerge,a.isconfirm,isnull(issendque,0)as issendque,
		b.ItemCode,b.ItemName,b.WHCode,b.ShelfCode,isnull(b.ShelfID,'') as ShelfID,isnull(QTY,0) as qty,isnull(b.zoneid,'') as zoneid,
		b.UnitCode,b.Price,isnull(b.DisCountWord,'') as DisCountWord,DisCountAmount,Amount,b.IsCancel as iscancelsub,isnull(b.BarCode,'') as barcode,
		isnull(c.name1,'') as arname,isnull(d.name,'') as salename
from	npmaster.dbo.TB_NP_DriveInSlipMaster a
		left join npmaster.dbo.TB_NP_DriveInSlipSub b on a.DocNo = b.docno and a.DocDate = b.docdate
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code 
where	a.docno = @vDocNo
order	by b.linenumber



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchGroupPicking
@vType as int,
@vDocNo as nvarchar(30)
as

set		dateformat dmy

if	@vType = 1
begin
select	distinct a.docno,zoneid
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join npmaster.dbo.TB_NP_PickingRequestSub b on a.docno = b.docno and a.docdate = b.docdate
where	a.docno = @vDocNo
end


if	@vType = 2
begin
select	distinct a.docno,b.zoneid
from	npmaster.dbo.TB_NP_QueueRequestPickingMaster a 
		left join npmaster.dbo.TB_NP_QueueRequestPicking b on a.docno = b.docno and a.docdate = b.docdate
where	a.docno = @vDocNo
end

if	@vType =3
begin
select	distinct a.docno,b.zoneid
from	npmaster.dbo.TB_NP_DriveInSlipMaster a 
		left join npmaster.dbo.TB_NP_DriveInSlipSub b on a.docno = b.docno and a.docdate = b.docdate
where	a.docno = @vDocNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchHodingBill
@vSearch as nvarchar(50)
as

set		dateformat dmy

if		@vSearch = ''
begin
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by a.docno
end

if		@vSearch <> ''
begin
select	*
from 
(
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) and
		a.docno like '%'+@vSearch+'%'
union
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		a.arcode like '%'+@vSearch+'%'
union
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		a.salecode like '%'+@vSearch+'%'
union
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		c.name1 like '%'+@vSearch+'%'
union
select	distinct 
		a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,a.salecode,isnull(d.name,'') as salename,
		a.cashiercode,isnull(e.name,'') as cashiername,a.machineno,a.netdebtamount
from	bcnpdisa.dbo.bpsholdingbill a
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
where	a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		d.name like '%'+@vSearch+'%'
) as	a
order	by docno
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchHodingBillDetails
@vDocNo as nvarchar(30)

as

set		dateformat dmy

select	a.docno,a.docdate,a.arcode,a.salecode,a.cashiercode,a.shiftno,a.machineno,a.machinecode,a.sumofitemamount,a.discountamount,a.taxamount,
		a.totalamount,b.itemcode,isnull(f.name1,'') as itemname,b.whcode,b.shelfcode,b.qty,b.price,b.discountword,b.discountamount,b.amount,b.unitcode,b.sorefno,
		b.linenumber,b.barcode,isnull(c.name1,'') as arname,isnull(d.name,'') as salename,1 as rate1,1 as rate2,
		isnull((select top 1 sourceid from npmaster.dbo.TB_NP_QuePickCenterMaster e where b.sorefno = docno order by queid desc),0) as type
from	bcnpdisa.dbo.bpsholdingbill a 
		left join bcnpdisa.dbo.bpsholdingbillsub b on a.docno = b.docno and a.docdate = b.docdate 
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code 
		left join dbo.bcsale e on a.cashiercode = e.code 
		left join dbo.bcitem f on b.itemcode = f.code
where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchHoldingDetails
@vDocNo as nvarchar(50)
as

set		dateformat dmy

select	a.DocNo,a.DocDate,isnull(a.CashierCode,'') as CashierCode,a.TotalAmount,a.NetDebtAmount,isnull(a.MyDescription ,'') as MyDescription,
	b.ItemCode,d.itemname,b.WHCode,b.ShelfCode,b.Qty,b.Price,b.Amount,b.NetAmount,b.UnitCode,b.StockType,b.LineNumber,b.BarCode,b.BillTime,b.PosStatus,
	c.docno as driveinno,c.docdate as driveindate,c.id,c.refid,c.pickzone,c.totalnetamount as driveinamount,d.pickqty,d.confirmqty,d.invqty
from	bcnpdisa.dbo.BPSHoldingBill a 
	inner join bcnpdisa.dbo.BPSHoldingBillsub b on a.docno = b.docno and a.docdate = b.docdate  
	inner join npmaster.dbo.tb_np_driveinslipmaster c on a.docno = c.billposno and a.docdate = c.docdate
	inner join npmaster.dbo.tb_np_driveinslipsub d on c.docno = d.docno and c.docdate = d.docdate and b.itemcode = d.itemcode and b.unitcode = d.unitcode
where	a.docno = @vDocNo and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by a.docno
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchItemChangePrice 
--@vSecMan as nvarchar(20),
@vDocDate as nvarchar(20)
as 

set	dateformat dmy
set	language us_english

SET LOCK_TIMEOUT 25000

SELECT  top 100 percent
	a.itemcode, 
	isnull(z.barcode,'')as barcode,
	b.Name1, 
	b.Name2, 
	a.unitcode,
	d.SalePrice1,
	isnull(e.PriceErect,0) as PriceErect, 
	cast(ltrim(str(day(dateupdate)))+'/'+ltrim(str(month(dateupdate)))+'/'+ltrim(str(year(dateupdate))) as datetime)as dateupdate1 ,
	a.dateupdate,
	b.defsalewhcode,
	a.printedUpdate,
	isnull(f.whcode,'') as whcode,isnull(f.shelfcode,'') as shelfcode,isnull(g.secman,'N/A')as secman
FROM    npmaster.dbo.nppricehistory a 
left  	join dbo.bcbarcodemaster z on a.itemcode = z.itemcode and a.unitcode = z.unitcode and z.activestatus = 1
left 	join dbo.BCITEM b ON a.ItemCode = b.Code 
left   	join dbo.BCPriceList d on a.itemcode = d.itemcode and a.unitcode = d.unitcode and 
	a.saletype = d.saletype and a.transporttype = d.transporttype
left 	join dbo.bcpriceerect e on a.itemcode = e.itemcode and a.unitcode = e.unitcode
left	join dbo.bcrecproduct2 f on a.itemcode = f.productcode and a.unitcode = f.unitcode
left	join bchistory.dbo.TB_IC_DrilldownHmxShelf g on a.itemcode = g.itemcode 
WHERE  oldprice1 <> newprice1 and 
	dateupdate < getdate() and a.saletype = 0 and a.transporttype = 0 and 
	a.PrintedUpDate = 0 and b.activestatus = 1 and f.whcode in ('012','014','020')  and a.activestatus = 1 and --isnull(g.secman,'') = @vSecMan and 
	cast(ltrim(str(day(dateupdate)))+'/'+ltrim(str(month(dateupdate)))+'/'+ltrim(str(year(dateupdate)))as datetime)= @vDocDate
group 	by  a.dateupdate,a.itemcode,z.barcode,b.Name1,b.Name2,a.unitcode,
	d.SalePrice1,e.PriceErect,b.defsalewhcode,a.printedUpdate,f.whcode,f.shelfcode,g.secman
order by a.dateupdate desc ,a.itemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchItemDetails
@vItemCode as nvarchar(20),
@vUnitCode as nvarchar(20),
@vBarCode as nvarchar(20),
@vWHCode as nvarchar(10)
as
/*set	dateformat dmy
SELECT	distinct isnull(a.Barcode,'') as barcode, isnull(a.ItemCode,'') as itemcode, isnull(b.Name1,'') as name1,isnull(b.Name2,'') as Name2,isnull(i.SHELFCODE,'-') as ShelfCode,
	isnull(i.whcode,h.whcode) as WHCode,isnull(a.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion,
	isnull(b.CategoryCode,'') as CategoryCode,isnull(b.DefStkUnitCode,'') as DefStkUnitCode, 
	isnull(b.ReserveQty,0) as ReserveQty,isnull(b.RemainOutQty,0) as RemainOutQty, 
	isnull(b.RemainInQty,0) as RemainInQty,isnull(g.Qty,0) as QTY,i.scandatetime
FROM	dbo.BCBarCodeMaster a
	left join dbo.BCITEM b 		on a.ItemCode = b.Code
	left join dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join dbo.bcpriceerect e 	on a.itemcode = e.itemcode
	left join dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join dbo.BCSTKLocation g 	on b.Code = g.ItemCode and d.unitcode = g.unitcode and c.whcode = g.whcode and g.shelfcode <> 'DMG' 
	left join dbo.bcitemwarehouse h on a.itemcode = h.itemcode and  h.whcode not in ('011')
--	left join npmaster.dbo.NP_ScanBarCode_Logs i on c.productcode = i.itemcode and c.whcode = i.whcode  
	left join (select top 1 itemcode,whcode,shelfcode,scandatetime from npmaster.dbo.NP_ScanBarCode_Logs i where i.itemcode = @vItemCode and i.whcode = @vWHCode order by scandatetime desc) i on c.productcode = i.itemcode  

WHERE     (a.ActiveStatus = 1) and a.itemcode = @vItemCode and a.unitcode = @vUnitCode and a.barcode = @vBarCode and i.whcode = @vWHCode
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,i.SHELFCODE,h.whcode,a.unitcode,
		d.SalePrice1,e.PriceErect,f.saleprice,b.CategoryCode,b.DefStkUnitCode,c.whcode,
		b.ReserveQty,b.RemainOutQty,b.RemainInQty,g.qty,i.scandatetime,i.whcode
order	by i.scandatetime desc
GO*/

set	dateformat dmy
select	distinct
	isnull(a.Barcode,'') as barcode, isnull(a.ItemCode,'') as itemcode, isnull(b.Name1,'') as name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,h.SHELFCODE) as ShelfCode,
	isnull(c.whcode,h.whcode) as WHCode,isnull(a.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion,
	isnull(b.CategoryCode,'') as CategoryCode,isnull(b.DefStkUnitCode,'') as DefStkUnitCode, 
	isnull(b.ReserveQty,0) as ReserveQty,isnull(b.RemainOutQty,0) as RemainOutQty, 
	isnull(b.RemainInQty,0) as RemainInQty,isnull(g.Qty,0) as QTY,c.updatedatetime
FROM	dbo.BCBarCodeMaster a
	left join dbo.BCITEM b 		on a.ItemCode = b.Code
	left join(select	top 1 * from
		(
		select itemcode,whcode,shelfcode,scandatetime as updatedatetime  from npmaster.dbo.NP_ScanBarCode_Logs where itemcode=@vItemCode
		union
		select productcode as itemcode,whcode,shelfcode,'' as updatedatetime from dbo.bcrecproduct2 where productcode= @vItemCode
		) as a
		where whcode =@vWHCode
		order by updatedatetime  desc) c on a.itemcode = c.itemcode
	--left join dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join dbo.bcpriceerect e 	on a.itemcode = e.itemcode and a.unitcode = e.unitcode
	left join dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join dbo.BCSTKLocation g 	on b.Code = g.ItemCode and d.unitcode = g.unitcode and c.whcode = g.whcode and g.shelfcode <> 'DMG' 
	left join dbo.bcitemwarehouse h on a.itemcode = h.itemcode and  h.whcode not in ('011')
	--left join npmaster.dbo.NP_ScanBarCode_Logs i on c.productcode = i.itemcode and c.whcode = i.whcode  
	--left join (select top 1 itemcode,whcode,shelfcode,scandatetime from npmaster.dbo.NP_ScanBarCode_Logs i where i.itemcode = '5001007' and i.whcode = '014' order by scandatetime desc) i on c.productcode = i.itemcode  

WHERE     (a.ActiveStatus = 1) and a.itemcode =@vItemCode and a.unitcode =@vUnitCode and a.barcode = @vBarCode and c.whcode =@vWHCode
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,h.whcode,a.unitcode,
		d.SalePrice1,e.PriceErect,f.saleprice,b.CategoryCode,b.DefStkUnitCode,c.whcode,
		b.ReserveQty,b.RemainOutQty,b.RemainInQty,g.qty,c.updatedatetime,c.shelfcode,h.shelfcode
order	by c.updatedatetime desc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchItemDetails_Market
@vItemCode as nvarchar(20),
@vUnitCode as nvarchar(20),
@vBarCode as nvarchar(20),
@vWHCode as nvarchar(10)
as
set	dateformat dmy
SELECT	distinct isnull(a.Barcode,'') as barcode, isnull(a.ItemCode,'') as itemcode, isnull(b.Name1,'') as name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'-') as ShelfCode,
	isnull(c.whcode,h.whcode) as WHCode,isnull(a.unitcode,'') as UnitCode,isnull(d.SalePrice1,0) as SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion,
	isnull(b.CategoryCode,'') as CategoryCode,isnull(b.DefStkUnitCode,'') as DefStkUnitCode, 
	isnull(b.ReserveQty,0) as ReserveQty,isnull(b.RemainOutQty,0) as RemainOutQty, 
	isnull(b.RemainInQty,0) as RemainInQty,isnull(g.Qty,0) as QTY
FROM	dbo.BCBarCodeMaster a
	left join dbo.BCITEM b 		on a.ItemCode = b.Code
	left join(select	top 1 * from
		(
		select itemcode,whcode,shelfcode,scandatetime as updatedatetime  from npmaster.dbo.NP_ScanBarCode_Logs where itemcode=@vItemCode
		union
		select productcode as itemcode,whcode,shelfcode,'' as updatedatetime from dbo.bcrecproduct2 where productcode= @vItemCode
		) as a
		where whcode =@vWHCode
		order by updatedatetime  desc) c on a.itemcode = c.itemcode
--	left join dbo.bcrecproduct2 c 	on a.itemcode = c.productcode
	left join dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join dbo.bcpriceerect e 	on a.itemcode = e.itemcode and a.unitcode = e.unitcode
	left join dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
	left join dbo.BCSTKLocation g 	on b.Code = g.ItemCode and d.unitcode = g.unitcode and c.whcode = g.whcode and g.shelfcode <> 'DMG' 
	left join dbo.bcitemwarehouse h on a.itemcode = h.itemcode and  h.whcode not in ('011') 
WHERE     (a.ActiveStatus = 1) and a.itemcode = @vItemCode and a.unitcode = @vUnitCode and a.barcode = @vBarCode and c.whcode = @vWHCode
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,h.whcode,a.unitcode,
		d.SalePrice1,e.PriceErect,f.saleprice,b.CategoryCode,b.DefStkUnitCode,c.whcode,
		b.ReserveQty,b.RemainOutQty,b.RemainInQty,g.qty
order	by a.barcode,c.whcode desc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchItemPriceDetails
@vItemCode as nvarchar(30),
@vUnitCode as nvarchar(30)
as
set	dateformat dmy
select 	itemcode,unitcode,saletype,transporttype,isnull(saleprice1,0) as saleprice1,isnull(saleprice2,0) as saleprice2
from 	dbo.bcpricelist 
where 	itemcode = @vItemCode and remark <> 'promotion' and unitcode = @vUnitCode
order	by itemcode,saletype,transporttype


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchItemSite
@vItemCode as nvarchar(20),
@vWHCode as nvarchar(10)
as
set	dateformat dmy
SELECT	TOP 100 PERCENT 
	dbo.BCITEM.Code AS ItemCode, dbo.BCITEM.Name1 AS ItemName, isnull(dbo.BCITEM.Name2,'') as name2, 
	isnull(dbo.BCITEM.CategoryCode,'') as CategoryCode,isnull(dbo.BCITEM.DefStkUnitCode,'') as DefStkUnitCode, 
	isnull(dbo.BCITEM.ReserveQty,0) as ReserveQty,isnull(dbo.BCITEM.RemainOutQty,0) as RemainOutQty, 
	isnull(dbo.BCITEM.RemainInQty,0) as RemainInQty,isnull(dbo.bcrecproduct.WHCode,'') as whcode, 
	isnull(dbo.bcrecproduct.ShelfCode,'') as shelfcode,isnull(dbo.BCSTKLocation.Qty,0) as QTY,isnull(dbo.BCITEM.StockQty,0) as stockqty
FROM  	dbo.BCITEM 
	left  outer join dbo.BCSTKLocation ON dbo.BCITEM.Code = dbo.BCSTKLocation.ItemCode and dbo.BCstklocation.shelfcode <> 'DMG' 
	left outer  join dbo.bcrecproduct on dbo.BCSTKLocation.ItemCode = dbo.bcrecproduct.productcode  and  
	dbo.BCstklocation.whCode = dbo.bcrecproduct.whcode
where 	dbo.BCITEM.ActiveStatus = 1 and dbo.bcitem.code = @vItemCode and dbo.bcrecproduct.WHCode = @vWHCode
ORDER BY dbo.BCITEM.Code


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE		procedure dbo.USP_NP_SearchLabelPriceList_UnitCode
@vItemCode as nvarchar(20),
@vUnitCode as nvarchar(20)
as
set	dateformat dmy

SELECT	a.Barcode, a.ItemCode,   b.Name1,isnull(b.Name2,'') as Name2,isnull(c.SHELFCODE,'') as ShelfCode,
	isnull(c.whcode,'') as WHCode,isnull(d.unitcode,'') as UnitCode,d.SalePrice1,isnull(e.PriceErect,0) as PriceErect,
	isnull(f.saleprice,0) as salePromotion
FROM	dbo.BCBarCodeMaster a
	left JOIN	dbo.BCITEM b 		ON a.ItemCode = b.Code
	left join  	dbo.bcrecproduct c 	on a.itemcode = c.productcode
	left join    	dbo.BPSPriceList d 	on a.itemcode = d.itemcode and a.unitcode = d.unitcode
	left join 	dbo.bcpriceerect e 	on a.itemcode = e.itemcode
	left join 	dbo.bpspromoprice f	on a.itemcode = f.itemcode and a.barcode = f.barcode
WHERE     (a.ActiveStatus = 1) and a.itemcode = @vItemCode and d.unitcode = @vUnitCode
GROUP BY  	a.Barcode,a.ItemCode,b.Name1,b.Name2,c.SHELFCODE,c.whcode,d.unitcode,
		d.SalePrice1,e.PriceErect,f.saleprice






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchListDocument
@vUserID as nvarchar(20),
@vTypeDoc as nvarchar(20)
as

select	*
from
(
Select	a.Docno,LastPrintDateTime,isnull(b.docno  ,'') as transferno  
from	NPMaster.dbo.NPPrintServer a
		left join dbo.bcstktransfer2 b on a.docno = b.depositno 
where	Printed = 0  and 
		LastPrintedUser = @vUserID  and 
		DoctypeID = @vTypeDoc  and b.iscancel = 0
/*union

Select	a.Docno,LastPrintDateTime,isnull(b.docno  ,'') as transferno
from	NPMaster.dbo.NPPrintlogs a
		left join dbo.bcstktransfer2 b on a.docno = b.depositno 
where	Printed = 1  and 
		LastPrintedUser = @vUserID  and 
		DoctypeID = @vTypeDoc  */
)as		result 
order	by LastPrintDateTime
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchNewDocNo
@vGroupDoc as int
as
select 	header,autonumber,docnumber  
from 	npmaster.dbo.NP_Generate_DocNo 
where 	headertype = @vGroupDoc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPOS
@vCasheirCode as nvarchar(30)
as
/*
set	dateformat dmy
select  docno,creatorcode,createdatetime,netdebtamount,salecode,createdatetime
from 	bcnp.dbo.bcarinvoice  
where   posstatus <> 0 and 
	iscancel = 0 and 
	day(createdatetime)= day(getdate())and 
	month(createdatetime)= month(getdate()) and 
	year(createdatetime)= year(getdate()) and 
	--cashiercode = '27007' and 
	machineno = '03' and 
	docno not in (select invoiceno from npmaster.dbo.TB_NP_PosPayGoodLogs)
order   by createdatetime desc
*/
set	dateformat dmy
select  docno,creatorcode,createdatetime,netdebtamount,salecode,createdatetime,cashiercode
from 	bcnp.dbo.bcarinvoice  
where   posstatus <> 0 and 
	iscancel = 0 and 
	day(createdatetime)= day(getdate())and 
	month(createdatetime)= month(getdate()) and 
	year(createdatetime)= year(getdate()) and 
	cashiercode = @vCasheirCode  and 
	docno not in (select invoiceno from npmaster.dbo.TB_NP_PosPayGoodLogs)
order   by createdatetime desc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchPayGoodsPrintReserve
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(10)
as
set	dateformat dmy
select  a.invoiceno,a.whcode,paynumber
from 	npmaster.dbo.np_paygoods a 
where 	checked = 0 and invoiceno = @vInvoiceNo and whcode = @vWHCode and 
	year(paydatetime) = year(getdate()) and 
	month(paydatetime) = month(getdate()) and day(paydatetime) = day(getdate()) 
order	by a.paydatetime desc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPickingLogs
@vSaleDocNo as nvarchar(20),
@vShelfGroup as nvarchar(2)
as
set dateformat dmy

select	count(pickingno) as vCount,isnull(pickingno,'') as pickingno
from	npmaster.dbo.np_pickingslip_logs
where 	saleorderno = @vSaleDocNo and shelfgroup = @vShelfGroup
group	by pickingno

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchPickingReq
@vDocno as nvarchar(20)
as	
set	dateformat dmy
select 	distinct docno,docdate
from  	npmaster.dbo.TB_CK_MobileDocument 
where 	year(docdate) = year(getdate()) and 
	month(docdate) = month(getdate()) and 
	day(docdate) = day(getdate()) and 
	docno = @vDocno
order	by docno 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPickingRequest
@vSaleDocNo as nvarchar(20),
@vShelfGroup as nvarchar(2),
@vDocDate as nvarchar(15)
as
set dateformat dmy

select	docno,saleorderno,shelfgroup,saleorderno,count(docno) as countpicking
from	npmaster.dbo.TB_NP_QueueManagement
where 	docno = @vSaleDocNo and shelfgroup = @vShelfGroup and docdate = @vDocDate
group 	by docno,saleorderno,shelfgroup,saleorderno


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPrintPicking
@vQueNo as int,
@vQueDocDate as nvarchar(20)
as

set		dateformat dmy

select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'')as refno,isnull(a.memberid,'') as memberid,a.sourceid,
		a.isconditionsend,a.quezone,a.quedate,a.questatus,a.quereqtime,a.quetime,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.unitcode,
		isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(f.qty/isnull(g.rate1,0),0) as stkqty,
		isnull(g.rate1,0) as rate1,isnull(g.rate2,0) as rate2,isnull(h.shelfcode,'-') as shelfid
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid = b.queid and a.quedocdate = b.quedocdate 
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
		left join dbo.bcsale e on a.salecode = e.code
		left join dbo.bcstkpacking g on b.itemcode = g.itemcode and b.unitcode = g.unitcode
		left join dbo.bcstklocation f on b.itemcode = f.itemcode and b.whcode = f.whcode and b.shelfcode = f.shelfcode
		left join dbo.bcrecproduct2 h on b.itemcode = h.productcode and b.whcode = h.whcode and b.shelfcode = h.fiscalshelf
where	a.queid = @vQueNo and a.quedocdate = @vQueDocDate 
order	by b.linenumber



select * from bcrecproduct2


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE 	procedure dbo.USP_NP_SearchPrintQTY
@vUserID as nvarchar(20)
as
set dateformat dmy
Select 	isnull(ItemCode,'') AS ItemCode, isnull(barcode,'') AS barcode, isnull(Name1,'')  AS Name1,
	isnull(Name2,'') AS Name2, isnull(QTY,0) AS QTY,
	isnull(PriceLevel,0) AS PriceLevel,isnull(Price,0)  AS Price, isnull(UnitCode,'') AS UnitCode, 
	isnull(UsedUser,'') AS UsedUser,isnull(Category_ID,'') AS Category_ID,isnull(WHCode,'') AS WHCode,
	isnull(ShelfCode,'') AS ShelfCode,isnull(VENDR_ID,'') AS VENDR_ID,isnull(remark,'') AS remark,isnull(SPrice,0) AS SPrice, 
	isnull(SOPNUM,'') AS SOPNUM,isnull(SOPQUAN,'') AS SOPQUAN,isnull(SOPDOC,'') AS SOPDOC,isnull(SOPQUAD,'') AS SOPQUAD, 
	isnull(SOPREQS,'') AS SOPREQS,isnull(SOPSALE,'') AS SOPSALE, 
	isnull(SOPCUST,'') AS SOPCUST,isnull(SOPSHPM,'') AS SOPSHPM,isnull(ONHAND,'') AS ONHAND,isnull(QTYALLOCATE,'') AS QTYALLOCATE,
	isnull(RemainOutQTY,'') AS RemainOutQTY,isnull(RemainInQTY,'') AS RemainInQTY, 
	isnull(ID,0) AS ID, isnull(Type,0) AS Type 
From 	dbo.NP_Label_Temp Where UsedUser = @vUserID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPulsePicking
@vDocNo	as nvarchar(20)
as
set	dateformat dmy
select 	 PickingNo,isnull(Picker,'') as Picker,Printed, ZoneLoc,WHCode,PickingType, PickingDate, isnull(MyDescription,'')  AS MyDescription
from 	npmaster.dbo.TB_IV_PulseOfPicking 
where 	pickingno = @vDocNo and 
	year(pickingdate) = year(getdate()) and 
	month(pickingdate) = month(getdate()) and 
	day(pickingdate) = day(getdate()) --and picker is not null and finish is not null
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPulsePickingExist
@vDocNo	as nvarchar(20)
as
set	dateformat dmy
select 	 PickingNo,Printed, finish,ZoneLoc,WHCode,PickingType, PickingDate, isnull(MyDescription,'')  AS MyDescription
from 	npmaster.dbo.TB_IV_PulseOfPicking 
where 	pickingno = @vDocNo and 
	year(pickingdate) = year(getdate()) and 
	month(pickingdate) = month(getdate()) and 
	day(pickingdate) = day(getdate()) and picker is not null and finish is not null
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchPulsePickingLogs
@vDocNo	as nvarchar(20)
as
set	dateformat dmy
select 	* 
from 	npmaster.dbo.TB_IV_PulseOfPicking 
where 	pickingno = @vDocNo and 
	year(printed) = year(getdate()) and 
	month(printed) = month(getdate()) and 
	day(printed) = day(getdate())
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueCenterBegin
@vZoneID as int
as

set 	dateformat dmy
set	language us_english

if 	@vZoneID = 1
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus = 0 and a.QueZone = 'A' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 2
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus = 0 and a.QueZone = 'B' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 3
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus = 0 and a.QueZone = 'C' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 4
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus = 0 and a.QueZone = 'X' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueCenterDetails
@vQueNo as int,
@vQueDocDate as nvarchar(20)
as

set 	dateformat dmy

--if @vSourceID = 1 
--begin
select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'')as refno,isnull(a.memberid,'') as memberid,a.sourceid,isnull(quepicker,'') as picker,
	a.isconditionsend,a.quezone,a.quedate,a.questatus,a.quereqtime,a.quetime,
	b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.unitcode,
	isnull(d.name1,'') as arname,isnull(e.name,'') as salename,isnull(c.stockqty,0) as stockqty,isnull(c.remainoutqty,0) as remainoutqty,(isnull(c.stockqty,0)-isnull(c.remainoutqty,0)) as remainsale
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
	left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid = b.queid and a.quedocdate = b.quedocdate 
	left join dbo.bcitem c on b.itemcode = c.code
	left join dbo.bcar d on a.arcode = d.code
	left join dbo.bcsale e on a.salecode = e.code
where	a.queid = @vQueNo and a.quedocdate = @vQueDocDate 
order	by linenumber
--end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueCenterFinish
@vZoneID as int
as

set 	dateformat dmy
set	language us_english

if 	@vZoneID = 1
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename,isnull(quepicker,'') as quepicker,quepickstatus
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus in (2,3) and quereceived = 0 and a.QueZone = 'A' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 2
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename,isnull(quepicker,'') as quepicker,quepickstatus
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus  in (2,3) and quereceived = 0 and a.QueZone = 'B' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 3
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename,isnull(quepicker,'') as quepicker,quepickstatus
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus  in (2,3) and quereceived = 0 and a.QueZone = 'C' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 4
begin
select 	QueID,QueDocDate,DocNo,DocDate,ARCode,a.SaleCode,isnull(RefNo,'') as refno,SourceID,QueZone,QueDate,IsConditionSend,
		isnull(QueReqTime,'') as quereqtime,QueStatus,QueTime,isnull(b.name1,'') as arname,isnull(c.name,'') as salename,isnull(quepicker,'') as quepicker,quepickstatus
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.salecode = c.code
where 	questatus  in (2,3) and quereceived = 0 and a.QueZone = 'X' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueCenterPicking
@vZoneID as int
as

set 	dateformat dmy
SET LOCK_TIMEOUT 10000

set	language us_english

if 	@vZoneID = 1
begin
select 	QueID,QueDocDate,Docno,Docdate,SourceID,isnull(ARCode,'') as ARCode,QueStatus,QueDate,QueStart,QueStop,
		isnull(QuePicker,'') as QuePicker,QueReceived,QueStatus,isnull(a.SaleCode,'') as SaleCode,
		isnull(RefNo,'') as RefNo,QueZone,QueTime,QueDescription,isnull(QueReqTime,'00:00') as QueReqTime,
		isnull(b.name1,'') as arname,isnull(c.name,'') as salename,
		case
		when len(cast(DATEPART(hour,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStart)as varchar(2))
		end as StartTime,
		case Questatus 
		when 1 then convert(nvarchar(30),(getdate()-QueStart),8 )
		end as PickingTime
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.SaleCode = c.code
where 	questatus = 1 and a.QueZone = 'A' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 2
begin
select 	QueID,QueDocDate,Docno,Docdate,SourceID,isnull(ARCode,'') as ARCode,QueStatus,QueDate,QueStart,QueStop,
		isnull(QuePicker,'') as QuePicker,QueReceived,QueStatus,isnull(a.SaleCode,'') as SaleCode,
		isnull(RefNo,'') as RefNo,QueZone,QueTime,QueDescription,isnull(QueReqTime,'00:00') as QueReqTime,
		isnull(b.name1,'') as arname,isnull(c.name,'') as salename,
		case
		when len(cast(DATEPART(hour,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStart)as varchar(2))
		end as StartTime,
		case Questatus 
		when 1 then convert(nvarchar(30),(getdate()-QueStart),8 )
		end as PickingTime
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.SaleCode = c.code
where 	questatus = 1 and a.QueZone = 'B' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 3
begin
select 	QueID,QueDocDate,Docno,Docdate,SourceID,isnull(ARCode,'') as ARCode,QueStatus,QueDate,QueStart,QueStop,
		isnull(QuePicker,'') as QuePicker,QueReceived,QueStatus,isnull(a.SaleCode,'') as SaleCode,
		isnull(RefNo,'') as RefNo,QueZone,QueTime,QueDescription,isnull(QueReqTime,'00:00') as QueReqTime,
		isnull(b.name1,'') as arname,isnull(c.name,'') as salename,
		case
		when len(cast(DATEPART(hour,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStart)as varchar(2))
		end as StartTime,
		case Questatus 
		when 1 then convert(nvarchar(30),(getdate()-QueStart),8 )
		end as PickingTime
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.SaleCode = c.code
where 	questatus = 1 and a.QueZone = 'C' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end

if 	@vZoneID = 4
begin
select 	QueID,QueDocDate,Docno,Docdate,SourceID,isnull(ARCode,'') as ARCode,QueStatus,QueDate,QueStart,QueStop,
		isnull(QuePicker,'') as QuePicker,QueReceived,QueStatus,isnull(a.SaleCode,'') as SaleCode,
		isnull(RefNo,'') as RefNo,QueZone,QueTime,QueDescription,isnull(QueReqTime,'00:00') as QueReqTime,
		isnull(b.name1,'') as arname,isnull(c.name,'') as salename,
		case
		when len(cast(DATEPART(hour,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStart)as varchar(2))
		end as StartTime,
		case Questatus 
		when 1 then convert(nvarchar(30),(getdate()-QueStart),8 )
		end as PickingTime
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale c on a.SaleCode = c.code
where 	questatus = 1 and a.QueZone = 'X' and QueDocDate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
order	by quedate desc,docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_SearchQueCheckOut
@vType as int,
@vSearch as nvarchar(50)
as

set	dateformat dmy


if	@vType = 1
begin
	if	@vSearch = ''
	begin
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1 --a.isconfirm = 0 and 
	order	by a.queid 
	end
	
	if	@vSearch <> ''
	begin
	select	*
	from
	(
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.arcode like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and 
	union		
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.refno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and 
	union		
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.queid like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and
	union		
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.docno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and
	union		
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.memberid,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and
	union		
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.barcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_PickingRequestSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.salecode,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 1--a.isconfirm = 0 and
	) as	result 
	order	by queid 
	end
end

if	@vType =2
begin
	if	@vSearch = ''
	begin
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 2 --a.isconfirm = 0 and 
	order	by a.queid 
	end
	
	if	@vSearch <> ''
	begin
	select	*
	from
	(
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.arcode like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		 invqty =0  and sourceid = 2 --a.isconfirm = 0 and
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.refno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		 invqty =0  and sourceid = 2 --a.isconfirm = 0 and
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(d.docno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 2 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.memberid,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 2 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.queid like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 2 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_QueueRequestPicking d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.salecode,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 2 --a.isconfirm = 0 and 
	) as aa
	order	by queid 
	end
end

if	@vType =3
begin
	if	@vSearch = ''
	begin
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 3 --a.isconfirm = 0 and 
	order	by a.queid 
	end
	
	if	@vSearch <> ''
	begin
	select	*
	from
	(
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.arcode like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		 invqty =0  and sourceid = 3 --a.isconfirm = 0 and
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.refno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		 invqty =0  and sourceid = 3 --a.isconfirm = 0 and
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(d.docno,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 3 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.memberid,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 3 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	a.queid like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 3 --a.isconfirm = 0 and 
	union
	select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.price,0) as price,isnull(d.discountamount,0) as discountamount,
		isnull(b.pickqty,0)*(isnull(d.price,0)-isnull(d.discountamount,0)) as netamount,isnull(d.itemcode,'') as barcode
	from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_DriveInSlipSub d on a.docno = d.docno and  b.itemcode = d.itemcode
	where	isnull(a.salecode,'') like '%'+@vSearch +'%' and a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and
		invqty =0  and sourceid = 3 --a.isconfirm = 0 and 
	) as aa
	order	by queid 
	end
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQuePayItem
@vType as int,
@vZoneID as nvarchar(2),
@vSearch as nvarchar(50)
as

set		dateformat dmy

if	@vType = 0
begin

select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(a.memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.name1,'') as arname
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and questatus = 2 and
		invqty =0  and checkqty =0 and b.zoneid = @vZoneID and a.queid = @vSearch
order	by b.linenumber
end



if	@vType = 1
begin

select	*
from
(
select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(a.memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,b.linenumber,
		isnull(d.name1,'') as arname
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and questatus = 2 and 
		invqty =0  and checkqty =0 and b.zoneid = @vZoneID and a.arcode like '%'+@vSearch+'%'

union

select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(a.memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,b.linenumber,
		isnull(d.name1,'') as arname
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and questatus = 2 and
		invqty =0  and checkqty =0 and b.zoneid = @vZoneID and a.refno like '%'+@vSearch+'%'

union

select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(a.memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,b.linenumber,
		isnull(d.name1,'') as arname
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and questatus = 2 and
		invqty =0  and checkqty =0 and b.zoneid = @vZoneID and isnull(a.memberid,'') like '%'+@vSearch+'%'
/*
union

select	a.queid,a.quedocdate,a.docno,a.docdate,a.arcode,a.salecode,isnull(a.refno,'') as refno,isnull(a.memberid,'') as memberid,
		sourceid,isnull(quedescription,'') as quedescription,isnull(quepicker,'') as quepicker,
		a.questart,a.questop,questatus,isnull(quepickstatus,'') as quepickstatus,quezone,
		b.itemcode,isnull(c.name1,'') as itemname,b.whcode,b.shelfcode,b.shelfid,b.zoneid,b.qty,b.pickqty,b.checkqty,b.invqty,b.unitcode,
		isnull(d.name1,'') as arname
from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcar d on a.arcode = d.code
where	a.quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)and questatus = 2 and
		invqty =0  and checkqty =0 and b.zoneid = @vZoneID and a.queid like '%'+@vSearch+'%'*/
)		as result
order	by queid,linenumber
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueDetails
@vDocno as nvarchar(20)
as
set	dateformat dmy

if 	@vDocno = ''
begin
select 	a.Docno,a.Docdate,a.DocType,a.Status,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartTime,isnull(a.StopDateTime,'') as StopTime,isnull(a.Picker,'') as Picker,a.PickingStatus,a.IsReceived, 
	a.SaleMan,isnull(c.name,'MobileApp') as salename,a.SaleOrderNo,a.WHCode,a.ARCode+'/'+b.name1 as arname,
	case
	when len(cast(DATEPART(hour,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.QueueDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.QueueDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.QueueDateTime)as varchar(2))
	end as PrintTime ,
	case
	when len(cast(DATEPART(hour,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.StartDateTime)as varchar(2))
	end as BeginTime ,
	case
	when len(cast(DATEPART(hour,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.StopDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.StopDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.StopDateTime)as varchar(2))
	end as FinishTime,
	case
	when status = 0 then
	'00:00:00'
	when status = 1 then
	convert(nvarchar(30),(getdate()-startdatetime),8 ) 
	when status = 2 then
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
	end as PickingTime, 
	case 
	when a. status  = 0 and a.IsReceived = 0 then 'รอจัดของ'
	when a. status  = 1 and a.IsReceived = 0 then 'กำลังจัดของ'
	when a. status  = 2 and a.IsReceived = 0 then 'รอจ่ายของ' 
	when a. status  = 2 and a.IsReceived = 1 then 'รับของแล้ว'
	end as StatusDescription
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on isnull(a.saleman,'') = c.code
order	by cast(docno as int)
end

if 	@vDocno <> ''
begin
select 	a.Docno,a.Docdate,a.DocType,a.Status,a.QueueDateTime, 
	isnull(a.StartDateTime,'') as StartTime,isnull(a.StopDateTime,'') as StopTime,isnull(a.Picker,'') as Picker,a.PickingStatus,a.IsReceived, 
	a.SaleMan,isnull(c.name,'MobileApp') as salename,a.SaleOrderNo,a.WHCode,a.ARCode+'/'+b.name1 as arname,
	case
	when len(cast(DATEPART(hour,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.QueueDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.QueueDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.QueueDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.QueueDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.QueueDateTime)as varchar(2))
	end as PrintTime ,
	case
	when len(cast(DATEPART(hour,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.StartDateTime)as varchar(2))
	end as BeginTime ,
	case
	when len(cast(DATEPART(hour,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.StopDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.StopDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.StopDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.StopDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.StopDateTime)as varchar(2))
	end as FinishTime,
	case
	when status = 0 then
	'00:00:00'
	when status = 1 then
	convert(nvarchar(30),(getdate()-startdatetime),8 ) 
	when status = 2 then
	convert(nvarchar(30),(stopdatetime-startdatetime),8 ) 
	end as PickingTime,
	case  
	when a. status  = 0 and a.IsReceived = 0 then 'รอจัดของ'
	when a. status  = 1 and a.IsReceived = 0 then 'กำลังจัดของ'
	when a. status  = 2 and a.IsReceived = 0 then 'รอจ่ายของ' 
	when a. status  = 2 and a.IsReceived = 1 then 'รับของแล้ว'
	end as StatusDescription
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on isnull(a.saleman,'') = c.code
where 	a.docno = @vDocno
order	by cast(docno as int)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueDoc
as
select 	a.*,b.name1 as arname,isnull(c.name,a.saleman) as salename 
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
where 	status = 0 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and zoneid in ('02','03')
order	by queuedatetime desc,docno desc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueFinish
as
select 	top 10 a.docno,a.docdate,a.picker,b.name1 as arname,a.timeid,picker,saleorderno,isnull(c.name,a.saleman) as salename,doctype
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
where 	status = 2 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate())  and
	isreceived = 0 and zoneid in ('02','03')
order	by stopdatetime desc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueItemDetails
@vPickingNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vDocType as int,
@vTimeID as int 
as
set 	dateformat dmy

if @vDocType = 1 
begin
select 	a.*,b.saleorderno,b.whcode as whcodeline,b.shelfgroup,isnull(itemcode,'') as itemcode,isnull(itemname,'') as itemname,isnull(qty,0) as qty,isnull(unitcode,'') as unitcode,
	isnull(c.whcode,b.whcode) as whcode,isnull(c.shelfgroup,'') as shelfgroup,isnull(c.shelfcode,'') as shelfcode
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join npmaster.dbo.NP_PickingSlip_Logs b on a.refdocno = b.pickingno
	left join 
		(
		select  distinct 
			top 100 percent a.DocNo,
			left(case 
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode = 'BAK' then 'K'
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode <> 'BAK' then 'M'
			when b.whcode = '014' and y.itemtype is  null  then 'H'
			when b.whcode = '020'  then 'E'
			--when b.whcode = '015'   then 'C'
			when b.whcode = '016'  then 'Y'
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
			end,1) as shelfgroup,
			isnull(itemcode,'') as itemcode,isnull(b.itemname,'') as itemname,isnull(qty,0) as qty,
			isnull(unitcode,'') as unitcode,b.whcode,isnull(z.shelfcode,'') as shelfcode
		
		FROM	dbo.BCSaleOrder a  
			LEFT OUTER JOIN dbo.BCSaleOrderSub b ON a.DocNo = b.DocNo 
			left join bcnp.dbo.bcitem g on b.itemcode = g.code
			left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa )z 
			on b.itemcode = z.productcode  and b.whcode = z.whcode
			left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype	
		WHERE    a.iscancel = 0 and a.docno = @vSaleOrderNo and b.whcode not in ('011','099','097','070','015')
		)as c on b.saleorderno = c.docno and b.whcode = c.whcode and b.shelfgroup = c.shelfgroup
	
where 	a.doctype = 1 and a.docno = @vPickingNo and a.timeid = @vTimeID and a.docdate = cast(rtrim(day(getdate()))+'/'+ rtrim(month(getdate()))+'/'+rtrim(year(getdate()))as datetime)
order	by queuedatetime
end
if @vDocType = 2 
begin
select 	a.*,isnull(b.Docno,'') as docnolink,isnull(b.ItemCode,'') as itemcode,isnull(b.ItemName,'') as itemname,isnull(b.QTY,0) as qty,
	isnull(b.UnitCode,'') as unitcode,isnull(replace(b.RefNo,' ',''),'') as refno ,isnull(c.shelfcode,'') as shelfcode
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join npmaster.dbo.TB_CK_MobileDocument b on a.refdocno = b.docno and replace(a.saleorderno,' ','')  = replace(b.refno,' ','') and 
	year(b.docdate) = year(getdate()) and month(b.docdate) = month(getdate()) and day(b.docdate) = day(getdate())
	left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa ) c
	on b.itemcode = c.productcode and c.whcode = '014'
where 	a.doctype = 2 and a.docno = @vPickingNo  and a.timeid = @vTimeID and 
	year(a.docdate) = year(getdate()) and month(a.docdate) = month(getdate()) and day(a.docdate) = day(getdate())
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueItemDetails1
@vQueueNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vDocType as int,
@vTimeID as int 
as
set 	dateformat dmy

if @vDocType = 1 
begin

select 	a.*,isnull(c.itemcode,'') as itemcode,isnull(c.itemname,'') as itemname,isnull(c.qty,0) as qty,isnull(c.unitcode,'') as unitcode,
	isnull(c.whcode,a.whcode) as whcode,isnull(c.shelfgroup,'') as shelfgroup,isnull(c.shelfcode,'') as shelfcode,isnull(b.pickqty,0) as pickqty,isnull(b.qty,0) as reqqty
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join 
		(
		select  distinct 
			top 100 percent a.DocNo,
			left(case 
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode = 'BAK' then 'K'
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode <> 'BAK' then 'M'
			when b.whcode = '014' and y.itemtype is  null  then 'H'
			when b.whcode = '020'  then 'E'
			--when b.whcode = '015'   then 'C'
			when b.whcode = '016'  then 'Y'
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
			end,1) as shelfgroup,
			isnull(itemcode,'') as itemcode,isnull(b.itemname,'') as itemname,isnull(qty,0) as qty,
			isnull(unitcode,'') as unitcode,b.whcode,isnull(z.shelfcode,'') as shelfcode
		
		FROM	dbo.BCSaleOrder a  
			LEFT OUTER JOIN dbo.BCSaleOrderSub b ON a.DocNo = b.DocNo 
			left join bcnp.dbo.bcitem g on b.itemcode = g.code
			left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa )z 
			on b.itemcode = z.productcode  and b.whcode = z.whcode
			left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype	
		WHERE    a.iscancel = 0 and a.docno = @vSaleOrderNo and b.whcode not in ('011','099','097','070','015')
		)as c on a.saleorderno = c.docno and a.whcode = c.whcode and a.shelfgroup = c.shelfgroup
	left join	npmaster.dbo.TB_NP_QueueManagementSub b on a.docno = b.pickingno and a.docdate = b.docdate and c.itemcode = b.itemcode and c.unitcode = b.unitcode
	
where 	a.doctype = 1 and a.docno = @vQueueNo and a.timeid = @vTimeID and a.docdate = cast(rtrim(day(getdate()))+'/'+ rtrim(month(getdate()))+'/'+rtrim(year(getdate()))as datetime)
order	by queuedatetime
end
if @vDocType = 2 
begin
select 	a.*,isnull(b.Docno,'') as docnolink,isnull(b.ItemCode,'') as itemcode,isnull(b.ItemName,'') as itemname,isnull(b.QTY,0) as qty,
	isnull(b.UnitCode,'') as unitcode,isnull(replace(b.RefNo,' ',''),'') as refno ,isnull(c.shelfcode,'') as shelfcode,isnull(d.pickqty,0) as pickqty,isnull(d.qty,0) as reqqty
from 	npmaster.dbo.TB_NP_QueueManagement a
	left join npmaster.dbo.TB_CK_MobileDocument b on a.refdocno = b.docno and replace(a.saleorderno,' ','')  = replace(b.refno,' ','') and 
	year(b.docdate) = year(getdate()) and month(b.docdate) = month(getdate()) and day(b.docdate) = day(getdate())
	left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa ) c
	on b.itemcode = c.productcode and c.whcode = '014'
	left join	npmaster.dbo.TB_NP_QueueManagementSub d on a.docno = d.pickingno and a.docdate = d.docdate and b.itemcode = d.itemcode and b.unitcode = d.unitcode
where 	a.doctype = 2 and a.docno = @vQueueNo  and a.timeid = @vTimeID and 
	year(a.docdate) = year(getdate()) and month(a.docdate) = month(getdate()) and day(a.docdate) = day(getdate())
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueItemDetails2
@vQueueNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vDocType as int,
@vTimeID as int,
@vDocDate as nvarchar(20) 
as
set 	dateformat dmy

if @vDocType = 1 
begin

select 	a.*,isnull(c.itemcode,'') as itemcode,isnull(c.itemname,'') as itemname,isnull(c.qty,0) as qty,
		isnull(c.unitcode,'') as unitcode,
		isnull(c.whcode,a.whcode) as whcode,isnull(c.shelfgroup,'') as shelfgroup,isnull(c.shelfcode,'') as shelfcode,
		isnull(b.pickqty,0) as pickqty,isnull(b.qty,0) as reqqty
from 	npmaster.dbo.TB_NP_QueueManagement a
		left join 
		(
			select  distinct 
			top		100 percent a.DocNo,
			left(case 
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode = 'BAK' then 'K'
			when b.whcode = '014' and y.itemtype is not null and b.shelfcode <> 'BAK' then 'M'
			when b.whcode = '014' and y.itemtype is  null  then 'H'
			when b.whcode = '020'  then 'E'
			when b.whcode = '016'  then 'Y'
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
			when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
			end,1) as shelfgroup,
			isnull(itemcode,'') as itemcode,isnull(b.itemname,'') as itemname,isnull(qty,0) as qty,
			isnull(unitcode,'') as unitcode,b.whcode,isnull(z.shelfcode,'') as shelfcode
			from	dbo.BCSaleOrder a  
				left join dbo.BCSaleOrderSub b ON a.DocNo = b.DocNo 
				left join bcnp.dbo.bcitem g on b.itemcode = g.code
				left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa )z 
				on b.itemcode = z.productcode  and b.whcode = z.whcode
				left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype	
				where    a.iscancel = 0 and a.docno = @vSaleOrderNo and b.whcode not in ('011','099','097','070','015')
				)as c on a.saleorderno = c.docno and a.whcode = c.whcode and a.shelfgroup = c.shelfgroup
				left join	npmaster.dbo.TB_NP_QueueManagementSub b on a.docno = b.pickingno and a.docdate = b.docdate and c.itemcode = b.itemcode and c.unitcode = b.unitcode	
where 	a.doctype = 1 and a.docno = @vQueueNo and a.timeid = @vTimeID and a.docdate = @vDocDate
order	by queuedatetime
end
if @vDocType = 2 
begin
select 	a.*,isnull(b.Docno,'') as docnolink,isnull(b.ItemCode,'') as itemcode,isnull(b.ItemName,'') as itemname,
		isnull(b.QTY,0) as qty,
		isnull(b.UnitCode,'') as unitcode,isnull(replace(b.RefNo,' ',''),'') as refno ,isnull(c.shelfcode,'') as shelfcode,
		isnull(d.pickqty,0) as pickqty,isnull(d.qty,0) as reqqty
from 	npmaster.dbo.TB_NP_QueueManagement a
		left join npmaster.dbo.TB_CK_MobileDocument b on a.refdocno = b.docno and replace(a.saleorderno,' ','')  = replace(b.refno,' ','') 
		left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa ) c
		on b.itemcode = c.productcode and c.whcode = '014'
		left join	npmaster.dbo.TB_NP_QueueManagementSub d on a.docno = d.pickingno and a.docdate = d.docdate and b.itemcode = d.itemcode and b.unitcode = d.unitcode
where 	a.doctype = 2 and a.docno = @vQueueNo  and a.timeid = @vTimeID and a.docdate = @vDocDate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueItemDetails3
@vQueueNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vDocType as int,
@vTimeID as int,
@vDocDate as nvarchar(20) 
as

set 	dateformat dmy

--Test

if @vDocType = 1 
begin

select	a.*,
		isnull(c.itemcode,'') as itemcode,isnull(c.itemname,'') as itemname,isnull(c.reqqty,0) as qty,
		isnull(c.unitcode,'') as unitcode,
		isnull(c.whcode,a.whcode) as whcode,isnull(c.zoneid,'') as shelfgroup,isnull(c.shelfcode,'') as shelfcode,
		isnull(d.pickqty,0) as pickqty,isnull(d.unitcode,'') as pickunitcode
from	npmaster.dbo.TB_NP_QueueManagement a
		left join npmaster.dbo.TB_NP_QueueRequestPickingMaster b on a.saleorderno = b.docno and a.docno = b.queueno and a.docdate = datepicking and a.timeid = b.socountnumber and a.shelfgroup = b.shelfgroup
		left join npmaster.dbo.TB_NP_QueueRequestPicking c on b.DocNo = c.DocNo and b.docdate = c.docdate and b.socountnumber = c.socountnumber and b.shelfgroup = c.zoneid
		left join npmaster.dbo.TB_NP_QueueManagementSub d on a.docno = d.pickingno and a.docdate = d.docdate and c.itemcode = d.itemcode and c.unitcode = d.unitcode
where	a.docno = @vQueueNo and a.docdate = @vDocDate and a.timeid = @vTimeID and a.saleorderno = @vSaleOrderNo
order	by c.linenumber
end

if @vDocType = 2 
begin
select 	a.*,isnull(b.Docno,'') as docnolink,isnull(b.ItemCode,'') as itemcode,isnull(b.ItemName,'') as itemname,
		isnull(b.QTY,0) as qty,
		isnull(b.UnitCode,'') as unitcode,isnull(replace(b.RefNo,' ',''),'') as refno ,isnull(c.shelfcode,'') as shelfcode,
		isnull(d.pickqty,0) as pickqty,isnull(d.qty,0) as reqqty
from 	npmaster.dbo.TB_NP_QueueManagement a
		left join npmaster.dbo.TB_CK_MobileDocument b on a.refdocno = b.docno and replace(a.saleorderno,' ','')  = replace(b.refno,' ','') 
		left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct2) as aaa ) c
		on b.itemcode = c.productcode
		left join	npmaster.dbo.TB_NP_QueueManagementSub d on a.docno = d.pickingno and a.docdate = d.docdate and b.itemcode = d.itemcode and b.unitcode = d.unitcode
where 	a.doctype = 2 and a.docno = @vQueueNo  and a.timeid = @vTimeID and a.docdate = @vDocDate
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueItemPickingDetails
@QueueNo as nvarchar(20)
as

declare	@vCheckItemCount as int
declare	@vCheckItemCountPickingRequest as int
declare	@vSaleOrderNo as nvarchar(20)

set	dateformat dmy
set	@vSaleOrderNo = (select isnull(saleorderno,'') as saleorder from npmaster.dbo.TB_NP_QueueManagement where docno = @QueueNo and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()))
set	@vCheckItemCount = (select isnull(count(itemcode),0) as vCount from npmaster.dbo.TB_NP_QueueManagementsub where pickingno = @QueueNo and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()))
set	@vCheckItemCountPickingRequest = (select isnull(count(itemcode),0) as vCount from npmaster.dbo.TB_CK_MobileDocument where refno = @vSaleOrderNo)

if 	@vCheckItemCount > 0  
begin
select 	a.docno,isreceived,pickingstatus,status,statusdesc,isnull(picker,'') as picker,
	b.itemcode,b.itemname,b.qty,b.pickqty,b.unitcode,
	a.arcode+'//'+isnull(c.name1,'') as arname,isnull(a.saleman,'')+'//'+isnull(d.name,a.saleman) as salename 
from 	npmaster.dbo.TB_NP_QueueManagement a
	inner join npmaster.dbo.TB_NP_QueueManagementsub b on a.docno = b.pickingno and a.docdate = b.docdate
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
where 	year(a.docdate) = year(getdate()) and month(a.docdate) = month(getdate())  and day(a.docdate) = day(getdate()) and
	docno = @QueueNo
order	by b.linenumber
end

if 	@vCheckItemCount = 0 and @vCheckItemCountPickingRequest = 0
begin
select	a.docno,isreceived,pickingstatus,status,statusdesc,isnull(picker,'') as picker,
	b.*,0 as pickqty,
	a.arcode+'//'+isnull(c.name1,'') as arname,isnull(a.saleman,'')+'//'+isnull(d.name,a.saleman) as salename
from 	npmaster.dbo.TB_NP_QueueManagement a inner join
	(
	select 	docno,docdate,arcode,itemcode,itemname,sum(qty) as qty,unitcode 
	from 	dbo.bcsaleordersub where docno = @vSaleOrderNo
	group 	by docno,docdate,arcode,itemcode,itemname,unitcode
	) as b on a.saleorderno = b.docno and a.arcode = b.arcode
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
end

if 	@vCheckItemCountPickingRequest > 0
begin
select	a.docno,isreceived,pickingstatus,status,statusdesc,isnull(picker,'') as picker,
	b.*,0 as pickqty,
	a.arcode+'//'+isnull(c.name1,'') as arname,isnull(a.saleman,'')+'//'+isnull(d.name,a.saleman) as salename
from 	npmaster.dbo.TB_NP_QueueManagement a inner join
	(
	select 	refno,docdate,'1' as arcode,itemcode,itemname,sum(qty) as qty,unitcode 
	from 	npmaster.dbo.TB_CK_MobileDocument where refno = @vSaleOrderNo
	group 	by refno,docdate,itemcode,itemname,unitcode
	) as b on a.saleorderno = b.refno and a.arcode = b.arcode
	left join dbo.bcar c on a.arcode = c.code
	left join dbo.bcsale d on a.saleman = d.code
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE	procedure dbo.USP_NP_SearchQueueLine
as
select 	a.*,isnull(b.name1,'')  as arname,isnull(c.name,a.saleman) as salename,
	case
	when len(cast(DATEPART(hour,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(hour,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(minute,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(minute,a.StartDateTime)as varchar(2))
	end+':'+
	case
	when len(cast(DATEPART(second,a.StartDateTime)as varchar(2))) <=1 then
	'0'+cast(DATEPART(second,a.StartDateTime)as varchar(2))
	else 
	cast(DATEPART(second,a.StartDateTime)as varchar(2))
	end as StartTime
from 	(select * from npmaster.dbo.TB_NP_QueueManagement where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate())  and zoneid in ('02','03')) a
	left join dbo.bcar b on a.arcode = b.code
	left join dbo.bcsale c on a.saleman = c.code
--where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate())  and zoneid in ('02','03')
order	by queuedatetime

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueueLogs
@vSaleDocNo as nvarchar(20),
@vShelfGroup as nvarchar(2)
as
set dateformat dmy

select	docno,saleorderno,shelfgroup,count(docno) as countpicking
from	npmaster.dbo.TB_NP_QueueManagement
where 	saleorderno = @vSaleDocNo and shelfgroup = @vShelfGroup
group 	by docno,saleorderno,shelfgroup
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQueuePrint
as
set	dateformat dmy
select 	top 1 jobid,docno,a.reportid,a.reporttype,printstatus,userprint,dateprint,
	b.reportname,case printstatus when 0 then 'ยังไม่ได้พิมพ์' else 'พิมพ์แล้ว' end as PrintText
from 	npmaster.dbo.TB_NP_CheckQueuePrint a left join 
	dbo.bcreportname b on a.reportid = b.repid and a.reporttype = b.reptype
where 	printstatus = 0 and jobid = '02'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchQuotation
@vTypeSearch as int,
@vDocSearch as nvarchar(30)
as
set dateformat dmy
if @vTypeSearch = 0
begin
	select 	* from
	(
	select 	docno,docdate,b.name1 as arname 
	from 	npmaster.dbo.tb_np_quotation a 
		inner join bcnp.dbo.bcar b on a.arcode = b.code
	where 	a.docno like @vDocSearch
	union
	select 	docno,docdate,b.name1 as arname 
	from 	npmaster.dbo.tb_np_quotation a 
		inner join bcnp.dbo.bcar b on a.arcode = b.code
	where 	b.name1 like @vDocSearch
	) as result 
	order	by docdate,docno desc
end
else
begin
	select 	docno,docdate,b.name1 as arname 
	from 	npmaster.dbo.tb_np_quotation a 
		inner join bcnp.dbo.bcar b on a.arcode = b.code
	order	by docdate,docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_SearchQuotationDetails
@vDocno as nvarchar(20)
as
set dateformat dmy
select 	a.DocNo,a.DocDate,BillType,ArCode,CreditDay,Validate,IsConditionSend, 
	a.SaleCode,TaxRate,a.IsCancel,SumOfItemAmount,TaxAmount,isnull(a.DiscountAmount,0) as DiscountAmount, 
	TotalAmount,NetAmount,isnull(MyDescription,'') as MyDescription,
	ItemCode,isnull(ItemName,'') as ItemName,QTY,Price,UnitCode,
	isnull(b.DisCountAmount,0) as DiscountAmountSub,isnull(SumDisCountAmount,0) as SumDisCountAmount,
	Amount,LineNumber,c.name1 as arname
from 	npmaster.dbo.tb_np_quotation a
	inner join npmaster.dbo.tb_np_quotationsub b on a.docno = b.docno
	inner join bcnp.dbo.bcar c on a.arcode = c.code
where 	a.docno = @vDocno
order	by linenumber

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchReqPicking
@vSearch as nvarchar(50)

as

set		dateformat dmy

if		@vSearch = ''
begin
select	top 100 a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	iscancel = 0
order	by docno desc
end

if		@vSearch <> ''
begin
select	*
from 
(		
select	a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.docno like '%'+@vSearch+'%' and iscancel = 0
union
select	a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.arcode like '%'+@vSearch+'%' and iscancel = 0
union
select	a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	isnull(a.memberid,'') like '%'+@vSearch+'%' and iscancel = 0
union
select	a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	isnull(a.refno,'') like '%'+@vSearch+'%' and iscancel = 0
union
select	a.DocNo,a.DocDate,a.ARCode,isnull(c.name1,'') as arname,a.SaleCode,isnull(d.name,'') as salename,isnull(refno,'') as refno,
		isnull(a.memberid,'') as memberid,a.NetDebtAmount		
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join dbo.bcar c on a.arcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.salecode like '%'+@vSearch+'%' and iscancel = 0
)	as	result
order	by docno
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchRequestConfirm
@vDocno as nvarchar(20)
as
declare @vExist as int
declare	@vBillStatus as int

set	@vExist = (select count(*) as vCount from dbo.bcreqconfirm where docno = @vDocno)
if 	@vExist = 0
begin
select 	'ไม่มีข้อมูลเอกสาร กรุณาตรวจสอบ' ScriptDescription
end
if 	@vExist <> 0
begin
set	@vBillStatus = (select count(*) as vCount from dbo.bcreqconfirm where docno = @vDocno and billstatus = 0)

if 	@vBillStatus = 0 
begin
select 	'เอกสารดังกล่าวได้ถูกอ้างไปทำใบสั่งซื้อเรียบร้อยแล้ว' ScriptDescription
end

if 	@vBillStatus <> 0
begin
select 	a.docno,a.docdate,b.linenumber,b.itemcode,c.name1 as itemname,qty,confirmqty,unitcode,b.stkrequestno,
	'' ScriptDescription 
from 	dbo.bcreqconfirm a
	inner join dbo.bcreqconfirmsub b on a.docno = b.docno and a.docdate = b.docdate
	inner join dbo.bcitem c on b.itemcode = c.code
where 	a.docno = @vDocno
end

end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchRequestQueueItem
@SaleOrderNo as nvarchar(20)
as

set	dateformat dmy
set	language us_english

select 	top 1 SaleOrderNo,SaleOrderDate,ARCode,SaleCode,RequestDate,RequestTime,RequestStatus,RequestCountItem,RequestCountQTY,PrintStatus,RequestAt
from	npmaster.dbo.TB_NP_PickingQueueRequest
where	SaleOrderNo = @SaleOrderNo
order	by RequestAt desc





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchSaleOrder
@vDocNo as nvarchar(20)
as

set		dateformat dmy

select	a.docno,a.docdate,a.arcode,a.salecode,a.billtype,a.billstatus,a.sostatus,a.sumofitemamount,isnull(a.discountword,'') as discountword,a.discountamount,
		a.taxamount,a.totalamount,a.netamount,a.creatorcode,a.createdatetime,a.isconditionsend,
		b.itemcode,b.itemname,b.whcode,b.shelfcode,b.qty,b.qty as remainqty,b.price,isnull(b.discountword,'') as discountwordsub,
		(b.discountamount/b.qty) as discountamountsub,b.amount,b.netamount as netamountsub,b.unitcode,b.packingrate1,b.packingrate2,b.itemtype,
		isnull(c.name1,'') as arname,isnull(d.name,'') as salename,isnull(f.zoneid,'X') as zoneid,
		isnull((select top 1 shelfcode as shelfid from dbo.bcrecproduct2 where b.itemcode = productcode and b.whcode = whcode and b.shelfcode = fiscalshelf order by roworder desc),'-') as ShelfID
from	dbo.bcsaleorder a
		left join dbo.bcsaleordersub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		left join dbo.bcar c on a.arcode = c.code 
		left join dbo.bcsale d on a.salecode = d.code
		left join dbo.bcitem e on b.itemcode = e.code
		left join npmaster.dbo.TB_NP_CategoryPickingZone f on e.categorycode = f.categorycode
where	a.docno =@vDocNo and a.iscancel = 0 --and billstatus = 0 and b.remainqty > 0
order	by b.linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchSaleOrderData
@vDocno as nvarchar(20)
as

set		dateformat dmy
set		language us_english
/*
xxxxxxxxxxxxxxxxx
*/
/*
select	a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,unitcode
from	dbo.bcsaleorder a 
		inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno = @vDocno group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		inner join dbo.bcar c on a.arcode = c.code
where	a.iscancel = 0 and a.billstatus = 0 and a.docno = @vDocno
order	by itemcode
*/

/*
select	a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,unitcode,
	isnull(left(case 
	when b.whcode = '014' and y.itemtype is not null and b.shelfcode = 'BAK' then 'K'
	when b.whcode = '014' and y.itemtype is not null and b.shelfcode <> 'BAK' then 'M'
	when b.whcode = '014' and y.itemtype is  null  then 'H'
	when b.whcode = '020' then 'H'
	when b.whcode = '016' then 'Y'
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
	end,1),'D') as shelfgroup
from	dbo.bcsaleorder a 
	inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(qty) as remainqty from dbo.bcsaleordersub where docno = @vDocno and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
	inner join dbo.bcar c on a.arcode = c.code
	left join dbo.bcitem g on b.itemcode = g.code
	left join (select distinct productcode,whcode,(select top 1 shelfcode from dbo.bcrecproduct2 where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from dbo.bcrecproduct2) as aaa )z 
	on b.itemcode = z.productcode  and b.whcode = z.whcode
	left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype
where	a.iscancel = 0 and a.billstatus <> 1 and a.docno = @vDocno 
order	by itemcode*/


select	*
from
(
select	0 as copy,0 as copy1,a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,qty as orderqty,unitcode
from	dbo.bcsaleorder a 
		inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(qty) as qty,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno = 'scv5101-0035' and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		inner join dbo.bcar c on a.arcode = c.code
		left join dbo.bcitem g on b.itemcode = g.code
where	a.billstatus <> 1 and a.docno = 'scv5101-0035' 
union
select	0 as copy,1 as copy1,a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,qty as orderqty,unitcode
from	dbo.bcsaleorder a 
		inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(qty) as qty,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno = 'scv5101-0035' and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		inner join dbo.bcar c on a.arcode = c.code
		left join dbo.bcitem g on b.itemcode = g.code
where	a.billstatus <> 1 and a.docno = 'scv5101-0035' 
union
select	1 as copy,0 as copy1,a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,qty as orderqty,unitcode
from	dbo.bcsaleorder a 
		inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(qty) as qty,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno = 'scv5101-0035' and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		inner join dbo.bcar c on a.arcode = c.code
		left join dbo.bcitem g on b.itemcode = g.code
where	a.billstatus <> 1 and a.docno = 'scv5101-0035' 
union
select	1 as copy,1 as copy1,a.docno,a.docdate,a.arcode,isnull(c.name1,'') as arname,b.whcode,b.shelfcode,b.itemcode,b.itemname,remainqty as qty,qty as orderqty,unitcode
from	dbo.bcsaleorder a 
		inner join (select docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode,sum(qty) as qty,sum(remainqty) as remainqty from dbo.bcsaleordersub where docno = 'scv5101-0035' and whcode not in ('070','080','097','099') group by  docno,docdate,arcode,whcode,shelfcode,itemcode,itemname,unitcode) b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode
		inner join dbo.bcar c on a.arcode = c.code
		left join dbo.bcitem g on b.itemcode = g.code
where	a.billstatus <> 1 and a.docno = 'scv5101-0035' 
) as 	a
order	by copy,docno,itemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchSaleOrderGroupShelf
@vDocNo as nvarchar(20)
as

set		dateformat dmy
set		language us_english

select	distinct a.docno,zoneid as shelfgroup
from	dbo.bcsaleorder a
		left join dbo.bcsaleordersub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode 
		left join dbo.bcitem c on b.itemcode = c.code
		left join npmaster.dbo.TB_NP_CategoryPickingZone d on c.categorycode = d.categorycode
where	a.docno = @vDocNo and a.iscancel = 0 --and a.billstatus = 0 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchSendQueuePicking
@vDocNo as nvarchar(30)
as

set		dateformat dmy

select	isnull(issendque ,0) as issendque
from	npmaster.dbo.TB_NP_PickingRequestMaster
where	docno = @vDocNo
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SearchStatusMinuteItem
as

select checkstock from dbo.bpsconfig 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SearchTransferFromDeposit
@vDocNo as nvarchar(20)
as


Select	isnull(b.docno ,'') as transferno
from	dbo.bcardeposit  a
		left join dbo.bcstktransfer2 b on a.docno = b.depositno 
where	a.docno = @vDocNo and b.iscancel = 0

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SelectDocnoChangePrice
@vScheduleDate as nvarchar(20),
@vItemCode as nvarchar(25),
@vPriceLevel as smallint,
@vsaleType as smallint,
@vTransSportType as smallint,
@vUnitCode as nvarchar(30)

as

set	dateformat dmy
select 	a.docno
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceSub a
	left join npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster b on a.docno = b.docno and a.docdate = b.docdate
where 	scheduledate = @vScheduleDate and 
	itemcode = @vItemCode and 
	pricelevel = @vPriceLevel and 
	saletype = @vsaleType and 
	transsporttype = @vTransSportType and 
	unitcode = @vUnitCode and 
	isupdate = 0

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_SelectItemChangePricePrintLabel
as

set	dateformat dmy
select 	a.docno,a.docdate,scheduledate,b.itemcode,itemname,b.unitcode,newprice,oldprice,isnull(isprintlabel ,0) as isprintlabel,
	c.barcode
from 	npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster a
	inner join npmaster.dbo.TB_NP_BasketItemUpdatePriceSub b on a.docno = b.docno and a.docdate = b.docdate
	left join dbo.bcbarcodemaster c on b.itemcode = c.itemcode and b.unitcode = c.unitcode
where 	isnull(isprintlabel,0)  = 0 and isconfirm = 1 and activestatus = 1 and isupdate = 1 and pricelevel = 1 and
	saletype = 0 and transsporttype = 0
order	by a.docno,b.itemcode


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_SelectItemReceivePrintLabel
@vDocno as nvarchar(20)
as
select   top 100 percent  a.docno,a.docdate,b.itemcode,b.itemname,b.whcode,b.cnqty,b.unitcode,c.barcode,d.saleprice1  
from bcapinvoice a 
left outer join bcapinvoicesub b on a.docno = b.docno and a.apcode = b.apcode and b.iscancel = 0
left outer join bcbarcodemaster c on b.itemcode = c.itemcode and b.itemcode = c.barcode
left outer join bpspricelist d on b.itemcode = d.itemcode 
where  a.docno = @vDocno
group by  a.docno,a.docdate,b.itemcode,b.itemname,b.whcode,b.cnqty,b.unitcode,c.barcode,b.linenumber,d.saleprice1
order by b.linenumber



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE 	procedure dbo.USP_NP_SelectReportName
@vRepID as int,
@vRepType as nvarchar(10)
as
set	dateformat dmy
SET LOCK_TIMEOUT 15000

select reportname from bcnp.dbo.bcreportname where repid =  @vRepID and reptype =@vRepType
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_ShowQueMonitor

as

set	dateformat dmy

select 	a.QueID,a.QueDocdate,a.SourceID,a.ARCode,a.QueStatus,a.QueDescription,a.QueDate, 
		isnull(a.QueStart,'') as QueStart1,a.QueStop,isnull(a.QuePicker,'') as QuePicker,a.QueStatus,a.QueReceived, 
		QuePickStatus,isnull(b.name1,'') as arname,
		case
		when len(cast(DATEPART(hour,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStart)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStart)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStart)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStart)as varchar(2))
		end as QueStart ,
		case
		when len(cast(DATEPART(hour,a.QueStop)as varchar(2))) <=1 then
		'0'+cast(DATEPART(hour,a.QueStop)as varchar(2))
		else 
		cast(DATEPART(hour,a.QueStop)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(minute,a.QueStop)as varchar(2))) <=1 then
		'0'+cast(DATEPART(minute,a.QueStop)as varchar(2))
		else 
		cast(DATEPART(minute,a.QueStop)as varchar(2))
		end+':'+
		case
		when len(cast(DATEPART(second,a.QueStop)as varchar(2))) <=1 then
		'0'+cast(DATEPART(second,a.QueStop)as varchar(2))
		else 
		cast(DATEPART(second,a.QueStop)as varchar(2))
		end as QueStop,
		convert(nvarchar(30),(getdate()-QueStart),8 ) as PickingTime,
		case questatus 
		when 0 then 'สินค้า'
		when 1 then ''
		when 2 then 'รอจ่ายของ' end as StatusDescription,
		Quereqtime
from 	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join dbo.bcar b on a.arcode = b.code
		left join dbo.bcsale d on a.salecode = d.code
where 	(a.quedocdate = rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) or (a.quedocdate < rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) and DATEPART (hh,getdate())<10)) and 
		quereceived = 0  
order	by quedocdate,queid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_StockRequestApprove
@vDocNo as nvarchar(20)
as

set	dateformat dmy

select	*,
case
when daysale > 0 then hmx/daysale 
when daysale = 0 then hmx/1
end as hmxavgperday,
case
when daysale > 0 then np/daysale 
when daysale = 0 then np/1
end as npavgperday,
case 
when daysale > 0 then hmx/case when (daysale/30) = 0 then 1 else (daysale/30) end
when daysale = 0 then hmx/1
end	as hmxavgper30,
case 
when daysale > 0 then np/case when (daysale/30) = 0 then 1 else (daysale/30) end
when daysale = 0 then np/1
end	as npavgper30,
case 
when daysale > 0 and hmx > 0 then onhandhmx/case when(hmx/daysale)=0 then 1 else (hmx/daysale)  end
when daysale > 0 and hmx = 0 then onhandhmx/1
when daysale = 0 and hmx > 0 then onhandhmx/case when(hmx/1)=0 then 1 else (hmx/1) end
when daysale = 0 and hmx = 0 then onhandhmx/1
end as hmxremainsale,
case 
when daysale > 0 and np > 0 then onhandnp/case when (np/daysale)=0 then 1 else (np/daysale) end
when daysale > 0 and np = 0 then onhandnp/1
when daysale = 0 and np > 0 then onhandnp/case when (np/1) = 0 then 1 else (np/1)end
when daysale = 0 and np = 0 then onhandnp/1
end as npremainsale,
rate*saleprice as realprice

from
(
select	a.docno,a.docdate,a.mydescription,a.creatorcode,
		h.name as workname,i.runnumber,i.countcopy,
		b.priority,isnull(b.wantday,0) as wantday,isnull(b.wantdate,getdate()) as wantdate, 
		b.itemcode,f.name1 as itemname,b.qty,b.unitcode,b.linenumber,isnull(f.itemstatus,0) as itemstatus,
		f.defbuyunitcode,isnull(c.rate,1) as rate,isnull(c.onhand,0) as onhandhmx,isnull(c.StkHMX,0) as StkRateHMX,isnull(d.onhand,0) as onhandnp,isnull(d.StkNP,0) as StkRateNP,
		c.stkunitcode,(isnull(c.onhand,0)/isnull(c.rate,1))+ (isnull(d.onhand,0)/isnull(d.rate,1)) as onhandbuyunit,
		isnull(e.remain,0) as remainorder,isnull(unitonorder,'') as unitonorder,
		isnull((select top 1 saleprice1 from dbo.bcpricelist where itemcode = b.itemcode and unitcode = f.defsaleunitcode and saletype = 0 and transporttype = 0),0) as saleprice,
		isnull((select top 1 docdate from (select a.docno,a.docdate,a.createdatetime,b.itemcode from dbo.bcarinvoice a inner join dbo.bcarinvoicesub b on a.docno = b.docno and a.docdate = b.docdate and a.arcode = b.arcode where a.iscancel = 0) as aa  where b.itemcode = itemcode  order by createdatetime desc),'') as lastdatesale,
		f.defsaleunitcode, isnull(f.leadtime,0) as leadtime,f.orderpoint,f.stockmin,f.stockmax,
		isnull((select top 1 countbasesaleqty from bchistory.dbo.TB_BC_ItemAverageSale where b.itemcode = itemcode and billtype = 0),0) as hmx,
		isnull((select top 1 countbasesaleqty from bchistory.dbo.TB_BC_ItemAverageSale where b.itemcode = itemcode and billtype = 1),0) as np,
		isnull((select top 1 datediff1 from bchistory.dbo.TB_BC_ItemAverageSale where b.itemcode = itemcode),1) as daysale,
		isnull((select top 1 price from (select a.docno,a.docdate,b.itemcode,b.qty,b.price,b.unitcode,isnull(b.discountword,'') as discountword,b.discountamount,a.createdatetime from dbo.bcapinvoice a inner join dbo.bcapinvoicesub b on a.docno = b.docno and a.docdate = b.docdate and a.apcode = b.apcode where a.iscancel = 0  and price > 0) as aa where b.itemcode = itemcode and b.unitcode = unitcode order by createdatetime desc ),0) as lastbuyprice,
		isnull((select top 1 discountword from (select a.docno,a.docdate,b.itemcode,b.qty,b.price,b.unitcode,isnull(b.discountword,'') as discountword,b.discountamount,a.createdatetime from dbo.bcapinvoice a inner join dbo.bcapinvoicesub b on a.docno = b.docno and a.docdate = b.docdate and a.apcode = b.apcode where a.iscancel = 0 ) as aa where b.itemcode = itemcode and b.unitcode = unitcode order by createdatetime desc ),'') as lastdiscountword,
		isnull((select top 1 discountamount/qty from (select a.docno,a.docdate,b.itemcode,b.qty,b.price,b.unitcode,isnull(b.discountword,'') as discountword,b.discountamount,a.createdatetime from dbo.bcapinvoice a inner join dbo.bcapinvoicesub b on a.docno = b.docno and a.docdate = b.docdate and a.apcode = b.apcode where a.iscancel = 0 ) as aa where b.itemcode = itemcode and b.unitcode = unitcode order by createdatetime desc ),0) as lastdiscountamount,
		isnull((select top 1 docno from (select a.docno,a.docdate,b.porefno,b.itemcode,b.qty,b.price,b.unitcode,isnull(b.discountword,'') as discountword,b.discountamount,a.createdatetime from dbo.bcapinvoice a inner join dbo.bcapinvoicesub b on a.docno = b.docno and a.docdate = b.docdate and a.apcode = b.apcode where a.iscancel = 0 ) as aa where b.itemcode = itemcode and b.unitcode = unitcode order by createdatetime desc ),'') as lastdocno,
		isnull(f.defbuywhcode,'') as whcode,isnull(a.departcode,'') as departcode,isnull(a.workman,'') as workman,isnull(g.name,'') as departname,
		isnull(f.baseprice,0) as baseprice
from	dbo.bcstkrequest a
		inner join	dbo.bcstkrequestsub b on a.docno = b.docno and a.docdate = b.docdate 		
		left join 
		(
			select 	*,
				case 
				when defbuyunitcode = stkunitcode then
				onhand
				else
				onhand/rate
				end as StkHMX	
			from
			(
			select 	a.code as itemcode,a.name1,isnull(a.defbuyunitcode,'') as defbuyunitcode,b.unitcode as rateunitcode,isnull(b.rate,1) as rate,c.onhand,c.unitcode as stkunitcode
			from 	bcitem a 
					inner join bcstkpacking b on a.code = b.itemcode and a.defbuyunitcode = b.unitcode
					inner join (	select itemcode,sum(qty) as onhand,unitcode
						from bcstklocation 
						where	shelfcode in ('AVL','BK3','RSV','PAK','SHW','VND')
						group by itemcode,unitcode) c on a.code = c.itemcode 
			)	as result
		) c on b.itemcode = c.itemcode 
		left join
		(
			select 	*,
				case 
				when defbuyunitcode = stkunitcode then
				onhand
				else
				onhand/rate
				end as StkNP
			from
			(
			select 	a.code as itemcode,a.name1,isnull(a.defbuyunitcode,'') as defbuyunitcode,b.unitcode as rateunitcode,isnull(b.rate,1) as rate,c.onhand,c.unitcode as stkunitcode
			from 	bcitem a 
					inner join bcstkpacking b on a.code = b.itemcode and a.defbuyunitcode = b.unitcode
					inner join (	select itemcode,sum(qty) as onhand,unitcode
						from bcstklocation 
						where	shelfcode in  ('BK1','BK2')
						group by itemcode,unitcode) c on a.code = c.itemcode 
			)	as result
		) d on b.itemcode = d.itemcode 
		left join	
		(
			select itemcode,sum(remainqty) as remain ,isnull(unitcode,'') as unitonorder from 
			(
				select * from
				(
					select a.*,b.docno as poxno from 
					(
						select * from
						(
							select 	a.docno,a.itemcode,b.porefno,a.qty,a.remainqty,a.unitcode,a.docdate,a.whcode
							from bcpurchaseordersub a 
								left join bcapinvoicesub b  on a.docno = b.porefno and a.itemcode = b.itemcode and a.unitcode = b.unitcode
							where a.iscancel = 0 and year(a.docdate) >= '2008' 
						) as aaa
						where porefno is null and whcode  not in  ('099','S00')
					) as a left join BCPOVOIDsub b on a.docno = b.pono and a.itemcode = b.itemcode and a.unitcode = b.unitcode
				) as aa where poxno is null 
			) as aaaa
			group  by itemcode,unitcode
			) e on b.itemcode = e.itemcode and b.unitcode = e.unitonorder
			left join dbo.bcitem f on b.itemcode = f.code and activestatus = 1
			left join bcdepartment g on a.departcode = g.code
			left join bcsale h on a.workman = h.code
			left join npmaster.dbo.TB_DC_RunNumberDocumentLogs i on a.docno = i.docno
where		a.iscancel = 0 and a.docno = @vDocno
) as		result
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateCarLicense
@vDocNo as nvarchar(20),
@vSOCountNumber as int,
@vCarLicense as nvarchar(30)
as
set	dateformat dmy

declare	@vExist as int
set		@vExist = (select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_QueueRequestPickingMaster where docno = @vDocNo and socountnumber = @vSOCountNumber)

if		@vExist > 0 
begin
update 	npmaster.dbo.TB_NP_QueueRequestPickingMaster set carlicense = @vCarLicense where docno = @vDocNo and socountnumber = @vSOCountNumber
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateCheckQtyQue
@vMergeNo as nvarchar(30),
@vUserID as nvarchar(30),
@vItemCode as nvarchar(20),
@vCheckQty as money
as

set		dateformat dmy

declare	@vCalcQTY as money


update	npmaster.dbo.tb_np_quepickcentermaster
set		mergeno = @vMergeNo,checker = @vUserID,checkoutdatetime = getdate()
where	queid in (	 
				select	queid 
				from	npmaster.dbo.tb_np_driveinmergetemp
				where	docno = @vMergeNo and 
						docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
				) and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)


update	npmaster.dbo.tb_np_quepickcenterSub
set		checkqty = @vCheckQty,mergeno = @vMergeNo
where	queid in (	 
				select	queid 
				from	npmaster.dbo.tb_np_driveinmergetemp
				where	docno = @vMergeNo and 
						docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) and 
						itemcode = @vItemCode 
				) and itemcode = @vItemCode  and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdateCountOfPrintPicking 
@vUserID as nvarchar(20),
@vPickingNo as nvarchar(20),
@vShelfGroup as nvarchar(2)
as
Update 	npmaster.dbo. NP_PickingSlip_Logs 
set 	lastuserprint = @vUserID , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 
where 	pickingno = @vPickingNo  and shelfgroup = @vShelfGroup

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateCustItemReceiptStatus
@vReceiptID as nvarchar(20),
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(20)
as
set	dateformat dmy
update 	npmaster.dbo.TB_QUE_CustItemReceipt set status = 1
	where docno = @vReceiptID and invoiceno = @vInvoiceNo and whcode = @vWHCode

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateDriveInCancelCheckOut
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vItemCode as nvarchar(20)
as

set		dateformat dmy

update	npmaster.dbo.tb_np_driveinslipsub
set		iscancel = 1 , confirmqty = 0 , invqty = 0 
where	docno = @vDocno and  itemcode = @vItemCode and iscancel = 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateDriveInCheckerItemCheckOut
@vDocno as nvarchar(20),
@vDocdate as nvarchar(20),
@vItemCode as nvarchar(20),
@vKeyQTY as money,
@vItemAmout as money
as

set		dateformat dmy

update	npmaster.dbo.tb_np_driveinslipmaster
set		ischecker =1
where	docno = @vDocno and iscancel = 0


update	npmaster.dbo.tb_np_driveinslipsub
set		confirmqty = @vKeyQTY ,amount= @vItemAmout
where	docno = @vDocno and  itemcode = @vItemCode and iscancel = 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateDriveInMergeTempConfirm
@vDocNo nvarchar(20),
@vDocDate nvarchar(20),
@vItemCode nvarchar(30),
@vPosNo as nvarchar(30)

as

set		dateformat dmy

update	npmaster.dbo.TB_NP_DriveInMergeTemp
set		isconfirm = 1,posbillno = @vPosNo,confirmdatetime = getdate()
where	docno = @vDocNo and docdate = @vDocDate and itemcode = @vItemCode

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateHoldBillQtyQue
@vMergeNo as nvarchar(30),
@vHoldBillNo as nvarchar(30),
@vCashierCode as nvarchar(30),
@vItemCode as nvarchar(20),
@vInvQty as money
as

set		dateformat dmy

declare	@vCalcQTY as money


update	npmaster.dbo.tb_np_quepickcentermaster
set		holdbillno = @vHoldBillNo,isconfirm = 1,cashiercode = @vCashierCode
where	queid in (	 
				select	queid 
				from	npmaster.dbo.tb_np_driveinmergetemp
				where	docno = @vMergeNo and 
						docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
				) and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)


update	npmaster.dbo.tb_np_quepickcenterSub
set		invqty = @vInvQty,holdbillno = @vHoldBillNo
where	queid in (	 
				select	queid 
				from	npmaster.dbo.tb_np_driveinmergetemp
				where	docno = @vMergeNo and 
						docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) and 
						itemcode = @vItemCode 
				) and itemcode = @vItemCode  and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure dbo.USP_NP_UpdateIsCancelQuotation
@vDocno as nvarchar(20)
as
declare @vExist as int 
set dateformat dmy

set @vExist = (select count(docno) as docno from npmaster.dbo.tb_np_quotation where docno = @vDocno)

if @vExist <> 0
begin
update npmaster.dbo.tb_np_quotation set iscancel = 1 where docno = @vDocno
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateMydescriptionQueueManagement
@vDocno as nvarchar(20),
@vTimeID as int,
@vMyDescription as nvarchar(150)
as

set	dateformat dmy
update 	npmaster.dbo.TB_NP_QueueManagement 
set	Mydescription = @vMyDescription
where 	docno = @vDocno  and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and timeid = @vTimeID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateMydescriptionQueueManagement1
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vTimeID as int,
@vMyDescription as nvarchar(150)
as

set	dateformat dmy
update 	npmaster.dbo.TB_NP_QueueManagement 
set	Mydescription = @vMyDescription
where 	docno = @vDocno  and docdate = @vDocDate and timeid = @vTimeID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateMydescriptionQueueManagement2
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vTimeID as int,
@vReasonID as int,
@vMyDescription as nvarchar(150)
as

set	dateformat dmy
update 	npmaster.dbo.TB_NP_QueueManagement 
set	Mydescription = @vMyDescription,pickreason = @vReasonID
where 	docno = @vDocno  and docdate = @vDocDate and timeid = @vTimeID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateMydescriptionQueueManagement3
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vTimeID as int,
@vReasonID as int,
@vMyDescription as nvarchar(150)
as

set	dateformat dmy
update 	npmaster.dbo.TB_NP_QueueManagement_Test 
set	Mydescription = @vMyDescription,pickreason = @vReasonID
where 	docno = @vDocno  and docdate = @vDocDate and timeid = @vTimeID





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdateNewDocNo
@vGroupDoc as int
as
Update 	npmaster.dbo.NP_Generate_DocNo  
set 	autonumber = autonumber + 1  
where 	headertype = @vGroupDoc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdatePayGoods
@vPayNumber as nvarchar(20)
as
set	dateformat dmy

update 	npmaster.dbo.np_paygoods 
set	checked = 1 --จ่ายของแล้ว
where 	paynumber = @vPayNumber and checked = 0  and 
	year(paydatetime) = year(getdate()) and 
	month(paydatetime) = month(getdate()) and 
	day(paydatetime) = day(getdate()) 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdatePayItemQtyQue
@vQueID as int,
@vItemCode as nvarchar(20),
@vQty as money
as

set		dateformat dmy


update	npmaster.dbo.tb_np_quepickcenterMaster
set	quereceived = 1 ,querecstatus = 1
where	queid = @vQueID and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)


update	npmaster.dbo.tb_np_quepickcenterSub
set		oncarqty = @vQty
where	queid = @vQueID and itemcode = @vItemCode  and quedocdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdatePickQueCenterSub
@vQueNo as int,
@vQueDocDate as nvarchar(20),
@vItemCode as nvarchar(30),
@vPickQTY as money
as

set 	dateformat dmy

update	npmaster.dbo.TB_NP_QuePickCenterSub 
set		pickqty = @vPickQTY
where	queid = @vQueNo and quedocdate = @vQueDocDate and itemcode = @vItemCode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdatePrintStatusQueueManagement
@vDocno as nvarchar(20),
@vPicker as nvarchar(20),
@vStatus as int,
@vTimeID as int,
@vPickStatus as int
as

declare @vStatusDesc as nvarchar(20)

if @vStatus = 1
begin 
	update 	npmaster.dbo.TB_NP_QueueManagement 
	set	status = 1,statusdesc = 'จัดของ',picker = @vPicker,startdatetime = getdate()
	where 	docno = @vDocno  and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and timeid = @vTimeID
end
if @vStatus = 2
begin 

if @vPickStatus = 1
begin
set @vStatusDesc ='ครบ'
end
if @vPickStatus = 2 
begin
set @vStatusDesc = 'ไม่ครบ'
end

update 	npmaster.dbo.TB_NP_QueueManagement 
	set	status = 2,statusdesc = @vStatusDesc,stopdatetime = getdate(),pickingstatus = @vPickStatus
	where 	docno = @vDocno and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and timeid = @vTimeID
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdatePrintStatusQueueManagement1
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vPicker as nvarchar(50),
@vStatus as int,
@vTimeID as int,
@vPickStatus as int
as

set	dateformat dmy

declare @vStatusDesc as nvarchar(20)

if @vStatus = 1
begin 
	update 	npmaster.dbo.TB_NP_QueueManagement 
	set	status = 1,statusdesc = 'จัดของ',picker = @vPicker,startdatetime = getdate()
	where 	docno = @vDocno  and docdate = @vDocDate and timeid = @vTimeID
end
if @vStatus = 2
begin 

if @vPickStatus = 1
begin
set @vStatusDesc ='ครบ'
end
if @vPickStatus = 2 
begin
set @vStatusDesc = 'ไม่ครบ'
end

update 	npmaster.dbo.TB_NP_QueueManagement 
	set	status = 2,statusdesc = @vStatusDesc,stopdatetime = getdate(),pickingstatus = @vPickStatus
	where 	docno = @vDocno and docdate = @vDocDate and timeid = @vTimeID
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdatePrintStatusQueueManagement2
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vPicker as nvarchar(20),
@vStatus as int,
@vTimeID as int,
@vPickStatus as int
as

--Test 

set	dateformat dmy

declare @vStatusDesc as nvarchar(20)

if @vStatus = 1
begin 
	update 	npmaster.dbo.TB_NP_QueueManagement_Test 
	set	status = 1,statusdesc = 'จัดของ',picker = @vPicker,startdatetime = getdate()
	where 	docno = @vDocno  and docdate = @vDocDate and timeid = @vTimeID
end
if @vStatus = 2
begin 

if @vPickStatus = 1
begin
set @vStatusDesc ='ครบ'
end
if @vPickStatus = 2 
begin
set @vStatusDesc = 'ไม่ครบ'
end

update 	npmaster.dbo.TB_NP_QueueManagement_Test
	set	status = 2,statusdesc = @vStatusDesc,stopdatetime = getdate(),pickingstatus = @vPickStatus
	where 	docno = @vDocno and docdate = @vDocDate and timeid = @vTimeID
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQuePickCenterReason
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vReasonID as int,
@vMyDescription as nvarchar(150)
as

set	dateformat dmy
update 	npmaster.dbo.TB_NP_QuePickCenterMaster 
set	quereasondesc = @vMyDescription,quereason = @vReasonID
where 	queid = @vQueID  and quedocdate = @vQueDocDate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueReceived
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vSaleOrderNo as nvarchar(25),
@vStatus as smallint,
@vDescription as nvarchar(300)

as
set	dateformat dmy

declare	@vCheckDescription as int
declare @vDescriptionExist as nvarchar(300)

set	@vCheckDescription = (select top 1 isnull(len(quereasondesc),0) as vLen from npmaster.dbo.TB_NP_QuePickCenterMaster  where queid =@vQueID  and quedocdate = @vQueDocDate and docno = @vSaleOrderNo)

if 	@vCheckDescription > 0 
begin
set		@vDescriptionExist = (select ltrim(rtrim(quereasondesc))+'/'+ltrim(rtrim(@vDescription)) as mydescription from npmaster.dbo.TB_NP_QuePickCenterMaster  where queid =@vQueID  and quedocdate = @vQueDocDate and docno = @vSaleOrderNo)
update 	npmaster.dbo.TB_NP_QuePickCenterMaster 
set 	quereceived = 1,querecstatus = @vStatus,querecdate = getdate(),quereasondesc = @vDescriptionExist
where 	queid =@vQueID  and quedocdate = @vQueDocDate and docno = @vSaleOrderNo
end

if 	@vCheckDescription = 0 
begin
update 	npmaster.dbo.TB_NP_QuePickCenterMaster 
set 	quereceived = 1,querecstatus =@vStatus,querecdate = getdate(),quereasondesc = @vDescription
where 	queid =@vQueID  and quedocdate = @vQueDocDate and docno = @vSaleOrderNo
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueStatusDetails
@vQueID as int,
@vQueDocDate as nvarchar(20),
@vPicker as nvarchar(50),
@vStatus as int,
@vQueTime as int,
@vPickStatus as int
as

set	dateformat dmy

declare @vStatusDesc as nvarchar(20)

if @vStatus = 1
begin 
	update 	npmaster.dbo.TB_NP_QuePickCenterMaster 
	set		questatus = 1,quedescription = 'จัดของ',quepicker = @vPicker,questart = getdate()
	where 	queid = @vQueID  and quedocdate = @vQueDocDate and Quetime = @vQueTime
end
if @vStatus = 2
begin 

if @vPickStatus = 1
begin
set @vStatusDesc ='ครบ'
end
if @vPickStatus = 2 
begin
set @vStatusDesc = 'ไม่ครบ'
end

update 	npmaster.dbo.TB_NP_QuePickCenterMaster 
set		questatus = 2,quedescription = @vStatusDesc,questop = getdate(),quepickstatus = @vPickStatus
where 	queid = @vQueID  and quedocdate = @vQueDocDate and Quetime = @vQueTime
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdateQueueCustRec
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20)
as

set		dateformat dmy

update	npmaster.dbo.tb_np_queuemanagement 
set		isreceived = 0 
where 	docno = @vDocno and docdate = @vDocDate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdateQueueCustRec_Test
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20)
as

set		dateformat dmy

update	npmaster.dbo.tb_np_queuemanagement_Test 
set		isreceived = 0 
where 	docno = @vDocno and docdate = @vDocDate

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_NP_UpdateQueueMyDescription
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vMyDescription as nvarchar(200)
as

set		dateformat dmy

update	npmaster.dbo.tb_np_queuemanagement 
set		mydescription = @vMyDescription 
where 	docno = @vDocno and docdate = @vDocDate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueueMyDescription1
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vMyDescription as nvarchar(200)
as

set		dateformat dmy
--Test
update	npmaster.dbo.tb_np_queuemanagement_test 
set		mydescription = @vMyDescription 
where 	docno = @vDocno and docdate = @vDocDate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueuePrintStatus
@vDocno as nvarchar(20)
as

set	dateformat dmy

update npmaster.dbo.TB_NP_CheckQueuePrint set printstatus = 1 where docno = @vDocno and printstatus = 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueueReceived
@vPickingNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20)
as
set	dateformat dmy
update 	npmaster.dbo.tb_np_queuemanagement set isreceived = 1
	where docno = @vPickingNo and saleorderno = @vSaleOrderNo
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueueReceivedStatus
@vPickingNo as nvarchar(20),
@vSaleOrderNo as nvarchar(20),
@vStatus as smallint
--@vDescription as nvarchar(200)

as
set	dateformat dmy

declare	@vCheckDescription as int
declare  @vDescriptionExist as nvarchar(200)

set	@vCheckDescription = (select isnull(len(mydescription),0) as vLen from npmaster.dbo.tb_np_queuemanagement  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)
/*
if 	@vCheckDescription > 0 
begin
set	@vDescriptionExist = (select ltrim(rtrim(mydescription))+'/'+ltrim(rtrim(@vDescription)) as mydescription from npmaster.dbo.tb_np_queuemanagement  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)
update 	npmaster.dbo.tb_np_queuemanagement 
set 	isreceived = @vStatus,custrecdatetime = getdate(),mydescription = @vDescriptionExist
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
end
*/
--if 	@vCheckDescription = 0 
--begin
update 	npmaster.dbo.tb_np_queuemanagement 
set 	isreceived = @vStatus,custrecdatetime = getdate()--,mydescription = @vDescription
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
--end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueueReceivedStatus1
@vPickingNo as nvarchar(20),
@vSaleOrderNo as nvarchar(25),
@vStatus as smallint,
@vDescription as nvarchar(200)

as
set	dateformat dmy

declare	@vCheckDescription as int
declare  @vDescriptionExist as nvarchar(200)

set	@vCheckDescription = (select top 1 isnull(len(mydescription),0) as vLen from npmaster.dbo.tb_np_queuemanagement  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)

if 	@vCheckDescription > 0 
begin
set	@vDescriptionExist = (select ltrim(rtrim(mydescription))+'/'+ltrim(rtrim(@vDescription)) as mydescription from npmaster.dbo.tb_np_queuemanagement  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)
update 	npmaster.dbo.tb_np_queuemanagement 
set 	isreceived = @vStatus,custrecdatetime = getdate(),mydescription = @vDescriptionExist
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
end

if 	@vCheckDescription = 0 
begin
update 	npmaster.dbo.tb_np_queuemanagement 
set 	isreceived = @vStatus,custrecdatetime = getdate(),mydescription = @vDescription
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateQueueReceivedStatus2
@vPickingNo as nvarchar(20),
@vSaleOrderNo as nvarchar(25),
@vStatus as smallint,
@vDescription as nvarchar(200)

as
set	dateformat dmy

declare	@vCheckDescription as int
declare  @vDescriptionExist as nvarchar(200)

set	@vCheckDescription = (select top 1 isnull(len(mydescription),0) as vLen from npmaster.dbo.tb_np_queuemanagement_Test  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)

if 	@vCheckDescription > 0 
begin
set	@vDescriptionExist = (select ltrim(rtrim(mydescription))+'/'+ltrim(rtrim(@vDescription)) as mydescription from npmaster.dbo.tb_np_queuemanagement_Test  where docno =@vPickingNo  and saleorderno = @vSaleOrderNo)
update 	npmaster.dbo.tb_np_queuemanagement_Test
set 	isreceived = @vStatus,custrecdatetime = getdate(),mydescription = @vDescriptionExist
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
end

if 	@vCheckDescription = 0 
begin
update 	npmaster.dbo.tb_np_queuemanagement_Test 
set 	isreceived = @vStatus,custrecdatetime = getdate(),mydescription = @vDescription
where 	docno = @vPickingNo and saleorderno = @vSaleOrderNo
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateRequestPickingQueue
@vDocNo as nvarchar(20),
@vCount as int,
@vQueueID as nvarchar(10),
@vShelfGroup as nvarchar(10)
as
set		dateformat dmy

update	npmaster.dbo.TB_NP_QueueRequestPickingMaster
set		queueno = @vQueueID
where	docno = @vDocno and socountnumber = @vCount and shelfgroup = @vShelfGroup
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateSendQueuePicking
@vType as int,
@vDocNo as nvarchar(30)
as

set		dateformat dmy

if	@vType = 1
begin
update	npmaster.dbo.TB_NP_PickingRequestMaster
set		issendque = 1 
where	docno = @vDocNo
end

if	@vType = 3
begin
update	npmaster.dbo.TB_NP_DriveInSlipMaster
set		issendque = 1 
where	docno = @vDocNo
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_NP_UpdateStatusCustItemReceipt
@vInputType as int,
@vDocno as nvarchar(20),
@vPrintCount as int,
@vChecker as nvarchar(30)

as

if @vInputType = 0
begin 
	update 	npmaster.dbo.TB_QUE_CustItemReceipt 
	set	Status = 1, PrintStatus = 1, PrintDateTime = getdate()
	where 	docno = @vDocno  and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and printcount = @vPrintCount
end

if @vInputType = 1
begin 
	update 	npmaster.dbo.TB_QUE_CustItemReceipt 
	set	status = 2,IsComplete = 1,CheckDateTime = getdate(),Checker = @vChecker
	where 	docno = @vDocno and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and printcount = @vPrintCount
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_QUE_CheckIsCancelInvoice
@vDocno as nvarchar(20)
as

set	dateformat dmy
set	language us_english
select 	iscancel,docno from dbo.bcarinvoice where docno = @vDocno

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_QUE_CheckItemReceiptUpdateCancel
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(10)
as
set	dateformat dmy
set	language us_english
select 	isnull(count(docno ),0) as vCount
from 	npmaster.dbo.TB_QUE_CustItemReceipt 
where 	invoiceno = @vInvoiceNo and whcode = @vWHCode and status <> 2 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_CheckShelfGroup
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(10)
as 
set	dateformat dmy
set	language us_english

select  distinct 
	top 100 percent a.DocNo,isnull(a.sorefno,'') as sorefno,
	left(case 
	when b.whcode = '014' and y.itemtype is not null  then 'M'
	when b.whcode = '014' and y.itemtype is  null  then 'H'
	when b.whcode = '020'  then 'H'
	when b.whcode = '015'  then 'C'
	when b.whcode = '016'  then 'Y'
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'D') not in ('A','B') then 'D' 
	when b.whcode = '010' and isnull(left(z.shelfcode,1),'')  in ('A','B') then left(z.shelfcode,1)
	end,1) as shelfgroup

FROM	dbo.bcarinvoice a  
	LEFT OUTER JOIN dbo.bcarinvoicesub b ON a.DocNo = b.DocNo and a.docdate = b.docdate
	left join bcnp.dbo.bcitem g on b.itemcode = g.code
	left join (select distinct productcode,whcode,(select top 1 shelfcode from bcnp.dbo.bcrecproduct where productcode = aaa.productcode and whcode = aaa.whcode and shelfcode <> '-') as shelfcode from (select productcode,whcode,shelfcode from bcnp.dbo.bcrecproduct) as aaa )z 
	on b.itemcode = z.productcode  and b.whcode = z.whcode
	left join npmaster.dbo.NP_ItemOutLet y on g.typecode = y.itemtype	
WHERE    a.iscancel = 0 and b.whcode not in ('011','000','070','097','099')   and a.docno = @vInvoiceNo and b.whcode = @vWHCode
order by a.docno




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_InsertCustItemReceipt
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vSaleType as int,
@vARCode as nvarchar(20),
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vZoneID as nvarchar(2)

as
declare	@vStatus as int
declare	@vPrintCount as int
declare	@vIsCancel as int
declare	@vIsComplete as int
declare	@vQueueDateTime as nvarchar(20)
declare 	@vPrintStatus as int

set	@vStatus = 0
set 	@vPrintCount = 1
set	@vIsCancel = 0
set 	@vIsComplete = 0
set	@vQueueDateTime = getdate()
set	@vPrintStatus = 0

set	dateformat dmy
set	language us_english

insert	into npmaster.dbo.TB_QUE_CustItemReceipt(DocNo,DocDate,QueueDateTime,PrintStatus,IsCancel,IsComplete,SaleType,ARCode,Status,InvoiceNo,WHCode,ZoneID,PrintCount)
values(@vDocNo,@vDocDate,@vQueueDateTime,@vPrintStatus,@vIsCancel,@vIsComplete,@vSaleType,@vARCode,@vStatus,@vInvoiceNo,@vWHCode,@vZoneID,@vPrintCount)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_InsertCustItemReceiptSub
@vDocNo as nvarchar(20),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(150),
@vQTY as int,
@vUnitCode as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(15),
@vLineNumber as int

as
declare	@vReceiptQTY as int
declare	@vPrintCount as int

set	@vReceiptQTY = 0
set 	@vPrintCount = 1
set	dateformat dmy
set	language us_english

insert	into npmaster.dbo.TB_QUE_CustItemReceiptSub(DocNo,DocDate,ItemCode,ItemName,QTY,ReceiptQTY,UnitCode,WHCode,ShelfCode,PrintCount,LineNumber)
values(@vDocNo,@vDocDate,@vItemCode,@vItemName,@vQTY,@vReceiptQTY,@vUnitCode,@vWHCode,@vShelfCode,@vPrintCount,@vLineNumber)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create	procedure dbo.USP_QUE_InsertLineItemReceipt
@vDocno as nvarchar(20),
@vDocDate as nvarchar(20),
@vItemCode as nvarchar(20),
@vItemName as nvarchar(120),
@vQTY as decimal,
@vReceiptQTY as decimal,
@vUnitCode as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfCode as nvarchar(10),
@vPrintCount as int,
@vLineNUmber as int
as
set	dateformat dmy
set	language us_english

insert	into npmaster.dbo.TB_QUE_CustItemReceiptSub (Docno,DocDate,ItemCode,ItemName,QTY,ReceiptQTY,UnitCode,WHCode,ShelfCode,PrintCount,LineNUmber)
values(@vDocno,@vDocDate,@vItemCode,@vItemName,@vQTY,@vReceiptQTY,@vUnitCode,@vWHCode,@vShelfCode,@vPrintCount,@vLineNUmber)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_SearchCheckCustItemReceipt
@vZoneID as int
as

set 	dateformat dmy
set	language us_english

if 	@vZoneID = 1
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid = '01' and iscancel = 0
order	by queuedatetime,invoiceno 
end

if 	@vZoneID = 2
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid in ('02','03')and iscancel = 0
order	by queuedatetime,invoiceno 
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_SearchCustItemReceipt
@vZoneID as int
as

set 	dateformat dmy
set	language us_english

if 	@vZoneID = 1
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 0 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid = '01' and iscancel = 0 and status = 0
order	by docno desc
end

if 	@vZoneID = 2
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 0 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid in ('02','03') and iscancel = 0 and status = 0
order	by docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_SearchCustItemReceiptChecking
@vZoneID as int
as

set 	dateformat dmy
set	language us_english

if 	@vZoneID = 1
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid = '01' and iscancel = 0 and status = 1
order	by docno desc
end

if 	@vZoneID = 2
begin
select 	a.*,b.name1 as arname
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join dbo.bcar b on a.arcode = b.code
where 	status = 1 and year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and
	a.zoneid in ('02','03') and iscancel = 0 and status = 1
order	by docno desc
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_SearchWHCodeCustReceiptItem
@vDocno as nvarchar(20),
@vWHCode as nvarchar(10)
as

set	dateformat dmy
set	language us_english
select 	a.docno,a.docdate,a.arcode,a.invoiceno,a.whcode,c.itemcode,c.itemname,c.qty,c.unitcode,c.shelfcode,d.name1 as arname 
from 	npmaster.dbo.TB_QUE_CustItemReceipt a
	left join bcnp.dbo.bcarinvoice b on a.invoiceno = b.docno
	left join bcnp.dbo.bcarinvoicesub c on b.docno = c.docno and b.docdate = b.docdate
	left join bcnp.dbo.bcar d on a.arcode = code
where 	year(a.docdate) = year(getdate()) and month(a.docdate) = month(getdate())  and day(a.docdate) = day(getdate()) and 
	a.docno = @vDocno and c.whcode = @vWHCode and a.iscancel = 0
order	by c.linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_UpdateCancelCustItemReceipt
@vInvoiceNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vUserID as nvarchar(30)
as

set	dateformat dmy
set	language us_english
update 	npmaster.dbo.TB_QUE_CustItemReceipt set iscancel = 1,cancelcode = @vUserID,canceldatetime = getdate() where invoiceno = @vInvoiceNo and whcode = @vWHCode

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_QUE_UpdateIsReceivedItem
@vRefNo as nvarchar(20),
@vWHCode as nvarchar(10),
@vShelfGroup as nvarchar(10)
as

set	dateformat dmy
set	language us_english
update 	npmaster.dbo.tb_np_queuemanagement set isreceived = 1 , status = 2 , pickingstatus = 1 where saleorderno = @vRefNo and whcode = @vWHCode  and shelfgroup = @vShelfGroup
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchDriveInMaster
@vDIPoint as nvarchar(10),
@vSearch as nvarchar(20)
as

set	dateformat dmy

if	@vSearch <> '' 
begin
	select	*
	from
	(
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and refno like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	union
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and docno like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	union
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and a.arcode like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	union
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and b.name1 like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	union
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and a.salecode like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	union
	select 	a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where 	pickzone = @vDIPoint and c.name like '%'+@vSearch+'%' and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	) as	result
	order 	by docdate desc,docno
end

if	@vSearch = '' 
begin
	select 	top 10 a.docno,a.docdate,isnull(a.refno,'') as refid,a.pickzone,isnull(totalnetamount,0) as totalnetamount,iscancel,isconfirm,
			isnull(a.arcode,'') as arcode,isnull(a.salecode,'') as salecode,isnull(b.name1,'') as arname,isnull(c.name,'') as salename
	from 	npmaster.dbo.tb_np_driveinslipmaster a
			left join dbo.bcar b on a.arcode = b.code
			left join dbo.bcsale c on a.salecode = c.code
	where	pickzone = @vDIPoint and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime) 
	order 	by a.docdate desc,a.docno
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchItemPickUp
@vRefNo as nvarchar(20)
as

set	dateformat dmy

if	@vRefNo <> '' 
begin
	select 	a.docno,a.docdate,a.id,a.refid,a.pickzone,isnull(ismerge,0) as ismerge,
		itemcode,itemname,whcode,shelfcode,qty,pickqty,unitcode,price,amount,isnull(barcode,'') as barcode
	from 	npmaster.dbo.tb_np_driveinslipmaster a
		inner join npmaster.dbo.tb_np_driveinslipsub b on a.docno = b.docno and a.docdate = b.docdate
	where 	refid = @vRefNo and ismerge = 0 and a.iscancel = 0 and b.iscancel = 0 and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	order 	by a.docdate,a.docno
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchItemPickUpCancel
@vRefNo as nvarchar(20)
as

set	dateformat dmy

if	@vRefNo <> '' 
begin
	select 	a.docno,a.docdate,a.id,a.refid,a.pickzone,itemcode,itemname,whcode,shelfcode,qty,pickqty,confirmqty,unitcode,price,amount,isnull(barcode,'')  as barcode
	from 	npmaster.dbo.tb_np_driveinslipmaster a
		inner join npmaster.dbo.tb_np_driveinslipsub b on a.docno = b.docno and a.docdate = b.docdate
	where 	refid = @vRefNo and b.iscancel = 1 and a.docdate = cast(rtrim(day(getdate()))+'/'+rtrim(month(getdate()))+'/'+rtrim(year(getdate())) as datetime)
	order 	by a.docdate,a.docno
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchPickUp
@vDocNo as nvarchar(20)
as

set	dateformat dmy


select 	a.docno,a.docdate,a.id,a.refid,a.pickzone,totalnetamount,itemcode,itemname,whcode,shelfcode,qty,pickqty,unitcode,price,amount,barcode
from 	npmaster.dbo.tb_np_driveinslipmaster a
		inner join npmaster.dbo.tb_np_driveinslipsub b on a.docno = b.docno and a.docdate = b.docdate
where 	a.docno = @vDocNo
order 	by a.docdate,a.docno,linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchReqPickingDetails
@vDocNo as nvarchar(30)

as

set		dateformat dmy

select	a.DocNo,a.DocDate,a.ARCode,a.SaleCode,isnull(d.name,'') as salename,issendque,isnull(refno,'') as refno,isnull(memberid,'') as memberid,a.BeforeTaxAmount,a.TaxAmount,a.NetDebtAmount,a.IsConditionSend,a.ReqTime,
	a.iscancel,isnull(a.MyDescription ,'') as MyDescription,
	b.ItemCode,c.name1 as itemname,b.QTY,b.Unitcode,b.Price,b.DisCountWord,b.DisCountAmount,b.NetAmount,b.WHCode,b.ShelfCode,b.ShelfID,b.ZoneID,b.BarCode,b.LineNumber
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join npmaster.dbo.TB_NP_PickingRequestSub b on a.docno =b.docno 
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.docno = @vDocNo 
order	by b.linenumber
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchReqPickingInformation
@vDocNo as nvarchar(30),
@vTimeID as int
as

set		dateformat dmy

select	a.QueID,a.QueDocDate,a.docno,a.ARCode,isnull(c.name1,'') as arname,isnull(b.ZoneID,'') as zoneid
from	npmaster.dbo.TB_NP_QuePickCenterMaster a 
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid =b.queid and a.quedocdate = b.quedocdate and a.quetime = b.quetime
		left join dbo.bcar c on a.arcode = c.code
where	a.docno = @vDocNo and a.QueTime = @vTimeID
group	by a.QueID,a.QueDocDate,a.docno,a.ARCode,c.name1,b.ZoneID
order	by a.QueID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchReqPickingInformationLastSend
@vDocNo as nvarchar(30),
@vQueDocDate as nvarchar(20),
@vTimeID as int
as

set		dateformat dmy


select	a.queid,b.quedocdate,a.docno,a.quedate,b.itemcode,isnull(c.name1,'') as itemname,isnull(b.qty,0) as qty,isnull(b.pickqty,0) as pickqty,b.unitcode,
		quezone,isnull(a.quepicker,'') as quepicker, 
		case	when questatus = 2 and qty > pickqty then 'ไม่ครบ' 
				when questatus = 2 and qty = pickqty then 'ครบ' 
				when questatus = 2 and qty < pickqty then 'เกิน'
				else  isnull(a.quedescription,'') 
		end		as quedescription

from	npmaster.dbo.TB_NP_QuePickCenterMaster a
		left join npmaster.dbo.TB_NP_QuePickCenterSub b on a.queid = b.queid and a.quedocdate = b.quedocdate and a.quetime = b.quetime
		left join dbo.bcitem c on b.itemcode = c.code
where	a.docno = @vDocNo and a.quedocdate = @vQueDocDate and 
		a.quetime = @vTimeID-- (select top 1 isnull(quetime,0) as quetime from npmaster.dbo.TB_NP_QuePickCenterMaster where docno = @vDocNo and quedocdate = @vQueDocDate order by quetime desc)
order	by a.quedocdate,a.queid
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.usp_np_SearchReqPickingItemZone
@vType as int,
@vDocNo as nvarchar(30),
@vZoneID as nvarchar(10),
@vTimeID as int

as

set		dateformat dmy

if 	@vType = 1 
begin
select	a.DocNo,a.DocDate,a.ARCode,a.SaleCode,isnull(d.name,'') as salename,a.BeforeTaxAmount,a.TaxAmount,a.NetDebtAmount,a.IsConditionSend,a.ReqTime,isnull(a.MyDescription ,'') as MyDescription,
		b.ItemCode,c.name1 as itemname,b.QTY,b.Unitcode,b.Price,b.DisCountWord,b.DisCountAmount,b.NetAmount,b.WHCode,b.ShelfCode,b.ShelfID,b.ZoneID,b.BarCode,b.LineNumber
from	npmaster.dbo.TB_NP_PickingRequestMaster a 
		left join npmaster.dbo.TB_NP_PickingRequestSub b on a.docno =b.docno 
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.docno = @vDocNo and b.zoneid = @vZoneID
order	by b.linenumber
end

if 	@vType = 2 
begin
set		dateformat dmy

select	a.DocNo,a.DocDate,a.ARCode,a.SaleCode,isnull(d.name,'') as salename,sumofitemamount as BeforeTaxAmount,TaxAmount,netamount as NetDebtAmount,
		a.IsConditionSend,'' as ReqTime,'' as MyDescription,
		b.ItemCode,c.name1 as itemname,b.reqQTY as qty,b.Unitcode,Price,'' as DisCountWord,DisCountAmount,itemamount as NetAmount,b.WHCode,b.ShelfCode,'' as ShelfID,b.ZoneID,'' as BarCode,b.LineNumber
from	npmaster.dbo.TB_NP_QueueRequestPickingMaster a 
		left join npmaster.dbo.TB_NP_QueueRequestPicking b on a.docno =b.docno and a.shelfgroup = b.zoneid and a.socountnumber = b.socountnumber
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.docno = @vDocNo and b.zoneid = @vZoneID and a.socountnumber = @vTimeID
order	by b.linenumber
end


if 	@vType = 3 
begin
set		dateformat dmy

select	a.DocNo,a.DocDate,a.ARCode,a.SaleCode,isnull(d.name,'') as salename,0 as BeforeTaxAmount,0 as TaxAmount,0 as NetDebtAmount,0 as IsConditionSend,'' as ReqTime,'' as MyDescription,
	b.ItemCode,c.name1 as itemname,b.qty,b.Unitcode,0 as Price,0 as DisCountWord,'' as DisCountAmount,0 as NetAmount,b.WHCode,b.ShelfCode,'' as ShelfID,b.ZoneID,'' as BarCode,b.LineNumber
from	npmaster.dbo.TB_NP_DriveInSlipMaster a 
		left join npmaster.dbo.TB_NP_DriveInSlipSub b on a.docno =b.docno and a.docdate = b.docdate
		left join dbo.bcitem c on b.itemcode = c.code
		left join dbo.bcsale d on a.salecode = d.code
where	a.docno = @vDocNo and b.zoneid = @vZoneID 
order	by b.linenumber
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



