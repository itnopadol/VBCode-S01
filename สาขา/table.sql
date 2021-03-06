if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInAuthorizeUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInAuthorizeUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInCheckOut]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInCheckOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInCheckOutSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInCheckOutSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInMergeTemp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInMergeTemp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInOutLet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInOutLet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInSlipMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInSlipMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_DriveInSlipSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_DriveInSlipSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_PickingQueueRequest]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_PickingQueueRequest]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_PickingRequestMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_PickingRequestMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_PickingRequestSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_PickingRequestSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QuePickCenterMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QuePickCenterMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QuePickCenterSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QuePickCenterSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QueueRequestPicking]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QueueRequestPicking]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QueueRequestPicking2008]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QueueRequestPicking2008]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QueueRequestPickingMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QueueRequestPickingMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_NP_QueueRequestPickingMaster2008]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_NP_QueueRequestPickingMaster2008]
GO

CREATE TABLE [dbo].[TB_NP_DriveInAuthorizeUser] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[Code] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[Name1] [varchar] (150) COLLATE Thai_CI_AS NULL ,
	[DutyCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[LevelID] [smallint] NULL ,
	[ActiveStatus] [smallint] NULL ,
	[CreatorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[LastEditorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[LastEditDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInCheckOut] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[Checker] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[PosNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[NetDebtAmount] [money] NULL ,
	[IsCancel] [smallint] NULL ,
	[IsConfirm] [smallint] NULL ,
	[CreatorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[CancelCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CancelDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInCheckOutSub] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[QTY] [money] NULL ,
	[BillQTY] [money] NULL ,
	[UnitCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[Amount] [money] NULL ,
	[BarCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsCancel] [smallint] NULL ,
	[LineNumber] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInMergeTemp] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[WHCode] [char] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfCode] [char] (10) COLLATE Thai_CI_AS NULL ,
	[QTY] [money] NULL ,
	[UnitCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[DisCountAmount] [money] NULL ,
	[Amount] [money] NULL ,
	[BarCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[RefNo] [varchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[QueID] [int] NULL ,
	[PosBillNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsConfirm] [smallint] NULL ,
	[ConfirmDateTime] [datetime] NULL ,
	[LineNumber] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInOutLet] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[ID] [int] NULL ,
	[ItemName] [varchar] (200) COLLATE Thai_CI_AS NULL ,
	[ActiveStatus] [smallint] NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInSlipMaster] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ArCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[SaleCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[MemberID] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[RefNo] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[PickZone] [varchar] (2) COLLATE Thai_CI_AS NULL ,
	[BeforeTaxAmount] [money] NULL ,
	[TaxAmount] [money] NULL ,
	[TotalNetAmount] [money] NULL ,
	[MergeNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsMerge] [smallint] NULL ,
	[IsConfirm] [smallint] NULL ,
	[IsCancel] [smallint] NULL ,
	[IsSendQue] [smallint] NULL ,
	[CreatorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[LastEditorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[LastEditDateTime] [datetime] NULL ,
	[ConfirmCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[ConfirmDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_DriveInSlipSub] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[ItemName] [varchar] (250) COLLATE Thai_CI_AS NOT NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfID] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[QTY] [money] NULL ,
	[UnitCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[DisCountWord] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[DisCountAmount] [money] NULL ,
	[Amount] [money] NULL ,
	[IsCancel] [smallint] NULL ,
	[BarCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[LineNumber] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_PickingQueueRequest] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[SaleOrderNo] [nvarchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[SaleOrderDate] [nvarchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NULL ,
	[ARCode] [nvarchar] (20) COLLATE Thai_CI_AS NULL ,
	[SaleCode] [nvarchar] (20) COLLATE Thai_CI_AS NULL ,
	[RequestDate] [datetime] NOT NULL ,
	[RequestTime] [nvarchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[RequestStatus] [smallint] NULL ,
	[RequestCountItem] [money] NULL ,
	[RequestCountQTY] [money] NULL ,
	[PrintStatus] [smallint] NULL ,
	[RequestDateHistory] [datetime] NULL ,
	[RequestTimeHistory] [nvarchar] (20) COLLATE Thai_CI_AS NULL ,
	[RequestFromPerson] [nvarchar] (50) COLLATE Thai_CI_AS NULL ,
	[LastEditRequestFrom] [nvarchar] (50) COLLATE Thai_CI_AS NULL ,
	[RequestAt] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_PickingRequestMaster] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ARCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[SaleCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[RefNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[MemberID] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[BeforeTaxAmount] [money] NULL ,
	[TaxAmount] [money] NULL ,
	[NetDebtAmount] [money] NULL ,
	[IsConditionSend] [smallint] NULL ,
	[ReqTime] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsCancel] [smallint] NULL ,
	[IsSendQue] [smallint] NULL ,
	[MyDescription] [varchar] (300) COLLATE Thai_CI_AS NULL ,
	[CreatorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[LastEditorCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[LastEditDateTime] [datetime] NULL ,
	[CancelCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CancelDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_PickingRequestSub] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[QTY] [money] NULL ,
	[UnitCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[DisCountWord] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[DisCountAmount] [money] NULL ,
	[NetAmount] [money] NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfCode] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfID] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[BarCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsCancel] [smallint] NULL ,
	[LineNUmber] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QuePickCenterMaster] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[QueID] [int] NOT NULL ,
	[QueDocDate] [datetime] NOT NULL ,
	[DocNo] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NULL ,
	[ARCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[SaleCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[RefNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[MemberID] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[SourceID] [smallint] NULL ,
	[CashierCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[MergeNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[HoldBillNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[IsConfirm] [smallint] NULL ,
	[IsConditionSend] [smallint] NULL ,
	[Checker] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[CheckOutDateTime] [datetime] NULL ,
	[QueDescription] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[QueZone] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[QueDate] [datetime] NULL ,
	[QuePicker] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[QueStart] [datetime] NULL ,
	[QueStop] [datetime] NULL ,
	[QueStatus] [smallint] NULL ,
	[QuePickStatus] [smallint] NULL ,
	[QueReqTime] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[QueReceived] [smallint] NULL ,
	[QueRecStatus] [smallint] NULL ,
	[QueRecDate] [datetime] NULL ,
	[QueReason] [int] NULL ,
	[QueReasonDesc] [varchar] (300) COLLATE Thai_CI_AS NULL ,
	[QueTime] [int] NOT NULL ,
	[IsCancel] [smallint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QuePickCenterSub] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[QueID] [int] NOT NULL ,
	[QueDocDate] [datetime] NOT NULL ,
	[ItemCode] [varchar] (30) COLLATE Thai_CI_AS NOT NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ShelfCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ShelfID] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[QTY] [money] NULL ,
	[PickQTY] [money] NULL ,
	[OnCarQTY] [money] NULL ,
	[CheckQTY] [money] NULL ,
	[InvQTY] [money] NULL ,
	[Unitcode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[BarCode] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[RefNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[MergeNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[HoldBillNo] [varchar] (30) COLLATE Thai_CI_AS NULL ,
	[QueTime] [int] NOT NULL ,
	[LineNumber] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QueueRequestPicking] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NULL ,
	[DatePicking] [datetime] NOT NULL ,
	[ItemCode] [varchar] (25) COLLATE Thai_CI_AS NOT NULL ,
	[ItemName] [varchar] (200) COLLATE Thai_CI_AS NULL ,
	[ReqQTY] [money] NOT NULL ,
	[UnitCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ShelfCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[DiscountAmount] [money] NULL ,
	[ItemAmount] [money] NULL ,
	[IsCancel] [smallint] NULL ,
	[SelectItemDateTime] [datetime] NULL ,
	[SOCountNumber] [int] NOT NULL ,
	[LineNumber] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QueueRequestPicking2008] (
	[RowOrder] [int] NOT NULL ,
	[DocNo] [varchar] (20) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NULL ,
	[DatePicking] [datetime] NOT NULL ,
	[ItemCode] [varchar] (25) COLLATE Thai_CI_AS NOT NULL ,
	[ItemName] [varchar] (200) COLLATE Thai_CI_AS NULL ,
	[ReqQTY] [money] NOT NULL ,
	[UnitCode] [varchar] (20) COLLATE Thai_CI_AS NULL ,
	[WHCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ShelfCode] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[Price] [money] NULL ,
	[DiscountAmount] [money] NULL ,
	[ItemAmount] [money] NULL ,
	[IsCancel] [smallint] NULL ,
	[SelectItemDateTime] [datetime] NULL ,
	[SOCountNumber] [int] NOT NULL ,
	[LineNumber] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QueueRequestPickingMaster] (
	[RowOrder] [int] IDENTITY (1, 1) NOT NULL ,
	[DocNo] [varchar] (50) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ARCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[DatePicking] [datetime] NOT NULL ,
	[BillType] [smallint] NULL ,
	[SOStatus] [smallint] NULL ,
	[IsCancel] [smallint] NULL ,
	[SaleCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[CarLicense] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[IsConditionSend] [smallint] NULL ,
	[SOCountNumber] [int] NOT NULL ,
	[QueueNo] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfGroup] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[DueDate] [datetime] NULL ,
	[PickStatus] [smallint] NULL ,
	[SumOfItemAmount] [money] NULL ,
	[TaxAmount] [money] NULL ,
	[NetAmount] [money] NULL ,
	[CreatorCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[LastEditorCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[LastEditDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TB_NP_QueueRequestPickingMaster2008] (
	[RowOrder] [int] NOT NULL ,
	[DocNo] [varchar] (50) COLLATE Thai_CI_AS NOT NULL ,
	[DocDate] [datetime] NOT NULL ,
	[ARCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[DatePicking] [datetime] NOT NULL ,
	[BillType] [smallint] NULL ,
	[SOStatus] [smallint] NULL ,
	[IsCancel] [smallint] NULL ,
	[SaleCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[CarLicense] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[IsConditionSend] [smallint] NULL ,
	[SOCountNumber] [int] NOT NULL ,
	[QueueNo] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[ShelfGroup] [varchar] (10) COLLATE Thai_CI_AS NOT NULL ,
	[ZoneID] [varchar] (10) COLLATE Thai_CI_AS NULL ,
	[DueDate] [datetime] NULL ,
	[PickStatus] [smallint] NULL ,
	[SumOfItemAmount] [money] NULL ,
	[TaxAmount] [money] NULL ,
	[NetAmount] [money] NULL ,
	[CreatorCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[CreateDateTime] [datetime] NULL ,
	[LastEditorCode] [varchar] (50) COLLATE Thai_CI_AS NULL ,
	[LastEditDateTime] [datetime] NULL 
) ON [PRIMARY]
GO

