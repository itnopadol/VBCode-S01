if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_AR_ARProFileSearch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_AR_ARProFileSearch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_AR_ARProfile]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_AR_ARProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_AR_SearchMemberID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_AR_SearchMemberID]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_CRM_EmployeeDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_CRM_EmployeeDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_MB_ScanBarCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_MB_ScanBarCode]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_AR_ARProFileSearch
@ARCode as nvarchar(20)
as

set		dateformat dmy

if		@ARCode = ''
begin
		select	top 100 a.createdatetime,a.code as arcode,isnull(a.name1,'') as arname,isnull(a.billaddress,'') as address,isnull(a.telephone,'-') as telephone,
				isnull(debtlimit1,0) as debtlimit1,isnull(groupcode,'') as groupcode,isnull(b.name,'') as groupname,isnull(groupofdebt,'') as groupofdebt,
				isnull(c.name,'') as debtname,isnull(creditmencode,'') as salecode,isnull(e.name,'') as salename,isnull(a.pressmencode,'') as pressmencode,
				isnull(f.name,'')as pressmenname,isnull(memberid,'') as memberid
		from	dbo.bcar a 
		left	join dbo.bcargroup b on a.groupcode = b.code
		left	join dbo.bcardebtgroup c on a.groupofdebt = c.code
		left	join dbo.bccusttype d on a.typecode = d.code
		left	join dbo.bcsale e on a.creditmencode = e.code
		left	join dbo.bcsale f on a.pressmencode = f.code
		where	a.activestatus = 1 
end

if		@ARCode <> ''
begin
select	*
from
(
select	a.createdatetime,a.code as arcode,isnull(a.name1,'') as arname,isnull(a.billaddress,'') as address,isnull(a.telephone,'-') as telephone,
		isnull(debtlimit1,0) as debtlimit1,isnull(groupcode,'') as groupcode,isnull(b.name,'') as groupname,isnull(groupofdebt,'') as groupofdebt,
		isnull(c.name,'') as debtname,isnull(creditmencode,'') as salecode,isnull(e.name,'') as salename,isnull(a.pressmencode,'') as pressmencode,
		isnull(f.name,'')as pressmenname,isnull(memberid,'') as memberid
from	dbo.bcar a 
left	join dbo.bcargroup b on a.groupcode = b.code
left	join dbo.bcardebtgroup c on a.groupofdebt = c.code
left	join dbo.bccusttype d on a.typecode = d.code
left	join dbo.bcsale e on a.creditmencode = e.code
left	join dbo.bcsale f on a.pressmencode = f.code
where	a.activestatus = 1 and a.code like '%'+@ARCode+'%'
union
select	a.createdatetime,a.code as arcode,isnull(a.name1,'') as arname,isnull(a.billaddress,'') as address,isnull(a.telephone,'-') as telephone,
		isnull(debtlimit1,0) as debtlimit1,isnull(groupcode,'') as groupcode,isnull(b.name,'') as groupname,isnull(groupofdebt,'') as groupofdebt,
		isnull(c.name,'') as debtname,isnull(creditmencode,'') as salecode,isnull(e.name,'') as salename,isnull(a.pressmencode,'') as pressmencode,
		isnull(f.name,'')as pressmenname,isnull(memberid,'') as memberid
from	dbo.bcar a 
left	join dbo.bcargroup b on a.groupcode = b.code
left	join dbo.bcardebtgroup c on a.groupofdebt = c.code
left	join dbo.bccusttype d on a.typecode = d.code
left	join dbo.bcsale e on a.creditmencode = e.code
left	join dbo.bcsale f on a.pressmencode = f.code
where	a.activestatus = 1 and a.name1 like '%'+@ARCode+'%'
) as	result
order	by arcode
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

CREATE	procedure dbo.USP_AR_ARProfile
@ARCode as nvarchar(20)
as

set		dateformat dmy
select	a.createdatetime,a.code as arcode,isnull(a.name1,'') as arname,isnull(a.billaddress,'') as address,isnull(a.telephone,'-') as telephone,
		isnull(debtlimit1,0) as debtlimit1,isnull(groupcode,'') as groupcode,isnull(b.name,'') as groupname,isnull(groupofdebt,'') as groupofdebt,
		isnull(c.name,'') as debtname,isnull(creditmencode,'') as salecode,isnull(e.name,'') as salename,isnull(a.pressmencode,'') as pressmencode,
		isnull(f.name,'')as pressmenname,isnull(memberid,'') as memberid,
		isnull((select top 1 name from dbo.bccontactlist where parentcode = a.code order by code),'') as contactname1,
		isnull((select top 1 telephone from dbo.bccontactlist where parentcode = a.code order by code),'') as contactPhone1,
		isnull((select top 1 name from dbo.bccontactlist where parentcode = a.code and code not in (select top 1 code from dbo.bccontactlist where parentcode = @ARCode  order by code)order by code),'') as contactname2,
		isnull((select top 1 telephone from dbo.bccontactlist where parentcode = a.code and code not in (select top 1 code from dbo.bccontactlist where parentcode = @ARCode  order by code)order by code),'') as contactPhone2,
		condpaycode,pricelevel,
		case pricelevel
		when 1 then 'ระดับราคาที่ 1'
		when 2 then 'ระดับราคาที่ 2'
		when 3 then 'ระดับราคาที่ 3'
		when 4 then 'ระดับราคาที่ 4'
		end as pricelevel1
from	dbo.bcar a 
left	join dbo.bcargroup b on a.groupcode = b.code
left	join dbo.bcardebtgroup c on a.groupofdebt = c.code
left	join dbo.bccusttype d on a.typecode = d.code
left	join dbo.bcsale e on a.creditmencode = e.code
left	join dbo.bcsale f on a.pressmencode = f.code
where	a.activestatus = 1 and a.code = @ARCode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE	procedure dbo.USP_AR_SearchMemberID
@vMemberID as nvarchar(30)
as

set		dateformat dmy

select	a.code as arcode,isnull(a.name1,'') as arname,isnull(memberid,'') as memberid
from	dbo.bcar a 
where	a.activestatus = 1 and a.memberid = @vMemberID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE dbo.USP_CRM_EmployeeDetails
@vType as int, 
@Employee as varchar(50)

as

if	@vType = 0
begin
	if @Employee = ''
	begin
		select	Code as EmpCode,name as EmpName 
		from	dbo.BCSale 
		where	 ActiveStatus=1
		order	by code
	end
	if @Employee <> ''
	begin
		select	*
		from
		(
		select	Code as EmpCode,name as EmpName 
		from	dbo.BCSale 
		where	code like '%'+@Employee +'%'  and ActiveStatus=1
		union
		select	Code as EmpCode,name as EmpName 
		from	dbo.BCSale 
		where	name like '%'+@Employee +'%'  and ActiveStatus=1
		) as 	result
		order	by EmpCode
	end 
end


if	@vType = 1
begin
select	Code as EmpCode,name as EmpName 
from	dbo.BCSale 
where	code = @Employee and ActiveStatus=1 
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

CREATE	procedure dbo.USP_MB_ScanBarCode
@vBarCode as nvarchar(20)
as
select  a.itemcode,a.barcode,isnull(b.name1,'') as itemname,isnull(b.defsaleunitcode,'') as unitcode
from    dbo.bcbarcodemaster a
		left join dbo.bcitem b on a.itemcode = b.code 
where   a.barcode = @vBarCode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

