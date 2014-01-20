USE [APT2012]
GO

/****** Object:  StoredProcedure [dbo].[spMakeIpPasswords]    Script Date: 12/12/2013 10:56:59 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spMakeIpPasswords]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spMakeIpPasswords]
GO

USE [APT2012]
GO

/****** Object:  StoredProcedure [dbo].[spMakeIpPasswords]    Script Date: 12/12/2013 10:56:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create procedure [dbo].[spMakeIpPasswords]
AS


declare @Length integer, @CharPool varchar(100), @LoopCount integer, @RandomString varchar(40),
@PoolLength integer

declare @ppsn varchar(20)

SET @CharPool = 
    'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
    
SET @PoolLength = DataLength(@CharPool)

delete from APT2012.dbo.IP

insert into APT2012.dbo.IP (Username)
select distinct username from IPPP.dbo.IP_WEB
where username like 'RY-%'

select @ppsn = MIN(username) from APT2012.dbo.IP

while @ppsn is not null
begin
	SET @Length = RAND() * 4 + 8    

	SET @LoopCount = 0
	SET @RandomString = ''

	WHILE (@LoopCount < @Length) BEGIN
		SELECT @RandomString = @RandomString + 
			SUBSTRING(@Charpool, CONVERT(int, RAND() * @PoolLength), 1)
		SELECT @LoopCount = @LoopCount + 1
	END

	update APT2012.dbo.IP
	set [Password] = @RandomString
	where username = @ppsn
	
	select @ppsn = MIN(username) from APT2012.dbo.IP where username > @ppsn
end

select * from APT2012.dbo.IP order by username
GO


