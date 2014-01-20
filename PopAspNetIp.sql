use APT2012
go

--BEGIN TRAN

insert into APT2012.dbo.aspnet_Users (
	ApplicationId,
	UserName,
	LoweredUserName,
	IsAnonymous,
	LastActivityDate )
select
	'7A1A0D49-5478-4058-BF29-E4E9DA3B4F53',
	username,
	LOWER(username),
	0,
	GETDATE()
from APT2012.dbo.IP

--select * from APT2012.dbo.aspnet_Users
--where UserName like 'RY-%'

insert into APT2012.dbo.aspnet_Membership (
	ApplicationId,
	UserId,
	Password,
	PasswordFormat,
	PasswordSalt,
	IsApproved,
	IsLockedOut,
	CreateDate,
	LastLoginDate,
	LastPasswordChangedDate,
	LastLockoutDate,
	FailedPasswordAnswerAttemptCount,
	FailedPasswordAnswerAttemptWindowStart,
	FailedPasswordAttemptCount,
	FailedPasswordAttemptWindowStart
)
select
	'7A1A0D49-5478-4058-BF29-E4E9DA3B4F53',
	U.UserId,
	I.Password,
	0,
	I.Password,
	1,
	0,
	GETDATE(),
	GETDATE(),
	GETDATE(),
	'1754-01-01',
	0,
	'1754-01-01',
	0,
	'1754-01-01'
from APT2012.dbo.IP I inner join aspnet_Users U on U.UserName = I.username

--select M.* from APT2012.dbo.aspnet_Membership M
--inner join aspnet_Users U on U.UserId = M.UserId
--where U.UserName like 'RY-%'



--select P.* from APT2012.dbo.aspnet_Profile P
--inner join aspnet_Users U on U.UserId = P.UserId
--where U.UserName like 'RY-%'

insert into aspnet_usersinRoles (
userid,
RoleId
)
select userid,'461A8413-7FB6-41A3-920F-D2CE845B1F33'	--Member
from aspnet_Users where UserName like 'RY-%'
--ROLLBACK TRAN

--BEGIN TRAN
update P
set		P.LastName = X.Surname,
		P.FirstName = X.Forename,
		P.Address1 = X.Address1,
		P.Address2 = X.Address2,
		P.Address3 = X.Address3,
		P.Address4 = X.Address4,
		P.PhoneHome = X.PhoneHome,
		P.PhoneMobile = X.PhoneMobile
from APT2012.dbo.aspnet_Profile P
inner join APT2012.dbo.aspnet_Users U on U.UserId = P.UserId
inner join
(
select	Username,
		Surname,
		Forename,
		Address1,
		Address2,
		Address3,
		Address4,
		PhoneHome,
		PhoneMobile
from IPPP.dbo.IP_WEB W
where username like 'RY-%'

except

select	U.UserName,
		P.LastName,
		P.FirstName,
		P.Address1,
		P.Address2,
		P.Address3,
		P.Address4,
		P.PhoneHome,
		P.PhoneMobile
from APT2012.dbo.aspnet_Profile P
inner join APT2012.dbo.aspnet_Users U on U.UserId = P.UserId
where U.username like 'RY-%'
) X
on X.Username = U.UserName

update M
set		M.Email = X.Email
from APT2012.dbo.aspnet_Membership M
inner join APT2012.dbo.aspnet_Users U on U.UserId = M.UserId
inner join
(
select	Username,
		Email
from IPPP.dbo.IP_WEB W
where username like 'RY-%'

except

select	U.UserName,
		M.Email
from APT2012.dbo.aspnet_Membership M
inner join APT2012.dbo.aspnet_Users U on U.UserId = M.UserId
where U.username like 'RY-%'
) X
on X.Username = U.UserName

delete from APT2012.dbo.ASP_PROFILE_UPDATES
where UserId in (
select UserId from APT2012.dbo.aspnet_Users where UserName like 'RY-%'
)

delete from APT2012.dbo.ASP_PROFILE_UPDATE_LOG
where UserId in (
select UserId from APT2012.dbo.aspnet_Users where UserName like 'RY-%'
)
--ROLLBACK TRAN

insert into Member (
	[PPSN],
	[PersonalStatus],
	[PersonalStatusDate],
	[DateOfBirth],
	[Forename],
	[Surname],
	[Sex],
	[PersonUID]
)
select	username,
	maritalstatus,
	getdate(),
	dateofbirth,
	forename,
	surname,
	gender,
	personuid
from IPPP.dbo.IP_WEB
where username like 'RY-%'

insert into MEmberScheme (
	MemberId,
	SchemeId,
	DateJoined
)
select	M.ID,
	1009817,
	getdate()
from	MEmber M
where PPSN like 'RY-%'

insert into APT2012.dbo.aspnet_Profile (
	UserId,
	FirstName,
	LastName,
	MemberId,
	SchemeId,
	LastUpdatedDate
)
select
	U.UserId,
	W.Forename,
	W.Surname,
	M.Id,
	MS.SchemeId,
	GETDATE()
from APT2012.dbo.IP I inner join aspnet_Users U on U.UserName = I.username
inner join IPPP.dbo.IP_WEB W on W.Username = I.username
inner join APT2012.dbo.Member M on M.PPSN = I.username
inner join APT2012.dbo.MemberScheme MS on MS.MemberId = M.id
