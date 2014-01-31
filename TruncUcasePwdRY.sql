use APT2012
go

update M set M.password = UPPER(left(M.password,10)),
PasswordSalt = UPPER(left(M.password,10))
from aspnet_Users U inner join aspnet_Membership M
on M.UserId = U.UserId
where U.UserName like 'RY-%'
and M.password <> UPPER(left(M.password,10))

delete from dbo.ASP_PASSWORD_UPDATES where UserName like 'RY-%'
delete from dbo.ASP_PASSWORD_UPDATE_LOG where UserName like 'RY-%'
