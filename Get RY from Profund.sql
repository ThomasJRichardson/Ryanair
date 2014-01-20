use APTTest
go
	select	
		left(P.Reference,4) + LOWER(SUBSTRING(P.Reference,5,15)) as Username,
		P.PersonUID,
		P.FamiliarName as ActiveDeferred,
		P.Salutation as PPSN,
		P.Surname,
		P.Forename,
		A.Line1 as Address1,
		A.Line2 as Address2,
		A.Line3 as Address3,
		A.Line4 as Address4,
		A.TelephoneNumber as PhoneHome,
		Mobile.Value as PhoneMobile,
		Email.Value as Email,
		P.DateOfBirth,
		P.Sex as Gender,
		case MAR.Value	when 'APA' then 'Separated'
				when 'DIV' then 'Divorced'
				when 'MAR' then 'Married'
				when 'SIN' then 'Single'
				else 'Unknown'
		end as MaritalStatus,

		EE.DateFirstEmployed as DateEmpStart,
		EEX.EffDate as DateEmpCease,
		SM1.DateJoinedScheme as DateJoinedScheme,
		isnull(JC1.LongDesc,'?') as SchemeCategory,
		SAL1.Value as FinalSalary,
		SAL2.Value as PensionableSalary,
		SAL3.Value as SchemeSalary,
		left(EE.PayrollNumber,1) as TransferIn,
		SM1.SchemeRetirementDate as NormalRetDate,

		P.PrevSurname as ServiceRecordVerified

	from APTTEST.dbo.Person P

	left outer join APTTEST.dbo.Communications Email on Email.ParentUID = P.PersonUID and Email.Catid = 802 and Email.EndDate is null
	left outer join APTTEST.dbo.Communications Mobile on Mobile.ParentUID = P.PersonUID and Mobile.Catid = 800 and Mobile.EndDate is null

	left outer join APTTEST.dbo.StringHistory MAR on MAR.ParentUID = P.PersonUID and MAR.CatID = 1200

	inner join APTTEST.dbo.Employee EE on EE.PersonUID = P.PersonUID
		and EE.DateFirstEmployed = (select min(DateFirstEmployed) from APTTEST.dbo.Employee where PersonUID = P.PersonUID)

	left outer join APTTEST.dbo.Address A on A.ParentUID = P.PersonUID and A.CatId = 1 and A.EndDate is null
	
	left outer join APTTEST.dbo.IntegerHistory EEX on EEX.ParentUID = EE.EmployeeUID and EEX.CatID = 4137 and EEX.Value = 4119

	left outer join APTTEST.dbo.IntegerHistory EEJC on EEJC.ParentUID = EE.EmployeeUID and EEJC.CatID = 1002

	left outer join APTTEST.dbo.CurrencyHistory SAL1 on SAL1.ParentUID = EE.EmployeeUID and SAL1.CatID = 201
	left outer join APTTEST.dbo.CurrencyHistory SAL2 on SAL2.ParentUID = EE.EmployeeUID and SAL2.CatID = 203
	left outer join APTTEST.dbo.CurrencyHistory SAL3 on SAL3.ParentUID = EE.EmployeeUID and SAL3.CatID = 25143

	left outer join APTTEST.dbo.JobClass JC1 on JC1.JobClassUID = EEJC.Value

	inner join APTTEST.dbo.SchemeMember SM1 on SM1.EmployeeUID = EE.EmployeeUID

	where P.Reference like 'RY-%'
	order by P.Reference;

go

use TomR
go
	delete from TomR.dbo.IP_WEB
	where username like 'RY-%'
go

		insert into Tomr.[dbo].[IP_WEB] (
		[Username],
		[PersonUID],
		[ServiceStatus],

		[PPSN],
		[Surname],
		[Forename],

		[Address1],
		[Address2],
		[Address3],
		[Address4],

		[PhoneHome],
		[PhoneMobile],
		[Email],

		[DateOfBirth],
		[Gender],
		[MaritalStatus],

		[ProfileLastUpdateBy],
		[ProfileLastUpdateAt],

		[DateEmpStart_1],
		[DateEmpCease_1],
		[DateJoinedScheme_1],
		[SchemeCategory_1],
		[FinalSalary_1],
		[PensionableSalary_1],
		[SchemeSalary_1],
		[TransferIn_1],
		[NormalRetDate_1],

		[ServiceLastUpdateAt],

		[DateEmpStart_m1],
		[DateEmpCease_m1],
		[DateJoinedScheme_m1],
		[SchemeCategory_m1],
		[FinalSalary_m1],
		[PensionableSalary_m1],
		[SchemeSalary_m1],
		[TransferIn_m1],
		[NormalRetDate_m1]
	)
	select	
		left(P.Reference,4) + LOWER(SUBSTRING(P.Reference,5,15)) as Username,
		P.PersonUID,
		isnull(P.FamiliarName,'?'),
		
		P.Salutation as PPSN,
		P.Surname,
		P.Forename,
		A.Line1 as Address1,
		A.Line2 as Address2,
		A.Line3 as Address3,
		A.Line4 as Address4,
		A.TelephoneNumber as PhoneHome,
		Mobile.Value as PhoneMobile,
		Email.Value as Email,
		P.DateOfBirth,
		P.Sex as Gender,
		case MAR.Value	when 'APA' then 'Separated'
				when 'DIV' then 'Divorced'
				when 'MAR' then 'Married'
				when 'SIN' then 'Single'
				else 'Unknown'
		end as MaritalStatus,

		'APT',
		getdate(),
		
		EE.DateFirstEmployed as DateEmpStart,
		EEX.EffDate as DateEmpCease,
		SM1.DateJoinedScheme as DateJoinedScheme,
		isnull(JC1.LongDesc,'?') as SchemeCategory,
		SAL1.Value as FinalSalary,
		SAL2.Value as PensionableSalary,
		SAL3.Value as SchemeSalary,
		left(EE.PayrollNumber,1) as TransferIn,
		SM1.SchemeRetirementDate as NormalRetDate,

		getdate(),

		EE.DateFirstEmployed as DateEmpStart,
		EEX.EffDate as DateEmpCease,
		SM1.DateJoinedScheme as DateJoinedScheme,
		isnull(JC1.LongDesc,'?') as SchemeCategory,
		SAL1.Value as FinalSalary,
		SAL2.Value as PensionableSalary,
		SAL3.Value as SchemeSalary,
		left(EE.PayrollNumber,1) as TransferIn,
		SM1.SchemeRetirementDate as NormalRetDate
		
	from APTTEST.dbo.Person P

	left outer join APTTEST.dbo.Communications Email on Email.ParentUID = P.PersonUID and Email.Catid = 802 and Email.EndDate is null
	left outer join APTTEST.dbo.Communications Mobile on Mobile.ParentUID = P.PersonUID and Mobile.Catid = 800 and Mobile.EndDate is null

	left outer join APTTEST.dbo.StringHistory MAR on MAR.ParentUID = P.PersonUID and MAR.CatID = 1200

	inner join APTTEST.dbo.Employee EE on EE.PersonUID = P.PersonUID
		and EE.DateFirstEmployed = (select min(DateFirstEmployed) from APTTEST.dbo.Employee where PersonUID = P.PersonUID)

	left outer join APTTEST.dbo.Address A on A.ParentUID = P.PersonUID and A.CatId = 1 and A.EndDate is null
	
	left outer join APTTEST.dbo.IntegerHistory EEX on EEX.ParentUID = EE.EmployeeUID and EEX.CatID = 4137 and EEX.Value = 4119

	left outer join APTTEST.dbo.IntegerHistory EEJC on EEJC.ParentUID = EE.EmployeeUID and EEJC.CatID = 1002

	left outer join APTTEST.dbo.CurrencyHistory SAL1 on SAL1.ParentUID = EE.EmployeeUID and SAL1.CatID = 201
	left outer join APTTEST.dbo.CurrencyHistory SAL2 on SAL2.ParentUID = EE.EmployeeUID and SAL2.CatID = 203
	left outer join APTTEST.dbo.CurrencyHistory SAL3 on SAL3.ParentUID = EE.EmployeeUID and SAL3.CatID = 25143

	left outer join APTTEST.dbo.JobClass JC1 on JC1.JobClassUID = EEJC.Value

	inner join APTTEST.dbo.SchemeMember SM1 on SM1.EmployeeUID = EE.EmployeeUID

	where P.Reference like 'RY-%';
