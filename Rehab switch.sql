select
P.Surname, P.Forename, P.NationalIDNumber, MI.MemberInvestmentUID,
MB.MemberBenefitUID as ParentUID,
0 as ContributionTypeCatID, 0 as ContributionReasonCatID, --?
GETDATE() as RequestDate,
I.InvestmentUID, --?
'M' as RequestLevel,
0 as SwitchAmount --?

from Person P
inner join Employee E on E.PersonUID = P.PersonUID
inner join SchemeMember SM on SM.EmployeeUID = E.EmployeeUID
inner join Scheme S on S.SchemeUId = SM.SchemeUID
inner join MemberBenefit MB on MB.MemberUID = SM.MemberUID
inner join MemberInvestment MI on MI.MemberBenefitUID = MB.MemberBenefitUID
inner join Investment I on I.InvestmentUID = MI.InvestmentUID
inner join dbo.vwMemberInvestmentCurrentFundValue V
on V.MemberBenefitUID = MB.MemberBenefitUID and V.InvestmentUID = I.InvestmentUID

where S.SchemeUId = 1279397
and I.InvestmentUID = 1274637
and V.Units <> 0

order by P.surname

/*
select * from dbo.vwMemberInvestmentDetails
where MemberInvestmentUID = 78409994

select * from dbo.vwMemberInvestmentCurrentFundValue
where MemberBenefitUID = 1284693
and InvestmentUID = 1274637
*/

--select * from Scheme where ShortDesc like 'Rehab%'

select * from SwitchRequestHistory
where ParentUID = 1284693

select * from SwitchProcessHistory
where ParentUID = 1284693
