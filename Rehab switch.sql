use TomR
go

select
R.PPSN, P.Surname, P.Forename, 
MB.MemberBenefitUID,
ILIM.MemberInvestmentUID as Investment_UID_Out,
SSGA.MemberInvestmentUID as Investment_UID_In,
R.UnitsOut, R.UnitsIn,
C1.CatIDCode as TV, C2.CatIDCode as Reason,
ILIM_TV.Units as Total_ILIM_units, ILIM_V.Units as Cat_ILIM_units,

R.UnitsOut * (ILIM_V.Units / ILIM_TV.Units) as adj_units_out,
R.UnitsIn * (ILIM_V.Units / ILIM_TV.Units) as adj_units_in

from TomR.dbo.RehabILIM2SSGA_Oct13 R

inner join APTLIVE.dbo.Person P on P.NationalIDNumber = R.PPSN collate SQL_Latin1_General_CP1_CI_AS
inner join APTLIVE.dbo.Employee E on E.PersonUID = P.PersonUID
inner join APTLIVE.dbo.SchemeMember SM on SM.EmployeeUID = E.EmployeeUID
inner join APTLIVE.dbo.Scheme S on S.SchemeUId = SM.SchemeUID
inner join APTLIVE.dbo.MemberBenefit MB on MB.MemberUID = SM.MemberUID
inner join APTLIVE.dbo.MemberInvestment ILIM on ILIM.MemberBenefitUID = MB.MemberBenefitUID
inner join APTLIVE.dbo.Investment ILIM_I on ILIM_I.InvestmentUID = ILIM.InvestmentUID
inner join APTLIVE.dbo.MemberInvestment SSGA on SSGA.MemberBenefitUID = MB.MemberBenefitUID
inner join APTLIVE.dbo.Investment SSGA_I on SSGA_I.InvestmentUID = SSGA.InvestmentUID

inner join (
SELECT mi.MemberBenefitUID, mi.InvestmentUID, SUM(ih.units) Units
FROM APTLive.dbo.investmenthistory ih 
	JOIN APTLive.dbo.memberinvestment mi ON mi.memberinvestmentuid = ih.ParentInvestmentUID 
GROUP BY mi.MemberbenefitUID, mi.InvestmentUID
) ILIM_TV
	on ILIM_TV.MemberBenefitUID = MB.MemberBenefitUID
	and ILIM_TV.InvestmentUID = ILIM.InvestmentUID
	
inner join (
SELECT mi.MemberBenefitUID, mi.InvestmentUID, ih.ContributionReasonCatID, ih.ContributionTypeCatID, SUM(ih.units) Units
FROM APTLive.dbo.investmenthistory ih 
	JOIN APTLive.dbo.memberinvestment mi ON mi.memberinvestmentuid = ih.ParentInvestmentUID 
GROUP BY mi.MemberbenefitUID, mi.InvestmentUID, ih.ContributionReasonCatID, ih.ContributionTypeCatID
) ILIM_V
	on ILIM_V.MemberBenefitUID = MB.MemberBenefitUID
	and ILIM_V.InvestmentUID = ILIM.InvestmentUID
	
inner join APTLive.dbo.CategoryID C1 on C1.CatID = ILIM_V.ContributionReasonCatID
inner join APTLive.dbo.CategoryID C2 on C2.CatID = ILIM_V.ContributionTypeCatID

where S.SchemeUId = 1279397		--Rehab Group Defined Contribution Pension Scheme
and ILIM.InvestmentUID = 1274637	--Irish Life Pension Irish Property - PPI
and SSGA.InvestmentUID = 17372976	--SSgA EUT Global Consensus Fund S20 - 3135

order by P.surname