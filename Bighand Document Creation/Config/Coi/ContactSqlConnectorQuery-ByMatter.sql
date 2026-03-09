SELECT DISTINCT TOP 200
       'M' + CAST(mc.ID AS nvarchar(50)) AS Id, 
       mc.EntityRef as ClientNumber,
       mc.MatterNo as MatterNumber,
       ISNULL(RTRIM(LTRIM(ent.LegalName)), '(blank)') as [Name],
       ISNULL(ent.Title,'') as Title, 
       ISNULL(ent.Forename, '') as Forename,
       ISNULL(ent.MiddleName, '') as MiddleName,
       ISNULL(ent.Surname, '')  as Surname,
       '' as Suffix,
       ISNULL(ent.Initials, '') as Initials,       
       ISNULL(ent.JobDescription, '') as JobTitle,  
       '' as Department, 
       ISNULL(addr.Country, '') as Country,
       ISNULL(ent.[Name] ,'') as Company,
       REPLACE(
              ISNULL(addr.Street1, '') + CHAR(13) + CHAR(10) + 
              ISNULL(addr.Street2, '') + CHAR(13) + CHAR(10) + 
              ISNULL(addr.Town, '') + CHAR(13) + CHAR(10) + 
              ISNULL(addr.County, '') + CHAR(13) + CHAR(10)  + 
              ISNULL(addr.Postcode, '') + CHAR(13) + CHAR(10) + 
              ISNULL(addr.Country, ''), 
              CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10), 
              CHAR(13) + CHAR(10)
       ) AS [Address],      
       ISNULL(mc.Salutation, '') as SalutationList,
       ISNULL(ent.EnvelopeName, '') as AddressList,
       ISNULL(tel.Number,'') as Phone, 
       ISNULL(eml.[Address],'') as Email, 
       ISNULL(fax.Number, '') as Fax, 
       ISNULL(ent.Tel_Mob, '') as Mobile,
       ety.[Description] as RelationshipType,
       case          
              when mc.EntityRef = Left(@client, 1) + RIGHT('00000000000000' + SUBSTRING(@client, 2, 14), 14) AND mc.MatterNo=@matter  then 'Matter Contacts'
              when mc.EntityRef = Left(@client, 1) + RIGHT('00000000000000' + SUBSTRING(@client, 2, 14), 14) THEN 'Client Contacts'
              else 'Other Contacts'
       end AS [Owner],
       '' as AddressType,
       '' as Reference,
       '' as Office,
       '' as CustomField1, 
       '' as CustomField2 
FROM 
       std_MatterContacts mc
       LEFT JOIN Entities                   ent           ON mc.ContactRef = ent.Code
       LEFT JOIN std_Addresses              addr		  ON mc.AddressID = addr.ID 
       LEFT JOIN std_TelephoneNumbers		tel           ON mc.TelephoneID = tel.ID 
       LEFT JOIN std_TelephoneNumbers		fax           ON mc.FaxId = fax.ID 
       LEFT JOIN std_EmailAddresses			eml           ON mc.EmailID = eml.ID  
       LEFT JOIN EntityTypes                ety           ON mc.EntityTypeRef = ety.Code    
WHERE
       mc.EntityRef = Left(@client, 1) + RIGHT('00000000000000' + SUBSTRING(@client, 2, 14), 14) AND
	   mc.MatterNo = @matter AND
	   mc.AddressID > -1	   
UNION
SELECT DISTINCT TOP 1
       'M' + CAST(mc.ID AS nvarchar(50)) AS Id, 
       mc.EntityRef as ClientNumber,
       mc.MatterNo as MatterNumber,
       ISNULL(RTRIM(LTRIM(ent.LegalName)), '(blank)') as [Name],
       ISNULL(ent.Title,'') as Title, 
       ISNULL(ent.Forename, '') as Forename,
       ISNULL(ent.MiddleName, '') as MiddleName,
       ISNULL(ent.Surname, '')  as Surname,
       '' as Suffix,
       ISNULL(ent.Initials, '') as Initials,       
       ISNULL(ent.JobDescription, '') as JobTitle,  
       '' as Department, 
       ISNULL(addr.Country, '') as Country,
       ISNULL(ent.[Name] ,'') as Company,
       REPLACE(
              ISNULL(m.CorrAddress1, '') + CHAR(13) + CHAR(10) + 
              ISNULL(m.CorrAddress2, '') + CHAR(13) + CHAR(10) + 
              ISNULL(m.CorrAddress3, '') + CHAR(13) + CHAR(10) + 
              ISNULL(m.CorrAddress4, '') + CHAR(13) + CHAR(10)  + 
              ISNULL(m.CorrStreet3, '') + CHAR(13) + CHAR(10) + 
              ISNULL(m.CorrPostcode, '') + CHAR(13) + CHAR(10), 
              CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10), 
              CHAR(13) + CHAR(10)
       ) AS [Address],      
       ISNULL(mc.Salutation, '') as SalutationList,
       ISNULL(ent.EnvelopeName, '') as AddressList,
       ISNULL(tel.Number,'') as Phone, 
       ISNULL(eml.[Address],'') as Email, 
       ISNULL(fax.Number, '') as Fax, 
       ISNULL(ent.Tel_Mob, '') as Mobile,
       ety.[Description] as RelationshipType,
       'Matter Correspondence Address' AS [Owner],
       '' as AddressType,
       '' as Reference,
       '' as Office,
       '' as CustomField1, 
       '' as CustomField2 
FROM 
       std_MatterContacts mc
       LEFT JOIN Entities                   ent           ON mc.ContactRef = ent.Code
	   LEFT JOIN Matters					m			  ON mc.MatterNo = m.Number AND ent.Code = m.EntityRef 
       LEFT JOIN std_Addresses              addr		  ON mc.AddressID = addr.ID 
       LEFT JOIN std_TelephoneNumbers		tel           ON mc.TelephoneID = tel.ID 
       LEFT JOIN std_TelephoneNumbers		fax           ON mc.FaxId = fax.ID 
       LEFT JOIN std_EmailAddresses			eml           ON mc.EmailID = eml.ID  
       LEFT JOIN EntityTypes                ety           ON mc.EntityTypeRef = ety.Code    
WHERE
       mc.EntityRef = Left(@client, 1) + RIGHT('00000000000000' + SUBSTRING(@client, 2, 14), 14) AND 
	   mc.MatterNo = @matter
UNION
SELECT DISTINCT TOP 1
       'M' + CAST(mc.ID AS nvarchar(50)) AS Id, 
       mc.EntityRef as ClientNumber,
       mc.MatterNo as MatterNumber,
       ISNULL(RTRIM(LTRIM(ent.LegalName)), '(blank)') as [Name],
       ISNULL(ent.Title,'') as Title, 
       ISNULL(ent.Forename, '') as Forename,
       ISNULL(ent.MiddleName, '') as MiddleName,
       ISNULL(ent.Surname, '')  as Surname,
       '' as Suffix,
       ISNULL(ent.Initials, '') as Initials,       
       ISNULL(ent.JobDescription, '') as JobTitle,  
       '' as Department, 
       ISNULL(addr.Country, '') as Country,
       ISNULL(ent.[Name] ,'') as Company,
       REPLACE(
              ISNULL(ent.Address1, '') + CHAR(13) + CHAR(10) + 
              ISNULL(ent.Address2, '') + CHAR(13) + CHAR(10) + 
              ISNULL(ent.Address3, '') + CHAR(13) + CHAR(10) + 
              ISNULL(ent.Address4, '') + CHAR(13) + CHAR(10)  + 
              ISNULL(ent.Street3, '') + CHAR(13) + CHAR(10) + 
              ISNULL(ent.Postcode, '') + CHAR(13) + CHAR(10), 
              CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10), 
              CHAR(13) + CHAR(10)
       ) AS [Address],      
       ISNULL(mc.Salutation, '') as SalutationList,
       ISNULL(ent.EnvelopeName, '') as AddressList,
       ISNULL(tel.Number,'') as Phone, 
       ISNULL(eml.[Address],'') as Email, 
       ISNULL(fax.Number, '') as Fax, 
       ISNULL(ent.Tel_Mob, '') as Mobile,
       ety.[Description] as RelationshipType,
       'Client Main Address' AS [Owner],
       '' as AddressType,
       '' as Reference,
       '' as Office,
       '' as CustomField1, 
       '' as CustomField2 
FROM 
       std_MatterContacts mc
       LEFT JOIN Entities                   ent           ON mc.ContactRef = ent.Code
       LEFT JOIN std_Addresses              addr		  ON mc.AddressID = addr.ID 
       LEFT JOIN std_TelephoneNumbers		tel           ON mc.TelephoneID = tel.ID 
       LEFT JOIN std_TelephoneNumbers		fax           ON mc.FaxId = fax.ID 
       LEFT JOIN std_EmailAddresses			eml           ON mc.EmailID = eml.ID  
       LEFT JOIN EntityTypes                ety           ON mc.EntityTypeRef = ety.Code    
WHERE
       mc.EntityRef = Left(@client, 1) + RIGHT('00000000000000' + SUBSTRING(@client, 2, 14), 14) 