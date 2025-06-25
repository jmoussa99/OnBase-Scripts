param(
	[string]$OutputDir = "C:\Temp\DatabaseReports"
)

# Ensure ImportExcel module is installed and imported
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
	Write-Host "ImportExcel module not found. Installing..."
	Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module ImportExcel -Force

# Create output folder if it doesn't exist
if (-not (Test-Path $OutputDir)) {
	New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# Prompt for ODBC DSN
$dsn = Read-Host "Enter the ODBC DSN name"

$username = Read-Host "Enter the SQL username"
$password = Read-Host "Enter the SQL password" -AsSecureString
$plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
$connectionString = "DSN=$dsn;Uid=$username;Pwd=$plainPassword;"

# Function to run SQL and export to XLSX
function Run-Query {
	param(
		[string]$Query,
		[string]$OutputFile
	)
	try {
		$connection = New-Object System.Data.Odbc.OdbcConnection($connectionString)
		$connection.Open()
		$command = $connection.CreateCommand()
		$command.CommandText = $Query
		$adapter = New-Object System.Data.Odbc.OdbcDataAdapter($command)
		$dataSet = New-Object System.Data.DataSet
		$adapter.Fill($dataSet) | Out-Null
		$connection.Close()

		if ($dataSet.Tables.Count -gt 0) {
			$dataSet.Tables[0] | Export-Excel -Path $OutputFile -WorksheetName "Results" -AutoSize -AutoFilter
		}
		else {
			Write-Warning "No data returned for $OutputFile"
		}
	}
 catch {
		Write-Error "Failed to run query: $_"
	}
}

# Queries to run
$queries = @(
	@{
		Query  = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; select type, SUM((size*8)/1024) as MB from sys.database_files group by type"
		Output = "$OutputDir\Database_Size.xlsx"
	},

	@{
		Query  = "select 'LifeCycles' as 'Category', count(*) as 'Count' from hsi.lifecycle union all select 'Queues' as 'Category', count(*) as 'Count' from hsi.lcstate"
		Output = "$OutputDir\Overview_Counts.xlsx"
	},
	@{
		Query  = @"
select rtrim(ps.productname) as ProductName, max(lu.usagecount) as MaxUsage
FROM hsi.licusage lu
INNER JOIN hsi.productsold ps ON lu.producttype = ps.producttype
where lu.logdate > DATEADD(dd, -90, getdate())
GROUP BY ps.productname
ORDER BY ps.productname
"@
		Output = "$OutputDir\Peak_License_Usage.xlsx"
	},
	@{
		Query  = @"
select rtrim(p.productname) as product, l.licensecount, p.regcopies, l.lpexpirationdate
from hsi.productsold p
inner join hsi.licensedproduct l on p.producttype = l.producttype
where p.numcopieslic > 0
order by 1
"@
		Output = "$OutputDir\Licensed_Modules.xlsx"
	},
	@{
		Query  = @"
select rtrim(d.diskgroupname) as 'diskgroup', 
d.numberofbackups as 'copies', d.ucautopromotespace as 'size', 
d.lastlogicalplatter as 'volumes', rtrim(p.lastuseddrive) as 'path'
from hsi.diskgroup d
inner join hsi.physicalplatter p on d.diskgroupnum = p.diskgroupnum
order by 1

"@
		Output = "$OutputDir\Disk_Group_Summary.xlsx"
	},
	@{
		Query  = @"
select count(itemname) from hsi.itemdata

"@
		Output = "$OutputDir\Num_Docs.xlsx"
	},
	@{
		Query  = @"
select count(distinct filepath) from hsi.itemdatapage

"@
		Output = "$OutputDir\Num_Files.xlsx"
	},
	@{
		Query  = @"
SELECT case activestatus
	when 1 then 'Inactive'
	when 2 then 'Deleted'
	else 'Active' end as 'status',
count(DISTINCT objectid) AS 'count'
from hsi.rmObject
GROUP BY activestatus
order by 1

"@
		Output = "$OutputDir\Workview_Objects.xlsx"
	},
	@{
		Query  = @"
SELECT
rtrim(l.lifecyclename) as 'lifecycle', s.statenum, 
rtrim(s.statename) as 'queue', 
case wvobj.activestatus
	when 1 then 'Inactive'
	else 'Deleted' end as 'status', 
count(DISTINCT w.contentnum) AS 'count'
from hsi.lifecycle l
inner join hsi.lcstate s on s.scope = l.lcnum
inner join hsi.workitemlc w on w.statenum = s.statenum
INNER JOIN hsi.rmObject wvobj ON (wvobj.objectid = w.contentnum)
WHERE wvobj.activestatus <> 0 
GROUP BY l.lifecyclename, s.statenum, s.statename, wvobj.activestatus
ORDER BY 1,3,4

"@
		Output = "$OutputDir\Inactive_Workview_Objects.xlsx"
	},
	@{
		Query  = @"
select 'Incomplete Commit' as 'Status', rtrim(dg.diskgroupname) as 'Disk Group'
, ic.physicalplatternum as 'Copy', ic.logicalplatternum as 'Volume'
, count(ic.itemnum) as '# of Docs'
from hsi.incompletecommit ic
inner join hsi.diskgroup dg on dg.diskgroupnum = ic.diskgroupnum
group by dg.diskgroupname, ic.logicalplatternum, ic.physicalplatternum
union all
select 'Incomplete Delete' as 'Status', rtrim(dg.diskgroupname) as 'Disk Group'
, id.physicalplatternum as 'Copy', id.logicalplatternum as 'Volume'
, count(id.itemnum) as '# of Docs'
from hsi.incompletedelete id
inner join hsi.diskgroup dg on dg.diskgroupnum = id.diskgroupnum
group by dg.diskgroupname, id.logicalplatternum, id.physicalplatternum
order by 1,2,3,4

"@
		Output = "$OutputDir\Platter_Management.xlsx"
	},
	@{
		Query  = @"
select 'Docs' as 'Type', count(itemnum) from hsi.trashcan
union all
select 'Folders' as 'Type', count(foldernum) from hsi.foldertrashcan

"@
		Output = "$OutputDir\Deleted_Docs_Folders.xlsx"
	},
	@{
		Query  = @"
select rtrim(t.itemtypename) as 'Document Type', count(a.itemnum) as 'Count'
FROM  hsi.itemdata a, hsi.doctype t,
      (select a.itemtypenum, delafter & 65535 as day,
                (delafter & 16711680)/65536 as month,
                (delafter/16777216) & 65535 as year
                from  hsi.doctype a, hsi.doctypeext b
                where (itemtypeflags & 4096) = 4096
                and  a.itemtypenum = b.itemtypenum
                )   as dates
where dates.itemtypenum = a.itemtypenum
and a.itemtypenum = t.itemtypenum
and  dateadd(dd, dates.day,dateadd(mm, dates.month,dateadd(yy, dates.year,a.itemdate))) < getdate()
group by t.itemtypename
order by 1

"@
		Output = "$OutputDir\Docs_Exceeding_Retention.xlsx"
	},
	@{
		Query  = @"
select 'Document Lock' as locktype, hsi.doccheckout.checkouttime, rtrim(hsi.useraccount.username) as 'username', hsi.doccheckout.itemnum as 'item'
from hsi.doccheckout 
	left outer join hsi.useraccount on (hsi.doccheckout.usernum = hsi.useraccount.usernum) 
	left outer join hsi.itemdata on (hsi.doccheckout.itemnum = hsi.itemdata.itemnum)
 UNION ALL
select 'Process Lock' as locktype, hsi.lockkeys.locktime, rtrim(hsi.useraccount.username) as 'username', '' as 'item'
from hsi.lockkeys 
	left outer join hsi.useraccount on (hsi.lockkeys.usernum = hsi.useraccount.usernum)
UNION ALL
Select 'Batch Lock' as locktype, hsi.lockprocess.locktime, rtrim(hsi.useraccount.username) as 'username', hsi.lockprocess.batchnum as 'item'
From hsi.lockprocess
	left outer join hsi.useraccount on (hsi.lockprocess.usernum = hsi.useraccount.usernum)
Where  hsi.lockprocess.status =  1 
order by 1,2

"@
		Output = "$OutputDir\Active_Locks.xlsx"
	},
	@{
		Query  = @"
select count(itemname) from hsi.itemdata where itemtypenum = 0

"@
		Output = "$OutputDir\Unindexed_Docs.xlsx"
	},
	@{
		Query  = @"
select dt.itemtypenum as 'DocTypeID', 
rtrim(dt.itemtypename) as 'DocumentType', count(id.itemname) as 'Count' 
from hsi.doctype dt 
left outer join hsi.itemdata id
on dt.itemtypenum = id.itemtypenum
group by dt.itemtypenum, dt.itemtypename
order by dt.itemtypename

"@
		Output = "$OutputDir\Docs_Per_Doc_Type.xlsx"
	},
	@{
		Query  = @"
select rtrim(ug.usergroupname) as 'name'
from hsi.usergroup ug
left outer join hsi.userxusergroup uxg on ug.usergroupnum = uxg.usergroupnum
where uxg.usergroupnum is null
order by 1

"@
		Output = "$OutputDir\UserGroups_Not_In_Use.xlsx"
	},
	@{
		Query  = @"
select distinct 'User Group' as 'type', rtrim(g.usergroupname) as 'item', rtrim(k.keytype) as 'securedby'
from hsi.usergroupseckeys s
inner join hsi.usergroup g on s.usergroupnum = g.usergroupnum
inner join hsi.keytypetable k on s.keytypenum = k.keytypenum
union all
select 'User Account' as 'type', rtrim(u.username) as 'item', rtrim(k.keytype) as 'securedby'
from hsi.useraccountseckeys s
inner join hsi.useraccount u on s.usernum = u.usernum
inner join hsi.keytypetable k on s.keytypenum = k.keytypenum
where u.licenseflag & 2 <> 2
union all
select  'Document Type' as 'type', rtrim(t.itemtypename) as 'item', rtrim(g.usergroupname) as 'securedby'
from hsi.usergroupconfig u
inner join hsi.usergroup g on u.usergroupnum = g.usergroupnum
inner join hsi.doctype t on u.itemtypenum = t.itemtypenum
where u.flags > 0
order by 1,2,3

"@
		Output = "$OutputDir\Security_Overrides.xlsx"
	},
	@{
		Query  = @"
select rtrim(ua.username) as 'name'
from hsi.useraccount ua
left outer join hsi.userxusergroup uxg on ua.usernum = uxg.usernum
where uxg.usernum is null
and ua.licenseflag & 2 <> 2
and ua.licenseflag & 65536 <> 65536
order by 1

"@
		Output = "$OutputDir\UserAccounts_Not_In_Use.xlsx"
	},
	@{
		Query  = @"
select 'Locked Accounts' as 'cat', count(usernum) as 'count'
from hsi.useraccount
where licenseflag & 2 <> 2 and disablelogin = 1
union all
select 'User Group Administrators' as 'cat', count(usernum) as 'count'
from hsi.useraccount
where licenseflag & 2 <> 2 and licenseflag & 512 = 512
union all
select 'Change Password Disabled' as 'cat', count(usernum) as 'count'
from hsi.useraccount
where licenseflag & 2 <> 2 and licenseflag & 16384 = 16384
union all
select 'Service Accounts' as 'cat', count(usernum) as 'count'
from hsi.useraccount
where licenseflag & 2 <> 2 and licenseflag & 65536 = 65536
union all
select 'Stale Accounts' as 'cat', count(usernum) as 'count'
from hsi.useraccount
where licenseflag & 2 <> 2 and lastlogon < DATEADD(dd, -90, getdate())

"@
		Output = "$OutputDir\Notable_User_Accounts.xlsx"
	},
	@{
		Query  = @"
select rtrim(cast(usernum as char)) as 'usernum', rtrim(username) as 'username', 
case when disablelogin = 1 then 'Yes' else 'No' end as 'Locked', 
case when licenseflag & 512 = 512 then 'Yes' else 'No' end as 'UGA',
case when licenseflag & 16384 = 16384 then 'Yes' else 'No' end as 'DisableChgPass',
case when licenseflag & 65536 = 65536 then 'Yes' else 'No' end as 'SvcAcct',
lastlogon
from hsi.useraccount
where licenseflag & 2 <> 2
and (disablelogin = 1 or licenseflag & 512 = 512 
			  or licenseflag & 16384 = 16384 
			  or licenseflag & 65536 = 65536
			  or lastlogon < DATEADD(dd, -90, getdate()))
order by 2

"@
		Output = "$OutputDir\Notable_User_Account_Settings.xlsx"
	},
	@{
		Query  = @"
select rtrim(d.itemtypename) as 'doctype',
case 
	when d.itemtypeflags & 12288 = 12288 then 'Permanent'
	when d.itemtypeflags & 4096 = 4096 then 'Static'
	when d.itemtypeflags & 8192 = 8192 then 'Dynamic'
	else 'None' end as 'retentiontype',
case 
	when d.itemtypeflags & 16834 = 16384 then 'Date Stored'
	when d.itemtypeflags & 32768 = 32768 then 'Keyword'
	when d.itemtypeflags & 12288 = 12288 then ''
	when d.itemtypeflags & 4096 = 4096 then 'Document Date'
	when d.itemtypeflags & 8192 = 8192 then 'Document Date'
	else '' end as 'dateoptions',
case when d.itemtypeflags & 32768 = 32768 then rtrim(k.keytype) else '' end as 'keyword',
case when e.delafter > 0 then e.delafter/16777216 else 0 end as 'years',
case when e.delafter > 0 then (e.delafter/16777216)/65536 else 0 end as 'months',
case when e.delafter > 0 then ((e.delafter/16777216)/65536)%65536 else 0 end as 'days'
from hsi.doctype d
left outer join hsi.doctypeext e on d.itemtypenum = e.itemtypenum
left outer join hsi.keytypetable k on e.delkeytypenum = k.keytypenum
order by 1

"@
		Output = "$OutputDir\Document_Retention.xlsx"
	},
	@{
		Query  = @"
SELECT CASE WHEN q.queuename IS NULL THEN '<ALL QUEUES>' ELSE RTRIM(q.queuename) END as 'Queue',
SUM(ISNULL(batchesscanned.BatchesScanned, 0)) as 'BatchesCreated',
SUM(ISNULL(batchesindexed.BatchesIndexed, 0)) as 'BatchesIndexed',
SUM(ISNULL(batchesscanned.DocsCreated, 0)) as 'DocumentsCreated',
SUM(ISNULL(batchesindexed.DocsIndexed, 0)) as 'DocumentsIndexed',
SUM(ISNULL(batchesscanned.Pages, 0)) as 'PagesCreated',
SUM(ISNULL(batchesindexed.PagesIndexed, 0)) as 'PagesIndexed',
SUM(ISNULL(scanmoredocs.ScanMoreDocsCreated, 0)) as 'DocumentsAdded',
SUM(ISNULL(scanmorepages.ScanMorePages, 0)) as 'PagesAdded',
avg(ISNULL(p.avgtime,0)) as 'AvgProcessingHours'
FROM hsi.scanqueue q
INNER JOIN 
	(SELECT distinct queuenum FROM hsi.scanninglog 
			WHERE (logdate > DATEADD(dd, -90, getdate()))) s  ON q.queuenum = s.queuenum
LEFT OUTER JOIN
	(SELECT 
		s2.queuenum, COUNT(DISTINCT s2.batchnum) AS 'BatchesScanned', SUM(s2.extrainfo1) AS 'DocsCreated', 
		SUM(s2.extrainfo2) AS 'Pages'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum in (1,200)) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.queuenum) batchesscanned
ON s.queuenum = batchesscanned.queuenum
LEFT OUTER JOIN
	(SELECT 
		s2.queuenum, 
	COUNT(distinct s2.batchnum) AS 'BatchesIndexed', 
	COUNT(distinct i.itemnum) AS 'DocsIndexed', 
	COUNT(idp.itemnum) AS 'PagesIndexed'
	FROM hsi.scanninglog s2
	INNER JOIN hsi.itemdata i ON s2.batchnum = i.batchnum
	INNER JOIN hsi.itemdatapage idp ON idp.itemnum = i.itemnum
	WHERE 
		(s2.actionnum = 202) AND (s2.eventnum = 2) AND (i.status = 0) AND (i.itemtypegroupnum <> 1) 
		AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.queuenum) batchesindexed
ON s.queuenum = batchesindexed.queuenum
LEFT OUTER JOIN
	(SELECT 
		s2.queuenum, SUM(s2.extrainfo1) AS 'ScanMoreDocsCreated'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 205) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.queuenum) scanmoredocs
ON s.queuenum = scanmoredocs.queuenum
LEFT OUTER JOIN
	(SELECT 
		s2.queuenum, SUM(s2.extrainfo2) AS 'ScanMorePages'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 230) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate())) 
	GROUP BY s2.queuenum) scanmorepages
ON s.queuenum = scanmorepages.queuenum
LEFT OUTER JOIN 
	(select starttime.queuenum, avg(datediff(hour, starttime.logdate, endtime.endtime)) as 'avgtime'
	from hsi.scanninglog starttime
	inner join (select batchnum as 'batch', logdate as 'endtime'
				from hsi.scanninglog
				where (actionnum = 8)
				and (logdate > DATEADD(dd, -90, getdate()))
				) endtime on endtime.batch = starttime.batchnum
	where (starttime.actionnum = 200)
	and (starttime.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY starttime.queuenum) p on q.queuenum = p.queuenum
GROUP BY q.queuename WITH ROLLUP

"@
		Output = "$OutputDir\Scan_Summary_By_Queue.xlsx"
	},
	@{
		Query  = @"
select CASE WHEN u.username IS NULL THEN '<ALL USERS>' ELSE RTRIM(u.username) END as 'UserName',
SUM(ISNULL(batchesscanned.BatchesScanned, 0)) as 'BatchesScanned',
SUM(ISNULL(batchesindexed.BatchesIndexed, 0)) as 'BatchesIndexed',
SUM(ISNULL(batchesscanned.DocsCreated, 0)) as 'DocumentsCreated',
SUM(ISNULL(batchesindexed.DocsIndexed, 0)) as 'DocumentsIndexed',
SUM(ISNULL(batchesscanned.Pages, 0)) as 'PagesCreated',
SUM(ISNULL(batchesindexed.PagesIndexed, 0)) as 'PagesIndexed',
SUM(ISNULL(scanmoredocs.ScanMoreDocsCreated, 0)) as 'ScanMoreDocuments',
SUM(ISNULL(scanmorepages.ScanMorePages, 0)) as 'ScanMorePages',
avg(ISNULL(p.avgmin,0)) as 'BatchAvgIndexMinutes',
SUM(ISNULL(imagequality.IQBatches, 0)) as 'BatchesImageQualityReviewed',
SUM(ISNULL(docslice.DocSlice, 0)) as 'BatchesSeparated',
SUM(ISNULL(nodocslice.NoDocSlice, 0)) as 'BatchesSkippedSeparation',
SUM(ISNULL(batchesreindexed.DocsReIndexed, 0)) as 'DocumentsRe-Indexed',
SUM(ISNULL(batchesreindexed.PagesReIndexed, 0)) as 'PagesRe-Indexed',
SUM(ISNULL(batchesindexed.DocTypesUsed, 0)) as 'DocumentTypesUsed',
SUM(ISNULL(qareview.QAbatches, 0)) as 'BatchesQAReviewed',
SUM(ISNULL(batchesdeleted.BatchesDeleted, 0)) as 'BatchesDeleted',
SUM(ISNULL(batchespurged.BatchesPurged, 0)) as 'BatchesPurged'
FROM hsi.useraccount u
INNER JOIN 
	(SELECT distinct usernum FROM hsi.scanninglog 
			WHERE (logdate > DATEADD(dd, -90, getdate()))) s  ON u.usernum = s.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(DISTINCT s2.batchnum) AS 'BatchesScanned', SUM(s2.extrainfo1) AS 'DocsCreated', 
		SUM(s2.extrainfo2) AS 'Pages'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum in (1,200)) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) batchesscanned
ON s.usernum = batchesscanned.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, 
	COUNT(distinct s2.batchnum) AS 'BatchesIndexed', 
	COUNT(distinct i.itemnum) AS 'DocsIndexed', 
	COUNT(idp.itemnum) AS 'PagesIndexed', 
	COUNT(distinct i.itemtypenum) AS 'DocTypesUsed'
	FROM hsi.scanninglog s2
	INNER JOIN hsi.itemdata i ON s2.batchnum = i.batchnum AND i.usernum = s2.usernum
	INNER JOIN hsi.itemdatapage idp ON idp.itemnum = i.itemnum
	WHERE 
		(s2.actionnum = 202) AND (s2.eventnum = 2) AND (i.status = 0) AND (i.itemtypegroupnum <> 1) 
		AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) batchesindexed
ON s.usernum = batchesindexed.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(distinct i.itemnum) AS 'DocsReIndexed', COUNT(distinct idp.itemnum) AS 'PagesReIndexed'
	FROM hsi.scanninglog s2
	INNER JOIN hsi.itemdata i ON s2.batchnum = i.batchnum
	INNER JOIN hsi.itemdatapage idp ON idp.itemnum = i.itemnum
	WHERE 
		(s2.actionnum = 213) AND (s2.eventnum = 2) AND (i.status = 0) and (i.itemtypegroupnum <> 1) 
		AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) batchesreindexed
ON s.usernum = batchesreindexed.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(s2.actionnum) AS 'IQBatches'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 217) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate())) 
	GROUP BY s2.usernum) imagequality
ON s.usernum = imagequality.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, SUM(s2.extrainfo1) AS 'ScanMoreDocsCreated'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 205) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) scanmoredocs
ON s.usernum = scanmoredocs.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, SUM(s2.extrainfo2) AS 'ScanMorePages'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 230) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate())) 
	GROUP BY s2.usernum) scanmorepages
ON s.usernum = scanmorepages.usernum
LEFT OUTER JOIN
	(SELECT s2.usernum, COUNT(s2.actionnum) AS 'DocSlice'
	FROM hsi.scanninglog s2
	WHERE 
	(s2.actionnum = 227) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) docslice
ON s.usernum = docslice.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(s2.actionnum) AS 'NoDocSlice'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 226) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) nodocslice
ON s.usernum = nodocslice.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(s2.actionnum) AS 'QABatches'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 218) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) qareview
ON s.usernum = qareview.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(s2.actionnum) AS 'BatchesDeleted'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 201) AND (s2.eventnum = 2) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) batchesdeleted
ON s.usernum = batchesdeleted.usernum
LEFT OUTER JOIN
	(SELECT 
		s2.usernum, COUNT(s2.actionnum) AS 'BatchesPurged'
	FROM hsi.scanninglog s2
	WHERE 
		(s2.actionnum = 9) AND (s2.eventnum = 1) AND (s2.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY s2.usernum) batchespurged
ON s.usernum = batchespurged.usernum
LEFT OUTER JOIN --indexing time for a batch is from actionnum 202 to 266
	(select starttime.usernum, avg(datediff(minute, starttime.logdate, endtime.endtime)) as 'avgmin'
	from hsi.scanninglog starttime
	inner join (select batchnum as 'batch', logdate as 'endtime'
				from hsi.scanninglog
				where (actionnum = 266)
				and (logdate > DATEADD(dd, -90, getdate()))
				) endtime on endtime.batch = starttime.batchnum
	where (starttime.actionnum = 202)
	and (starttime.logdate > DATEADD(dd, -90, getdate()))
	GROUP BY starttime.usernum) p on s.usernum = p.usernum
GROUP BY u.username WITH ROLLUP

"@
		Output = "$OutputDir\User_Summary_by_Date.xlsx"
	},
	@{
		Query  = @"
Select rtrim(queuename) as 'Queue'
, case when status = 0 then 'Awaiting Index'
	when status = 1 then 'Index in Progress'
	when status = 2 then 'Awaiting Commit'
	when status in (3,4,5,6) then 'Incomplete Commit'
	when status = 9 then 'Incomplete Purge'
	when status = 14 then 'Awaiting OCR'
	when status = 17 then 'Checked Out Disconnected Scanning'
	when status = 18 then 'Disconnected Scan Incomplete Upload'
	when status = 19 then 'Incomplete Archive'
	when status = 20 then 'Secondary Awaiting Index'
	when status = 21 then 'Secondary Index in Progress'
	when status = 22 then 'Failed Automatic OCR'
	when status = 23 then 'Awaiting Doc Separation'
	when status = 24 then 'Line Item Separation'
	when status = 25 then 'ADF Error Queue'
	when status = 26 then 'Awaiting Re-Index'
	when status = 27 then 'Re-Index in Progress'
	when status = 28 then 'Check Error Queue'
	when status = 29 then 'ADF Decisioning Queue'
	when status = 30 then 'Administrator Repair'
	when status = 31 then 'Awaiting QA Image Quality Review'
	when status = 32 then 'Awaiting QA Review'
	when status = 33 then 'Awaiting QA ReScan'
	when status = 34 then 'Awaiting Manager Resolution'
	when status = 35 then 'Awaiting QA Re-Index'
	when status = 36 then 'QA Re-Index in Progress'
	when status = 37 then 'In process'
	when status = 38 then 'Awaiting PDF Conversion'
	when status = 39 then 'Scheduled Processes'
	when status = 40 then 'Error Correction Queue'
	when status = 41 then 'Awaiting Transfer to Host'
	when status = 43 then 'Awaiting External Index'
	when status = 44 then 'Awaiting Barcode Processing'
	when status = 45 then 'ADF Decision Error Queue'
	when status = 46 then 'Awaiting Image Process'
	when status = 47 then 'Custom Process'
	when status = 48 then 'Ad Hoc Re-scan'
	when status = 49 then 'Branch Capture Balancing Queue'
	when status = 50 then 'Branch Capture In Process Queue'
	when status = 51 then 'Awaiting Zonal OCR'
	when status = 52 then 'Awaiting Ad Hoc Zonal OCR'
	when status = 53 then 'Pull Slips'
	when status = 54 then 'Awaiting Ad Hoc Verification'
	when status = 55 then 'QA Review In Progress'
	when status = 56 then 'Synchronization Pending'
	when status = 57 then 'Synchronization Complete'
	when status = 58 then 'Synchronization Failed'
	when status = 59 then 'Synchronization Queued'
	when status = 60 then 'Synchronization Processed'
	when status = 61 then 'Export Awaiting Transfer'
	when status = 62 then 'Export Pending Verification'
	when status = 63 then 'Export Complete'
	when status = 64 then 'Export Error'
	when status = 65 then 'Synchronization History'
	when status = 66 then 'Doc Transfer Export History'
	when status = 67 then 'Awaiting Formless Indexing'
	when status = 68 then 'Awaiting Queue Sorting'
	else 'Other' end as 'Batch Status'
, Count(*) as 'Count'
From hsi.archivedqueue
where status <> 8
group by queuename, status
UNION ALL
select rtrim(queuename) as 'Queue'
, 'Committed' as 'Batch Status', count(*) as 'Count'
from hsi.archivedcommitq
group by queuename
UNION ALL
Select '<ALL QUEUES>' as 'Queue'
, case when status = 0 then 'Awaiting Index'
	when status = 1 then 'Index in Progress'
	when status = 2 then 'Awaiting Commit'
	when status in (3,4,5,6) then 'Incomplete Commit'
	when status = 9 then 'Incomplete Purge'
	when status = 14 then 'Awaiting OCR'
	when status = 17 then 'Checked Out Disconnected Scanning'
	when status = 18 then 'Disconnected Scan Incomplete Upload'
	when status = 19 then 'Incomplete Archive'
	when status = 20 then 'Secondary Awaiting Index'
	when status = 21 then 'Secondary Index in Progress'
	when status = 22 then 'Failed Automatic OCR'
	when status = 23 then 'Awaiting Doc Separation'
	when status = 24 then 'Line Item Separation'
	when status = 25 then 'ADF Error Queue'
	when status = 26 then 'Awaiting Re-Index'
	when status = 27 then 'Re-Index in Progress'
	when status = 28 then 'Check Error Queue'
	when status = 29 then 'ADF Decisioning Queue'
	when status = 30 then 'Administrator Repair'
	when status = 31 then 'Awaiting QA Image Quality Review'
	when status = 32 then 'Awaiting QA Review'
	when status = 33 then 'Awaiting QA ReScan'
	when status = 34 then 'Awaiting Manager Resolution'
	when status = 35 then 'Awaiting QA Re-Index'
	when status = 36 then 'QA Re-Index in Progress'
	when status = 37 then 'In process'
	when status = 38 then 'Awaiting PDF Conversion'
	when status = 39 then 'Scheduled Processes'
	when status = 40 then 'Error Correction Queue'
	when status = 41 then 'Awaiting Transfer to Host'
	when status = 43 then 'Awaiting External Index'
	when status = 44 then 'Awaiting Barcode Processing'
	when status = 45 then 'ADF Decision Error Queue'
	when status = 46 then 'Awaiting Image Process'
	when status = 47 then 'Custom Process'
	when status = 48 then 'Ad Hoc Re-scan'
	when status = 49 then 'Branch Capture Balancing Queue'
	when status = 50 then 'Branch Capture In Process Queue'
	when status = 51 then 'Awaiting Zonal OCR'
	when status = 52 then 'Awaiting Ad Hoc Zonal OCR'
	when status = 53 then 'Pull Slips'
	when status = 54 then 'Awaiting Ad Hoc Verification'
	when status = 55 then 'QA Review In Progress'
	when status = 56 then 'Synchronization Pending'
	when status = 57 then 'Synchronization Complete'
	when status = 58 then 'Synchronization Failed'
	when status = 59 then 'Synchronization Queued'
	when status = 60 then 'Synchronization Processed'
	when status = 61 then 'Export Awaiting Transfer'
	when status = 62 then 'Export Pending Verification'
	when status = 63 then 'Export Complete'
	when status = 64 then 'Export Error'
	when status = 65 then 'Synchronization History'
	when status = 66 then 'Doc Transfer Export History'
	when status = 67 then 'Awaiting Formless Indexing'
	when status = 68 then 'Awaiting Queue Sorting'
	else 'Other' end as 'Batch Status'
, Count(*) as 'Count'
From hsi.archivedqueue
where status <> 8
group by status
UNION ALL
select '<ALL QUEUES>' as 'Queue'
, 'Committed' as 'Batch Status', count(*) as 'Count'
from hsi.archivedcommitq
UNION ALL
select rtrim(q.queuename) as 'Queue', 'No batches in this scan queue' as 'Batch Status', 0 as 'Count'
from hsi.scanqueue q
left outer join (select distinct queuename from hsi.archivedqueue
				union all
				select distinct queuename from hsi.archivedcommitq) l on l.queuename = q.queuename
where l.queuename is null

"@
		Output = "$OutputDir\Scan_Queue_Summary.xlsx"
	},
	@{
		Query  = @"
select 'Ad Hoc Import' as 'method', COUNT(i.itemnum) as 'count'
from hsi.itemdata i
where i.batchnum = 0
and i.datestored > DATEADD(dd, -90, getdate())
group by i.batchnum
UNION ALL
select 'Scan / Sweep' as 'method', SUM(s.extrainfo1) as 'count'
from hsi.scanninglog s
--where s.actionnum in (200,205)
where s.actionnum = 1
and s.logdate > DATEADD(dd, -90, getdate())
group by s.actionnum
UNION ALL
select case when p.parsingmethod = 1 then 'COLD'
       when p.parsingmethod = 41 then 'COLD - Visual PDF'
       when p.parsingmethod = 73 then 'EDI'
       when p.parsingmethod = 74 then 'EDI'
       when p.parsingmethod = 76 then 'EDI'
       when p.parsingmethod = 122 then 'DrIP'
       when p.parsingmethod & 4 = 4 then 'DIP'
       end  as 'method',
COUNT(i.itemnum) as 'count'
FROM hsi.itemdata i, hsi.parsedqueue p
        WHERE i.batchnum = p.batchnum
        AND (p.parsingmethod & 4 = 4 or p.parsingmethod in (1,41,73,74,122))
        AND i.datestored > DATEADD(dd, -90, getdate())
group by p.parsingmethod

"@
		Output = "$OutputDir\Docs_Ingested_Method.xlsx"
	},
	@{
		Query  = @"
--scan queues
select 'Scan Queue' as 'Import Method'
, rtrim(queuename) as 'Name'
, RTRIM(sweepdir) AS 'Import Directory'
, '' as 'Index'
from hsi.scanqueue
union all
--most other import methods
select case when parsingmethod = 1 then 'COLD'
       when parsingmethod = 40 then 'Autofill Keyword Processor'
       when parsingmethod = 41 then 'COLD - Visual PDF'
       when parsingmethod = 64 then 'Keyword Update - Global'
       when parsingmethod = 65 then 'Keyword Update - Doc Type'
       when parsingmethod = 73 then 'EDI'
       when parsingmethod = 74 then 'EDI'
       when parsingmethod = 76 then 'EDI'
       when parsingmethod = 122 then 'DrIP'
       when parsingmethod & 4 = 4 then 'DIP' --4, 65540, 131076, 262148
       else 'Other'
       end as 'Import Method'
, RTRIM(parsefilename) AS 'Name'
, RTRIM(defdirname) AS 'Import Directory'
, RTRIM(deffilename) AS 'Index'
from hsi.parsefiledesc
where parsingmethod not in (46,47)
ORDER BY 1,2

"@
		Output = "$OutputDir\Common_Import_Methods.xlsx"
	},
	@{
		Query  = @"
SELECT RTRIM(ssaccountname) AS 'Name'
, RTRIM(serveraddress) AS 'Server Address'
, RTRIM(mailacctusername) AS 'Email Account'
FROM hsi.ssaccount
order by 1

"@
		Output = "$OutputDir\Mailbox_Importer.xlsx"
	},
	@{
		Query  = @"
select s.keysettablenum, rtrim(s.keysetname) as 'keysetname', rtrim(t.keytype) as 'primary',
case when s.flags & 32 = 32 then 'external' else 'internal' end as 'source'
from hsi.keywordset s
inner join hsi.keysetxkeytype x on s.keysettablenum = x.keysettablenum
join hsi.keytypetable t on x.keytypenum = t.keytypenum
where x.seqnum = 0
order by 2

"@
		Output = "$OutputDir\Autofill_Keyword_Sets.xlsx"
	},
	@{
		Query  = @"
SELECT CASE p.schedtype
	WHEN 1 THEN 'Parse Format COLD/DIP'
	WHEN 2 THEN 'Parse Job'
	WHEN 3 THEN 'Scan Queue Sweep'
	WHEN 7 THEN 'Automatic Commit - Scan'
	WHEN 8 THEN 'Automatic Commit - COLD/DIP'
	WHEN 13 THEN 'Document Maintenance Purge'
	WHEN 18 THEN 'Signature Polling'
	WHEN 19 THEN 'Parse Format - Document Retention'
	ELSE 'Other' END AS 'Process Type', 
RTRIM(p.schedprocname) AS 'Process Name', 
RTRIM(r.registername) AS 'Assigned Workstation', 
p.lastprocdate AS 'Last Run'
FROM hsi.scheduledprocess p
INNER JOIN hsi.registeredusers r ON r.registernum = p.registernum
order by 1,2

"@
		Output = "$OutputDir\Client_Scheduled_Tasks.xlsx"
	},
	@{
		Query  = @"
select rtrim(keytype) as 'keytype',
CASE datatype
	WHEN 0 THEN 'NULL'
	WHEN 1 THEN 'Numeric (20)'
	WHEN 2 THEN 'Alphanum DT (' + convert(varchar,keytypelen) + ')'
	WHEN 3 THEN 'Currency'
	WHEN 4 THEN 'Date'
	WHEN 5 THEN 'Float'
	WHEN 6 THEN 'Numeric (9)'
	WHEN 9 THEN 'Date/Time'
	WHEN 10 THEN 'Alphanum ST (' + convert(varchar,keytypelen) + ')'
	WHEN 11 THEN 'Specific Currency'
	WHEN 12 THEN 'Alphanum DT (' + convert(varchar,keytypelen) + ')'
	WHEN 13 THEN 'Alphanum ST (' + convert(varchar,keytypelen) + ')'
	ELSE 'Unknown' END as 'DataType',
CASE 
	WHEN keytypeflags < 0 then 'Using Data Set: External'
	WHEN keytypeflags > 600000000 then 'Using Data Set: Custom'
	WHEN keytypeflags & 34603008 = 34603008 then 'Using Data Set: Descending'
	WHEN keytypeflags & 33554432 = 33554432 then 'Using Data Set: Ascending'
	ELSE '' END as 'Extras',
CASE
	WHEN keytypemask > '' then 'Mask: ' + keytypemask
	ELSE '' END as 'Mask'
FROM hsi.keytypetable
where datatype in (3,4,9,11) or (keytypeflags & 33554432 = 33554432) or (keytypeflags & 1 = 1)
union all
select k.keytype, 
CASE datatype
	WHEN 0 THEN 'NULL'
	WHEN 1 THEN 'Numeric (20)'
	WHEN 2 THEN 'Alphanum DT (' + convert(varchar,keytypelen) + ')'
	WHEN 3 THEN 'Currency'
	WHEN 4 THEN 'Date'
	WHEN 5 THEN 'Float'
	WHEN 6 THEN 'Numeric (9)'
	WHEN 9 THEN 'Date/Time'
	WHEN 10 THEN 'Alphanum ST (' + convert(varchar,keytypelen) + ')'
	WHEN 11 THEN 'Specific Currency'
	WHEN 12 THEN 'Alphanum DT (' + convert(varchar,keytypelen) + ')'
	WHEN 13 THEN 'Alphanum ST (' + convert(varchar,keytypelen) + ')'
	ELSE 'Unknown' END as 'DataType',
'Not In Use',''
from hsi.keytypetable k
left outer join hsi.itemtypexkeyword x on x.keytypenum = k.keytypenum
where x.keytypenum is null
order by 2,1

"@
		Output = "$OutputDir\Notable_Keyword_Types.xlsx"
	},
	@{
		Query  = @"
select rtrim(q.cqname) as 'Custom Query'
, case when q.cqflags & 512 = 512 then 'Custom Written Sql'
		when q.cqflags & 2048 = 2048 then 'By Document Type'
		when q.cqflags & 4096 = 4096 then 'By Document Type Group'
		when q.cqflags & 8192 = 8192 then 'By Keyword'
		when q.cqflags & 524288 = 524288 then 'By Folder Type'
		else 'By Full Text' end as Type
, case when q.cqflags & 65536 = 65536 or q.cqflags & 2097152 = 2097152 then 'Yes' else 'No' end as 'Restrict by Rights'
, case when q.cqflags & 262144 = 262144 then 'Yes' else 'No' end as 'Minimize Duplicate Docs'
, case when q.cqflags & 16384 = 16384 then 'Yes' else 'No' end as 'Workflow Filter'
, case when q.cqflags & 131072 = 131072 then 'Yes' else 'No' end as 'Folder Filter'
, case when q.cqflags & 8388608 = 8388608 then 'Yes' else 'No' end as 'Mobile Accessible'
, case when q.cqusekeys = 1 then 'Yes' else 'No' end as 'Keyword Edit Fields'
, case when q.cqflags & 2 = 2 then 'Yes' else 'No' end as 'Text Search Button'
, case when q.cqflags & 256 = 256 then 'Yes' else 'No' end as 'Use HTML Form'
, count(x.cqnum) as '# User Groups'
from hsi.customquery q
left outer join (select u.cqnum, g.usergroupname from hsi.usergcustomquery u
				inner join hsi.usergroup g on g.usergroupnum = u.usergroupnum) x on x.cqnum = q.cqnum
group by q.cqname, q.cqflags, q.cqviewall, q.cqusedate, q.cqusekeys
order by q.cqname

"@
		Output = "$OutputDir\Custom_Query_Settings.xlsx"
	},
	@{
		Query  = @"
select t.vbscriptnum, t.vbscriptname, u.usergroupname
from hsi.vbscripttable t
left outer join hsi.vbscripthooks h on t.vbscriptnum = h.vbscriptnum
left outer join hsi.usergroup u on h.usergroupnum = u.usergroupnum
order by t.vbscriptnum, u.usergroupname

"@
		Output = "$OutputDir\VB_Scripts.xlsx"
	},
	@{
		Query  = @"
select case when ft.foldertypeflags & 536870912 = 536870912 then 'Workflow, Exclude Primary Doc'
		when ft.foldertypeflags & 4096 = 4096 then 'Workflow'
		else 'Client' end as 'Usage'
, rtrim(ftp.foldertypename) as 'Parent Folder'
, rtrim(ft.foldertypename) as 'Folder'
, case when ft.foldertypeflags & 3 = 3 then 'Static and Dynamic Docs'
		when ft.foldertypeflags & 1 = 1 then 'Dynamic Doc Types'
		when ft.foldertypeflags & 2 = 2 then 'Static Docs'
		when ft.foldertypeflags & 4 = 4 then 'Dynamic Doc Type Groups'
		else 'None (Folders Only)' end as 'Contents'
, count(ug.foldertypenum) as '# User Groups'
, case when kt.numkeywords is null then 0 else kt.numkeywords end as '# Keyword Types'
, case when dt.numdocs is null then 0 else dt.numdocs end as '# Doc Types'
from hsi.foldertype ft
left outer join hsi.foldertype ftp on ftp.foldertypenum = ft.prntfoldertypenum
left outer join (select u.foldertypenum, g.usergroupname from hsi.usergfoldertype u
				inner join hsi.usergroup g on g.usergroupnum = u.usergroupnum
				) ug on ug.foldertypenum = ft.foldertypenum
left outer join (select foldertypenum, count(keytypenum) as 'numkeywords' 
				from hsi.foldtypexkeyword group by foldertypenum
				) kt on kt.foldertypenum = ft.foldertypenum
left outer join (select foldertypenum, count(itemtypenum) as 'numdocs' 
				from hsi.dynfoldinfoseq group by foldertypenum
				) dt on dt.foldertypenum = ft.foldertypenum
group by ft.foldertypename, ftp.foldertypename, ft.foldertypeflags, kt.numkeywords, dt.numdocs
order by 1,2,3

"@
		Output = "$OutputDir\Folder_Types.xlsx"
	},
	@{
		Query  = @"
select rtrim(l.lifecyclename) as 'LifeCycle', rtrim(s.statename) as 'Queue', count(w.itemnum) as 'Count'
from hsi.itemlc w
inner join hsi.lifecycle l on l.lcnum = w.lcnum
inner join hsi.lcstate s on s.statenum = w.statenum
group by l.lifecyclename, l.lcnum, s.statename, s.statenum
union all
select rtrim(l.lifecyclename) as 'LifeCycle', rtrim(s.statename) as 'Queue', count(w.contentnum) as 'Count'
from hsi.workitemlc w
inner join hsi.lifecycle l on l.lcnum = w.lcnum
left outer join hsi.lcstate s on s.statenum = w.statenum
group by l.lifecyclename, l.lcnum, s.statename, s.statenum
order by 3 desc,1,2

"@
		Output = "$OutputDir\Queue_Counts.xlsx"
	},
	@{
		Query  = @"
SELECT 'User Group' as 'UserOrGroup', rtrim(ug.usergroupname) as 'UserOrGroupName', 'n/a' as 'Lifecycle', 'n/a' as 'Queue',
'Administrative Processing Privileges - Workflow' as 'Privilege'
FROM hsi.usergroup ug
WHERE ug.userprivilege3 & 268435456 = 268435456
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Queue Administration'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 1 = 1
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'See Other Users Documents'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 2 = 2
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Execute System Work'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 4 = 4
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Execute Timer'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 8 = 8
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Ad hoc Routing'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 16 = 16
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Override Auto-Feed'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 32 = 32
UNION ALL
select 'User Group', rtrim(g.usergroupname), rtrim(l.lifecyclename), rtrim(s.statename), 'Ownership Administration'
from hsi.lcstateusergprivs p
inner join hsi.usergroup g on p.usergroupnum = g.usergroupnum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 64 = 64
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Queue Administration'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 1 = 1
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'See Other Users Documents'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 2 = 2
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Execute System Work'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 4 = 4
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Execute Timer'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 8 = 8
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Ad hoc Routing'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 16 = 16
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Override Auto-Feed'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 32 = 32
UNION ALL
select 'User', rtrim(a.username), rtrim(l.lifecyclename), rtrim(s.statename), 'Ownership Administration'
from hsi.lcstateuserprivs p
inner join hsi.useraccount a on p.usernum = a.usernum
inner join hsi.lcstate s on p.statenum = s.statenum
inner join hsi.lifecycle l on l.lcnum = s.scope
WHERE p.userprivilege0 & 64 = 64
order by 1,2,3,4,5

"@
		Output = "$OutputDir\Admin_Processing_Priv.xlsx"
	},
	@{
		Query  = @"
select * from 
(select rtrim(l.lifecyclename) as 'LifeCycle', l.lcnum, rtrim(s.statename) as 'Queue', s.statenum,
case when s.queuetype & 1052674 = 1052674 then 'Allocated Percentage (User Groups)'
	when s.queuetype & 1056770 = 1056770 then 'Allocated Percentage (Users)'
	when s.queuetype & 135170 = 135170 then 'By Priority (User Groups)'
	when s.queuetype & 139266 = 139266 then 'By Priority (Users)'
	when s.queuetype & 69634 = 69634 then 'n Order (User Groups)'
	when s.queuetype & 73730 = 73730 then 'In Order (Users)'
	when s.queuetype & 4198402 = 4198402 then 'Keyword Based (User Groups)'
	when s.queuetype & 4202498 = 4202498 then 'Keyword Based (Users)'
	when s.queuetype & 8392706 = 8392706 then 'Match Keyword to User Name (User Groups)'
	when s.queuetype & 8396802 = 8396802 then 'Match Keyword to User Name (Users)'
	when s.queuetype & 266242 = 266242 then 'Rules Based (User Groups)'
	when s.queuetype & 270338 = 270338 then 'Rules Based (Users)'
	when s.queuetype & 274434 = 274434 then 'Rules Based (Roles)'
	when s.queuetype & 528386 = 528386 then 'Shortest Queue (User Groups)'
	when s.queuetype & 532482 = 532482 then 'Shortest Queue (Users)'
	else 'Unbalanced'
	end as 'LoadBalanceMethod',
case when s.queuetype & 134217728 = 134217728 then 'Yes' else 'No' end as 'SecurityKeywords',
s.obrefresh as 'InboxRefreshRate'
from hsi.lcstate s
inner join hsi.lifecycle l on s.scope = l.lcnum) a
where a.LoadBalanceMethod <> 'Unbalanced' or a.SecurityKeywords = 'Yes' or a.InboxRefreshRate > 0
order by 1,3

"@
		Output = "$OutputDir\Notable_Queue_Settings.xlsx"
	},
	@{
		Query  = @"
select 1 as 'SortOrder', 'Action' as 'Type', a.actionnum as 'ID', rtrim(a.actionname) as 'Name', 
'' as 'ChildTask', '' as 'ChildType', rtrim(l.lifecyclename) as 'LifeCycle'
from hsi.action a
left outer join hsi.lifecycle l on l.lcnum = a.scope
where a.flags & 1073741824 = 1073741824
union all
select 2 as 'SortOrder', 'Rule' as 'Type', r.rulenum as 'ID', rtrim(r.rulename) as 'Name', 
'' as 'ChildTask', '' as 'ChildType', rtrim(l.lifecyclename) as 'LifeCycle'
from hsi.ruletable r
left outer join hsi.lifecycle l on l.lcnum = r.scope
where r.flags & 524288 = 524288
union all
SELECT 3 as 'SortOrder', 'Task' as 'Type', t.tasklistnum as 'ID', rtrim(t.tasklistname) as 'Name', 
'' as 'ChildTask', '' as 'ChildType', rtrim(l.lifecyclename) as 'LifeCycle'
FROM hsi.tasklist t
left outer join hsi.lifecycle l on l.lcnum = t.scope
where t.flags & 262144 = 262144
and t.tasklistname NOT IN ('On True','On False')
union all
select 4 as 'SortOrder', rtrim(t.tasklistname) as 'Type', t.tasklistnum as 'ID', '' as 'Name',
x.tasknum as 'ChildTask', case when x.flags = 1 then 'Rule' 
							   when x.flags = 2 then 'Action' 
							   else '' end as 'ChildType', 
rtrim(l.lifecyclename) as 'LifeCycle'
FROM hsi.tasklist t
left outer join hsi.tasklistxtask x on x.tasklistnum = t.tasklistnum
left outer join hsi.lifecycle l on l.lcnum = t.scope
where t.flags & 262144 = 262144
and t.tasklistname IN ('On True','On False')
union all
select 5 as 'SortOrder', 'Timer' as 'Type', t.timernum as 'ID', rtrim(t.timername) as 'Name', 
'' as 'ChildTask', '' as 'ChildType', rtrim(l.lifecyclename) as 'LifeCycle'
from hsi.lctimer t
left outer join hsi.lifecycle l on l.lcnum = t.scope
where t.flags & 524288 = 524288
union all
select 6 as 'SortOrder', 'Custom Log Entry' as 'Type', a.actionnum as 'ID', rtrim(a.actionname) as 'Name',
'' as 'ChildTask', '' as 'ChildType', rtrim(l.lifecyclename) as 'LifeCycle'
from hsi.action a
left outer join hsi.lifecycle l on l.lcnum = a.scope
where a.actiontype = 96
order by 1,3


"@
		Output = "$OutputDir\Configured_Logging.xlsx"
	},
	@{
		Query  = @"
SELECT case when a.actiontype > 0 then 'Yes'
       else 'No'
       end as 'InUse',
RTRIM(l.lifecyclename) AS 'LifeCycle', nl.notilistnum, 
RTRIM(nl.notilistname) AS 'Notification',
CASE WHEN n.email > '' THEN RTRIM(n.email)
       ELSE ''
       END AS 'Recipient'
FROM hsi.notificationlist nl
LEFT OUTER JOIN hsi.lifecycle l ON nl.scope = l.lcnum
INNER JOIN hsi.notification n ON nl.notilistnum = n.notilistnum
LEFT OUTER JOIN hsi.action a on a.notilistnum = nl.notilistnum
where a.actiontype is null or n.email > ''
ORDER BY 1,2,5

"@
		Output = "$OutputDir\Workflow_Notifications.xlsx"
	},
	@{
		Query  = @"
select rtrim(l.lifecyclename) as 'LifeCycle',
'Existing Notification' as 'Type',
RTRIM(ap.rmapplicationname) AS 'Application', RTRIM(c.classname) AS 'Class',
RTRIM(n.rmname) AS 'Notification', n.notificationid AS 'ID',
CASE WHEN a.actionname IS NULL THEN 'Notification Not in Use' ELSE RTRIM(a.actionname) end AS 'Action',
p.actionnum AS 'ActionID'
from hsi.rmnotification n
inner join hsi.rmclass c on c.classid = n.ownerobjectid
inner JOIN hsi.rmapplicationclasses ac ON ac.classid = c.classid
inner JOIN hsi.rmapplication ap ON ap.rmapplicationid = ac.rmapplicationid
LEFT outer join hsi.actionprops p on p.propertyvalue = convert(varchar,n.notificationid)
                                                              AND p.propertyname = 'NotificationID'
LEFT outer join hsi.action a on a.actionnum = p.actionnum
left outer join hsi.lifecycle l on l.lcnum = a.lcnum
union all
select rtrim(l.lifecyclename) as 'LifeCycle',
'Custom Notification' as 'Type',
'' AS 'Application', '' AS 'Class',
'' as 'Notification', 0 AS 'ID',
CASE WHEN t.tasknum IS NULL THEN 'Notification Not in Use' ELSE RTRIM(a.actionname) end AS 'Action', 
p.actionnum AS 'ActionID'
from hsi.actionprops p
inner join hsi.action a on a.actionnum = p.actionnum
left outer join hsi.tasklistxtask t on t.tasknum = a.actionnum
left outer join hsi.lifecycle l on l.lcnum = a.scope
where p.propertyname = 'NotificationSource' and p.propertyvalue = '2'

"@
		Output = "$OutputDir\Workview_Notifications.xlsx"
	},
	@{
		Query  = @"
select rtrim(l.lifecyclename) as 'lifecycle', rtrim(s.statename) as 'queue', 
rtrim(t.timername) as 'timername', t.timernum,
case when t.flags & 256 = 256 then 
	(case when t.flags & 64 = 64 then 'M' else '' end +
	case when t.flags & 32 = 32 then 'Tu' else '' end +
	case when t.flags & 16 = 16 then 'W' else '' end +
	case when t.flags & 8 = 8 then 'Th' else '' end +
	case when t.flags & 3 = 3 then 'F' else '' end +
	case when t.flags & 2 = 2 then 'Sa' else '' end +
	case when t.flags & 1 = 1 then 'Su' else '' end + ' ' +
	case when t.hours < 10 then '0' else '' end +
	convert(varchar,t.hours) + ':' +
	case when t.minutes < 10 then '0' else '' end +
	convert(varchar,t.minutes)) end as 'certaintime',
case when t.flags & 74240 = 74240 then convert(varchar,t.days) + ' days'
	when t.flags & 37376 = 37376 then convert(varchar,t.hours) + ' hours'
	when t.flags & 18944 = 18944 then convert(varchar,t.minutes) + ' minutes' end as 'everytime',
case when t.flags & 74752 = 74752 then convert(varchar,t.days) + ' days'
	when t.flags & 37888 = 37888 then convert(varchar,t.hours) + ' hours'
	when t.flags & 19456 = 19456 then convert(varchar,t.minutes) + ' minutes' end as 'aftertime',
t.lastexecuted
from hsi.lctimer t
left outer join hsi.lcstatextimer x on x.timernum = t.timernum
left outer join hsi.lcstate s on s.statenum = x.statenum
left outer join hsi.lifecycle l on l.lcnum = t.scope
where t.flags & 67108864 <> 67108864
order by 1,2,3

"@
		Output = "$OutputDir\Legacy_Timers.xlsx"
	},
	@{
		Query  = @"
SELECT CASE
	WHEN f.workerpoolname IS NULL THEN '<Unassigned>'
	ELSE RTRIM (f.workerpoolname)
	END AS TaskGroup,
RTRIM (c.schedtaskname) AS JobName,
CASE
	WHEN c.schedtasktype = '1' then 'Workflow Timer'
	WHEN c.schedtasktype = '9' then 'Dashboards Data Provider Export'
	WHEN c.schedtasktype = '10' then 'Unity Script'
	WHEN c.schedtasktype = '22' then 'EIS Workflow Messaging Archiver'
	WHEN c.schedtasktype = '23' then 'EIS Workflow Messaging Cleaner'
	WHEN c.schedtasktype = '30' then 'Purge Execution History'
	WHEN c.schedtasktype = '34' then 'Capture Process'
	WHEN c.schedtasktype = '40' then 'Platter Deletion Processing'
	WHEN c.schedtasktype = '41' then 'Disk Group Analysis Processing'
	WHEN c.schedtasktype = '42' then 'Incomplete Commit Queue Processing'
	WHEN c.schedtasktype = '43' then 'Incomplete Delete Queue Processing'
	ELSE 'Other'
	END AS SchedulerType,
CASE
	WHEN b.schedulename IS NULL THEN 'Custom/Ad-Hoc Schedule'
	ELSE RTRIM (b.schedulename)
	END AS ScheduleName,
case
	when g.SCHEDULETYPE = '1' then convert(varchar,g.REPEATPERIOD) + ' Minute'
	when g.SCHEDULETYPE = '2' then convert(varchar,g.REPEATPERIOD) + ' Hour'
	when g.SCHEDULETYPE = '3' then convert(varchar,g.REPEATPERIOD) + ' Day (' + CONVERT(VARCHAR(5), g.EXECUTEAT, 8) + ')'
	when g.SCHEDULETYPE = '4' then convert(varchar,g.REPEATPERIOD) + ' Business Day'
	when G.SCHEDULETYPE = '5' then 'Weekly (' + case when G.DAYSOFWEEK & 1 > 0 then 'Su' else '' end +
	case when G.DAYSOFWEEK & 2 > 0 then 'M' else '' end +
	case when G.DAYSOFWEEK & 4 > 0 then 'Tu' else '' end +
	case when G.DAYSOFWEEK & 8 > 0 then 'W' else '' end +
	case when G.DAYSOFWEEK & 16 > 0 then 'Th' else '' end +
	case when G.DAYSOFWEEK & 32 > 0 then 'F' else '' end +
	case when G.DAYSOFWEEK & 64 > 0 then 'Sa' else '' end + ' ' + CONVERT(VARCHAR(5), g.EXECUTEAT, 8) + ')'
		when g.SCHEDULETYPE = '7' then 'Monthly Fixed Date'
		when g.SCHEDULETYPE = '9' then 'Monthly Relative'
		when g.SCHEDULETYPE = '10' then 'Annual'
		when g.SCHEDULETYPE = '11' then 'Full Calendar'
		else convert(varchar,g.scheduletype)
		end as ScheduleInterval,
a.latestexecutionstart AS LastRunStart,
CASE a.jobstatus
	WHEN 1 THEN 'New'
	WHEN 2 THEN 'Disabling'
	WHEN 3 THEN 'OK'
	WHEN 4 THEN 'Running'
	WHEN 5 THEN 'Inactive'
	WHEN 6 THEN 'ERROR!'
	WHEN 7 THEN 'Disabled'
	WHEN 8 THEN 'Cancelling'
	WHEN 9 THEN 'Cancelled'
	ELSE 'UNKNOWN (' + convert(varchar,a.jobstatus) + ')'
	END AS JobStatus,
RTRIM(s.servername) AS 'servername'
FROM hsi.schedulertask c
LEFT OUTER JOIN hsi.schedulerjob a ON a.schedtasknum = c.schedtasknum
LEFT OUTER JOIN hsi.schedulerschedule b ON a.schedulenum = b.schedulenum
LEFT OUTER JOIN hsi.schedulertaskhistory e ON e.schedtasknum = a.schedtasknum 
											AND a.latestexecutionstart = e.schedexecutionstart 
											AND a.latestexecutionend = e.schedexecutionend
LEFT OUTER JOIN hsi.schedulerservice s ON s.schedservicenum = e.schedservicenum
LEFT OUTER JOIN hsi.schedulerworkerpool f ON a.workerpoolnum = f.workerpoolnum
LEFT OUTER JOIN hsi.scheduleitem g ON g.schedulenum = a.schedulenum
order by 1,2


"@
		Output = "$OutputDir\Unity_Scheduler_Tasks.xlsx"
	}
)
	# Execute all queries
	foreach ($q in $queries) {
		Write-Host "Running query and exporting to $($q.Output)"
		Run-Query -Query $q.Query -OutputFile $q.Output
	}

	Write-Host "All queries completed. Reports saved in $OutputDir"
	pause