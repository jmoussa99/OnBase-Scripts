/* database settings at a glance */
select database_id, name, user_access_desc, compatibility_level, recovery_model_desc, page_verify_option_desc, is_auto_create_stats_on, is_auto_update_stats_on, is_auto_close_on, is_auto_shrink_on
from sys.databases

/* when was the last full database backup */
select d.name, MAX(b.backup_finish_date) as last_backup_finish_date
from master.sys.databases d
left outer join msdb.dbo.backupset b on d.name = b.database_name
and b.type = 'D'
where d.database_id not in (2,3)
group by d.name

/* where are the backups going */
select top 100 physical_device_name, media_set_id
from msdb.dbo.backupmediafamily
order by 2 desc

/* who is using outdated odbc drivers */
/* NOTE: edit the last line to match the current database version */
SELECT A.session_id
 , B.login_name
 , B.host_name
 , A.client_net_address
 , B.client_interface_name
 , A.protocol_type
 , CAST(A.protocol_version AS VARBINARY(9))
 ,driver_version =
 CASE SUBSTRING(CAST(A.protocol_version AS BINARY(4)), 1,1)
 WHEN 0x70 THEN 'SQL Server 7.0'
 WHEN 0x71 THEN 'SQL Server 2000'
 WHEN 0x72 THEN 'SQL Server 2005'
 WHEN 0x73 THEN 'SQL Server 2008'
 WHEN 0x74 THEN 'SQL Server 2012+'
 ELSE 'Unknown driver'
 END
 FROM sys.dm_exec_connections A
 INNER JOIN sys.dm_exec_sessions B ON A.session_id = B.session_id
 WHERE B.client_interface_name = 'ODBC'
 AND SUBSTRING(CAST(A.protocol_version AS BINARY(4)), 1,1) <> 0x74


Database level queries
--------------------------------------
/* run this first to prevent table locking */
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

/* database (type 0) and transaction log (type 1) size */
select type, SUM((size*8)/1024) as MB
from sys.database_files
group by type


/* overview counts */
select 'LifeCycles' as 'Category', count(*) as 'Count'from hsi.lifecycle
union all
select 'Queues' as 'Category', count(*) as 'Count' from hsi.lcstate


/* peak license usage over last 90 days */
/* Config | Utils | Core-Based Settings | Log License Usage must be enabled for this report. */
select rtrim(ps.productname) as ProductName, max(lu.usagecount) as MaxUsage
FROM hsi.licusage lu
INNER JOIN hsi.productsold ps ON lu.producttype = ps.producttype
where lu.logdate > DATEADD(dd, -90, getdate())
GROUP BY ps.productname
ORDER BY ps.productname


/* modules licensed and registered */
select rtrim(p.productname) as 'product', l.licensecount, p.regcopies, l.lpexpirationdate
from hsi.productsold p
inner join hsi.licensedproduct l on p.producttype = l.producttype
where p.numcopieslic > 0
order by 1


/* disk group summary */
select rtrim(d.diskgroupname) as 'diskgroup', 
d.numberofbackups as 'copies', d.ucautopromotespace as 'size', 
d.lastlogicalplatter as 'volumes', rtrim(p.lastuseddrive) as 'path'
from hsi.diskgroup d
inner join hsi.physicalplatter p on d.diskgroupnum = p.diskgroupnum
order by 1


/* number of docs */
select count(itemname) from hsi.itemdata


/* number of files */
select count(distinct filepath) from hsi.itemdatapage


/* WorkView Objects in OnBase */
SELECT case activestatus
	when 1 then 'Inactive'
	when 2 then 'Deleted'
	else 'Active' end as 'status',
count(DISTINCT objectid) AS 'count'
from hsi.rmObject
GROUP BY activestatus
order by 1


/* Inactive and Deleted WorkView Objects in Unity Life Cycles */
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


/* Number of Unindexed Documents */
select count(itemname) from hsi.itemdata where itemtypenum = 0


/* Number of Documents per Document Type */
select dt.itemtypenum as 'DocTypeID', 
rtrim(dt.itemtypename) as 'DocumentType', count(id.itemname) as 'Count' 
from hsi.doctype dt 
left outer join hsi.itemdata id
on dt.itemtypenum = id.itemtypenum
group by dt.itemtypenum, dt.itemtypename
order by dt.itemtypename


/* User Groups Not in Use */
select rtrim(ug.usergroupname) as 'name'
from hsi.usergroup ug
left outer join hsi.userxusergroup uxg on ug.usergroupnum = uxg.usergroupnum
where uxg.usergroupnum is null
order by 1


/* Security Overrides */
select 'User Group' as 'type', rtrim(g.usergroupname) as 'item', rtrim(k.keytype) as 'securedby'
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


/* User Accounts Not in Any Groups*/
select rtrim(ua.username) as 'name'
from hsi.useraccount ua
left outer join hsi.userxusergroup uxg on ua.usernum = uxg.usernum
where uxg.usernum is null
and ua.licenseflag & 2 <> 2
and ua.licenseflag & 65536 <> 65536
order by 1


/* Notable User Account Statistics */
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


/* Notable User Account Settings */
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

/* Documents Ingested by Method */
select 'Ad Hoc Import' as 'type', COUNT(i.itemnum) as 'count'
from hsi.itemdata i
where i.batchnum = 0
and i.datestored > DATEADD(dd, -90, getdate())
group by i.batchnum
UNION ALL
select 'Scan / Sweep' as 'type', SUM(s.extrainfo1) as 'count'
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
       end  as 'type',
COUNT(i.itemnum) as 'count'
FROM hsi.itemdata i, hsi.parsedqueue p
        WHERE i.batchnum = p.batchnum
        AND (p.parsingmethod & 4 = 4 or p.parsingmethod in (1,41,73,74,122))
        AND i.datestored > DATEADD(dd, -90, getdate())
group by p.parsingmethod


/* Autofill Keyword sets */
select s.keysettablenum, rtrim(s.keysetname) as 'keysetname', rtrim(t.keytype) as 'primary',
case when s.flags & 32 = 32 then 'external' else 'internal' end as 'source'
from hsi.keywordset s
inner join hsi.keysetxkeytype x on s.keysettablenum = x.keysettablenum
join hsi.keytypetable t on x.keytypenum = t.keytypenum
where x.seqnum = 0
order by 2


/* Notable Keyword Types */
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


/* VB Scripts */
select t.vbscriptnum, t.vbscriptname, t.vbscript, u.usergroupname from hsi.vbscripttable t
left outer join hsi.vbscripthooks h on t.vbscriptnum = h.vbscriptnum
left outer join hsi.usergroup u on h.usergroupnum = u.usergroupnum
order by t.vbscriptnum, u.usergroupname


/* Queue Counts */
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


/* Admin Processing Privileges */
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


/* Notable Queue Settings */
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


/* Configured Logging */
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


/* Workflow Notifications */
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


/* WorkView Notifications */
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


/* Legacy Timers */
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
where t.flags & 67108864 <> 67108864 --exclude unity timers
order by 1,2,3


/* Unity Scheduler Tasks */
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
