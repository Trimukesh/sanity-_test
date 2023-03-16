$debug = $true;
#Add-PSSnapin SqlServerCmdletSnapin110
#Add-PSSnapin SqlServerProviderSnapin110
cls
write-host $servernames
$starttime = Get-Date
write-host $starttime

#[string[]]$servernames = Get-Content -Path D:\\scripts\\ServersList.txt
[string[]]$servernames = $($env:servernames)
foreach($servername in $servernames)
{
write-host Starting Server $servername
$dataSource  =  $servername -replace "/","\\"
##setup data source

$database = "master"                                 ##Database name
$TableHeader = "SQL Server Health Check Report"      ##The title of the HTML page
$path = "D:\\"
$name = $dataSource -replace "\\\\","_"
$OutputFile_new = $path + $name + \'.html\'             ##The file location

$a = "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TABLE{width: 70%;border-width: 1px;border-style: solid;border-color: #2F4F4F;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: #2F4F4F;;background-color:#8FBC8F}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: #2F4F4F;}"
$a = $a + "</style>"

$colorTagTable = @{
                    Stopped = \' bgcolor="RED">Stopped<\';
					Read_Write = \' bgcolor="Green">Read_Write<\';
                    Read_Only = \' bgcolor="Red">Read_Only<\';
                    Running = \' bgcolor="Green">Running<\';
                    OFFLINE = \' bgcolor="RED">OFFLINE<\';
                    ONLINE  = \' bgcolor="Green">ONLINE<\';
					RESTORING = \' bgcolor="RED">RESORING<\';
					RECOVERING = \' bgcolor="RED">RECOVERING<\';
					RECOVERY_PENDING = \' bgcolor="RED">RECOVERY PENDING<\';
					SUSPECT = \' bgcolor="RED">SUSPECT<\';
					EMERGENCY = \' bgcolor="RED">EMERGENCY<\';
                
                   } 


$userid= $($env:username)
$Password=$($env:password)
#$userid= $cred.username
#$password=$cred.GetNetworkCredential().password
##Create a string variable with all our connection details 
#$connectionDetails = "Provider=sqloledb; " + "Data Source=$dataSource; 
#" + "Initial Catalog=$database; " + "Integrated Security=SSPI;" 
$connectionDetails ="Persist Security Info=False;" + "User ID= $userid;" + "Password= $Password;" + "Initial Catalog=Master;" + "Data Source=$dataSource;" 


##**************************************
##Calculating SQL Server Information
##**************************************
$sql_server_info = "select @@servername as [SQLNetworkName], 
CAST( SERVERPROPERTY(\'MachineName\') AS NVARCHAR(128)) AS [MachineName],
CAST( SERVERPROPERTY(\'ServerName\')AS NVARCHAR(128)) AS [SQLServerName],
serverproperty(\'edition\') as [Edition],
serverproperty(\'productlevel\') as [Servicepack],
CAST( SERVERPROPERTY(\'InstanceName\') AS NVARCHAR(128)) AS [InstanceName],
SERVERPROPERTY(\'Productversion\') AS [ProductVersion],@@version as [Serverversion]"

##Connect to the data source using the connection details and T-SQL command we provided above, 
##and open the connection
$connection = New-Object System.Data.SqlClient.SqlConnection $connectionDetails
$command1 = New-Object system.data.sqlclient.sqlcommand $sql_server_info,$connection
$connection.Open()

##Get the results of our command into a DataSet object, and close the connection
$dataAdapter = New-Object System.Data.sqlclient.sqlDataAdapter $command1
$dataSet1 = New-Object System.Data.DataSet
$dataAdapter.Fill($dataSet1)
$connection.Close()


##Return all of the rows and pipe it into the ConvertTo-HTML cmdlet, 
##and then pipe that into our output file
$frag1 = $dataSet1.Tables | Select-Object -Expand Rows |select -Property MachineName,SQLServerName,Edition,InstanceName,Serverversion  | ConvertTo-HTML -AS Table -Fragment -PreContent \'<h3 style="color:#2F4F4F">SQL Server Info</h3>\'|Out-String


write-host $frag1


##**************************************
##Database states
##**************************************
$SQLServerDatabaseState = "
IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N\'tempdb..#tmp_database\'))
BEGIN
drop table #tmp_database
END

declare @count int
declare @name varchar(128)
declare @state_desc varchar(128)

select @count = COUNT(*) from sys.databases where state_desc not in (\'ONLINE\')
create table #tmp_database (name nvarchar(128),state_desc nvarchar(128))
if @count > 0
        begin
            Declare Cur1 cursor for select name,state_desc from sys.databases 
            where state_desc not in (\'ONLINE\')
        open Cur1
            FETCH NEXT FROM Cur1 INTO @name,@state_desc
            WHILE @@FETCH_STATUS = 0
                BEGIN
                    insert into #tmp_database values(@name,@state_desc)
                FETCH NEXT FROM Cur1 INTO @name,@state_desc
                END
            CLOSE Cur1
            DEALLOCATE Cur1
        end
else
    begin
        insert into #tmp_database values(\'ALL DATABASES ARE\',\'ONLINE\')
    end
if @count > 0
   begin
     Declare Cur2 cursor for
	 select name,state_desc from sys.databases 
            where state_desc not in (\'ONLINE\',\'OFFLINE\',\'RESTORING\',\'RECOVERING\',\'RECOVERY PENDING\',\'SUSPECT\',\'EMERGENCY\',\'OFFLINE\')
	        open Cur2
            FETCH NEXT FROM Cur2 INTO @name,@state_desc
            WHILE @@FETCH_STATUS = 0
                BEGIN
                    insert into #tmp_database values(@name,\'HUNG STATE\')
                FETCH NEXT FROM Cur2 INTO @name,@state_desc
                END
            CLOSE Cur2
            DEALLOCATE Cur2
    end

	     
select name as DBName ,state_desc as DBStatus from #tmp_database
"

$connection = New-Object System.Data.SqlClient.SqlConnection $connectionDetails
$command2 = New-Object system.data.sqlclient.sqlcommand $SQLServerDatabaseState,$connection
$connection.Open()

##Get the results of our command into a DataSet object, and close the connection
$dataAdapter = New-Object System.Data.sqlclient.sqlDataAdapter $command2
$dataSet2 = New-Object System.Data.DataSet
$dataAdapter.Fill($dataSet2)
$connection.Close()

$frag2 = $dataSet2.Tables | Select-Object -Expand Rows |Select -Property DBName,DBStatus | 
ConvertTo-HTML -AS Table -Fragment -PreContent \'<h3 style="color:#2F4F4F">SQLServer Databases State</h3>\'|Out-String

$colorTagTable.Keys | foreach { $frag2 = $frag2 -replace ">$_<",($colorTagTable.$_) }

write-host $frag2 


##**************************************
##Database Read-Write
##**************************************
$SQLServerDatabaseState = "
IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N\'tempdb..#tmp_database\'))
BEGIN
drop table #tmp_database
END

declare @count int
declare @name varchar(128)
declare @state_desc varchar(128)

select @count = COUNT(*) from sys.databases where name in (SELECT name FROM sys.databases where is_read_only=1)
create table #tmp_database (name nvarchar(128),state_desc nvarchar(128))
if @count > 0
        begin
            Declare Cur1 cursor for SELECT name,is_read_only FROM sys.databases where name in (SELECT name FROM sys.databases where is_read_only=1)
          open Cur1
            FETCH NEXT FROM Cur1 INTO @name,@state_desc
            WHILE @@FETCH_STATUS = 0
                BEGIN
                   
                    insert into #tmp_database values(@name,\'Read_Only\')
                FETCH NEXT FROM Cur1 INTO @name,@state_desc
                END
            CLOSE Cur1
            DEALLOCATE Cur1
        end
else 
    begin
        insert into #tmp_database values(\'ALL DATABASES ARE\',\'Read_Write\')
    end

select name as DBName ,state_desc as DBStatus from #tmp_database
"

$connection = New-Object System.Data.SqlClient.SqlConnection $connectionDetails
$command3 = New-Object system.data.sqlclient.sqlcommand $SQLServerDatabaseState,$connection
$connection.Open()

##Get the results of our command into a DataSet object, and close the connection
$dataAdapter = New-Object   System.Data.sqlclient.sqlDataAdapter $command3
$dataSet3 = New-Object System.Data.DataSet
$dataAdapter.Fill($dataSet3)
$connection.Close()

$frag3 = $dataSet3.Tables | Select-Object -Expand Rows |Select -Property DBName,DBStatus | 
ConvertTo-HTML -AS Table -Fragment -PreContent \'<h3 style="color:#2F4F4F">SQLServer Databases Mode</h3>\'|Out-String

$colorTagTable.Keys | foreach { $frag3 = $frag3 -replace ">$_<",($colorTagTable.$_) }

write-host $frag3 

##**************************************
##SQL Server Services Status
##**************************************
$SQLServerDatabaseState = "
IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE ID = OBJECT_ID(N\'tempdb..#tmp_database\'))
BEGIN
drop table #tmp_database
END

declare @count int
declare @name varchar(128)
declare @state_desc varchar(128)

select @count = COUNT(*) from sys.dm_server_services where servicename in (select servicename from sys.dm_server_services where status_desc =\'Running\')
create table #tmp_database (name nvarchar(128),state_desc nvarchar(128))
if @count > 0
        begin
            Declare Cur1 cursor for select servicename,status_desc from sys.dm_server_services where status_desc =\'Running\'
          open Cur1
            FETCH NEXT FROM Cur1 INTO @name,@state_desc
            WHILE @@FETCH_STATUS = 0
                BEGIN
                    insert into #tmp_database values(@name,@state_desc)
                FETCH NEXT FROM Cur1 INTO @name,@state_desc
                END
            CLOSE Cur1
            DEALLOCATE Cur1
        end
else 
    begin
        insert into #tmp_database values(\'ALL SQL Server Services are\',\'Running\')
    end

select name as ServiceName ,state_desc as Status from #tmp_database
"

$connection = New-Object System.Data.SqlClient.SqlConnection $connectionDetails
$command4 = New-Object system.data.sqlclient.sqlcommand $SQLServerDatabaseState,$connection
$connection.Open()

##Get the results of our command into a DataSet object, and close the connection
$dataAdapter =  New-Object System.Data.sqlclient.sqlDataAdapter $command4
$dataSet4 = New-Object System.Data.DataSet
$dataAdapter.Fill($dataSet4)
$connection.Close()

$frag4 = $dataSet4.Tables | Select-Object -Expand Rows |Select -Property ServiceName,Status | 
ConvertTo-HTML -AS Table -Fragment -PreContent \'<h3 style="color:#2F4F4F">SQL Server Services Status</h3>\'|Out-String

$colorTagTable.Keys | foreach { $frag4 = $frag4 -replace ">$_<",($colorTagTable.$_) }

write-host $frag4 


##**************************************
##Final Code to Combine all fragments
##**************************************

ConvertTo-HTML -head $a -PostContent $frag1,$frag2,$frag3,$frag4 -PreContent \'<h1 style="color:#2F4F4F"><center><U>SQL Server Heatlh Check Report</U></center></h1>\'| Out-File $OutputFile_new

#$fromaddress = "Aparna.Sagi@teachforamerica.org" 
$fromaddress= $($env:From)
#$toaddress = "Aparna.Sagi@teachforamerica.org" 
$toaddress= $($env:To)
$Subject = "SQL Server Status Report($dataSource)" 
$body = Get-Content $OutputFile_new
#$attachment = "D:\\UWDBRSTR01.html" 
$smtpserver = "plsmtp1.prod.tfanet.org" 
  
 
$message = new-object System.Net.Mail.MailMessage 
$message.From = $fromaddress 
$message.To.Add($toaddress) 
#$message.CC.Add($CCaddress) 
#$message.Bcc.Add($bccaddress) 
$message.IsBodyHtml = $True 
$message.Subject = $Subject 
#$attach = new-object Net.Mail.Attachment($attachment) 
#$message.Attachments.Add($attach) 
$message.body = $body 
$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message) 

$Stoptime = Get-Date
Write-host $Stoptime

}

remove-variable starttime
remove-variable servernames
remove-variable servername
remove-variable dataSource
remove-variable database
remove-variable TableHeader
remove-variable path
remove-variable name
remove-variable OutputFile_new
remove-variable a
remove-variable colorTagTable
remove-variable connectionDetails
remove-variable connection
remove-variable dataAdapter
remove-variable sql_server_info
remove-variable Stoptime

remove-variable dataSet1
remove-variable dataSet2
remove-variable dataSet3
remove-variable dataSet4



remove-variable command1
remove-variable command2
remove-variable command3
remove-variable command4



remove-variable frag1
remove-variable frag2
remove-variable frag3
remove-variable frag4



'''
}
}
}
}
