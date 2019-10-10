<#
Author: Alex Willard
Created Date: 10/10/2019
Desc: Pulls from SQL Table to create, update, suspend user
#>



#Globals
$SQL = @{
    SERVER = 'smdsi-wh\SQLEXPRESS'
    DB = 'Datawarehouse'
    USERNAME = 'sa'
    PASSWORD = 'Master2019!'
}
$FDATE = Get-Date -Format "yyyy.MM.dd HH:mm"
$LOGGING = ".\Logs\$($FDATE).log"
$BASEOU = 'OU=DSI,DC=GTNC,DC=local'

#Settings
$ErrorActionPreference = 'Stop'

#Functions
function Check-User {
    param (
        $EmployeeId
    )
    [bool] $Exist = Get-ADUser -Filter {employeeID -eq $EmployeeId}
    return $Exist
}

function Get-SamFromLogon {
    param (
        $Logon
    )
    return $Logon.substring($Logon.indexOf('\') + 1,$Logon.Length - $Logon.indexOf('\') - 1)
}

#Start Logging
Try{
    Start-Transcript -Path "$LOGGING" -Force
}Catch{
    Write-Host "Transcript failed to start."
}

#SQL Connection String
Try 
{ 
    $SQLConnection = New-Object System.Data.SQLClient.SQLConnection 
    $SQLConnection.ConnectionString ="server=$($SQL.SERVER);database=$($SQL.DB);User ID = $($SQL.USERNAME);Password = $($SQL.PASSWORD);" 
    $SQLConnection.Open() 
    "['$(Get-Date -Format g)'] Connection to SQL established:"
} 
catch 
{ 
    "['$(Get-Date -Format g)'] Failed to connect SQL Server:"
} 

#Pulls all info from Datawarehouse.dbo.Employees
$SqlCommand = New-Object System.Data.SqlClient.SqlCommand
$SqlCommand.CommandText = "
--Users to Create
SELECT *,
(SELECT TOP 1 AdLogonName FROM Datawarehouse.dbo.Employees ee WHERE e.ReportsToId = ee.AdpId) as ManagerSam FROM Datawarehouse.dbo.Employees e
INNER JOIN Datawarehouse.dbo.JobTitles jt ON jt.JobTitleId = e.JobTitleId
INNER JOIN Datawarehouse.dbo.Locations l ON l.LocationId = e.LocationId
WHERE e.ActiveValue = 1 AND e.Created = 0

--Users to Suspend
SELECT * FROM Datawarehouse.dbo.Employees e
WHERE ActiveValue = 0 AND Created = 1

--Users to Update
SELECT *,
(SELECT TOP 1 AdLogonName FROM Datawarehouse.dbo.Employees ee WHERE e.ReportsToId = ee.AdpId) as ManagerSam FROM Datawarehouse.dbo.Employees e
INNER JOIN Datawarehouse.dbo.JobTitles jt ON jt.JobTitleId = e.JobTitleId
INNER JOIN Datawarehouse.dbo.Locations l ON l.LocationId = e.LocationId
WHERE ActiveValue = 1 AND Updated = 1"
$SqlCommand.Connection = $SQLConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCommand
$SqlDataset = New-Object System.Data.DataSet
$SqlAdapter.Fill($SqlDataset) | Out-Null

#Employees to create
$Users2Create = $SqlDataset.Tables[0]

foreach ($Data in $Users2Create) {
    #Check for User
    if(Check-User $Data.AdpId){
        "User: $($Data.FirstLast) already exists!"
        Continue
    }

    $Location = $Adp.Location.trim().replace(' ', '').replace(',', '')
    $OU = "OU=$Location,$BASEOU"

    #Check for OU
    if(!(Get-ADOrganizationalUnit -Filter {distinguishedName -eq $OU})){
        New-ADOrganizationalUnit -Name $Location -Path $BASEOU
        "OU: $OU created"
    }

    New-ADUser -Name $Data.FirstLast `
    -AccountPassword ($Data.AdPassword | ConvertTo-SecureString -AsPlainText -Force) `
    -Surname $Data.LastName `
    -StreetAddress $Data.Address1 `
    -State $Data.State `
    -UserPrincipalName $($Data.FirstName + '.' + $Data.LastName + '@Servicemasterdsioffice.onmicrsoft.com') `
    -PostalCode $Data.Zip `
    -MobilePhone $Data.MobilePhone `
    -OfficePhone $Data.DeskExtension `
    -City $Data.City `
    -PasswordNeverExpires 0 `
    -CannotChangePassword 0 `
    -DisplayName $Data.FirstLast `
    -EmailAddress $Data.Email `
    -SamAccountName Get-SamFromLogon -Logon $Data.AdAccountName `
    -EmployeeID $Data.AdpId `
    -Enabled 1 `
    -GivenName $Data.Firstname `
    -Manager Get-SamFromLogon -Logon $Data.ManagerSam `
    -Office $Data.City `
    -Path $OU `
    -OtherAttributes @{employeeNumber = $Data.EmployeeId;
    title = $Data.JobTitle;}
}


#Employees to Suspend
$Users2Suspend = $SqlDataset.Tables[1]

#Employees to update
$Users2Update = $SqlDataset.Tables[2]