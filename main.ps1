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
    if($Logon -isnot [System.DBNull]){
        return $Logon.substring($Logon.indexOf('\') + 1,$Logon.Length - $Logon.indexOf('\') - 1)
    }
}

function Add-DelegateGSuiteAccount {
    param (
        $UserEmail,
        $DelegateEmail
    )
    gam $UserEmail delegate to $DelegateEmail
}

function Set-SqlCreated {
    param (
        $AdpId
    )
    "User: $($Data.FirstLast) already exists!"
    $SqlCommand.CommandText = "
    UPDATE Employees
    SET Created = 1
    WHERE AdpId = '$AdpId'
    "
        $SqlCommand.ExecuteNonQuery() | Out-Null
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
        Set-SqlCreated -AdpId $Data.AdpId
        Continue
    }

    $Location = $Data.Location.trim().replace(' ', '').replace(',', '')
    $OU = "OU=$Location,$BASEOU"

    #Check for OU
    if(!(Get-ADOrganizationalUnit -Filter {distinguishedName -eq $OU})){
        New-ADOrganizationalUnit -Name $Location -Path $BASEOU
        "OU: $OU created"
    }


    $User = New-ADUser -Name $($Data.FirstLast) `
    -AccountPassword $($Data.AdPassword | ConvertTo-SecureString -AsPlainText -Force) `
    -Surname $($Data.LastName) `
    -StreetAddress $($Data.Address1) `
    -State $($Data.State) `
    -UserPrincipalName $($Data.FirstName + '.' + $Data.LastName + '@Servicemasterdsioffice.onmicrsoft.com') `
    -PostalCode $($Data.Zip) `
    -MobilePhone $($Data.MobilePhone) `
    -OfficePhone $($Data.DeskExtension) `
    -City $($Data.City) `
    -PasswordNeverExpires 0 `
    -CannotChangePassword 0 `
    -DisplayName $($Data.FirstLast) `
    -EmailAddress $($Data.Email) `
    -SamAccountName $(Get-SamFromLogon -Logon $Data.AdLogonName) `
    -EmployeeID $($Data.AdpId) `
    -Enabled 1 `
    -GivenName $($Data.Firstname) `
    -Manager $(Get-SamFromLogon -Logon $Data.ManagerSam) `
    -Office $($Data.City) `
    -Path $OU `
    -OtherAttributes @{employeeNumber = $($Data.EmployeeId);
    title = $($Data.JobTitle;)} 

    if($Data.JobTitle -eq 'Account Manager'){
        #Gives RDP Access to all servers 
        $Group = Get-ADGroup -Identity 'CN=Server Access_S,OU=Server Access,OU=Groups,DC=GTNC,DC=local'
        Add-ADGroupMember -Identity $Group -Members $User
    }
    if(Check-User -EmployeeId $Data.AdpId){
        Set-SqlCreated -AdpId $Data.AdpId
        "$($Data.AdLogonName) Created!"
    }else{
        "$($Data.AdLogonName) Failed Creation"
    }
}

exit

#Employees to Suspend
$Users2Suspend = $SqlDataset.Tables[1]

foreach ($Data in $Users2Suspend) {
    Get-ADUser $(Get-SamFromLogon -Logon $Data.AdLogonName) | Set-ADUser -Enabled 0
    
}

#Disable users that aren't matched
<#
$SqlCommand.CommandText = "SELECT * FROM Datawarehouse.dbo.Employees"
$SqlAdapter.SelectCommand = $SqlCommand
$SqlTempDataset = New-Object System.Data.DataSet
$SqlAdapter.Fill($SqlTempDataset)

$Results = $SqlTempDataset.Tables[0]
$Users = Get-ADUser -Filter * -SearchBase $BASEOU

foreach($User in $Users){
    
}
#>



#Employees to update
$Users2Update = $SqlDataset.Tables[2]

foreach ($Data in $Users2Update) {
    $Location = $Data.Location.trim().replace(' ', '').replace(',', '')
    $OU = "OU=$Location,$BASEOU"

    #Check for OU
    if(!(Get-ADOrganizationalUnit -Filter {distinguishedName -eq $OU})){
        New-ADOrganizationalUnit -Name $Location -Path $BASEOU
        "OU: $OU created"
    }

    $User = Get-ADUser -Filter {employeeId -eq $Data.AdpId}
    $User | Set-ADUser -DisplayName $Data.FirstLast `
    -GivenName $Data.FirstName `
    -Surname $Data.LastName `
    -StreetAddress $Data.Location `
    -City $Data.City `
    -State $Data.State `
    -PostalCode $Data.Zip `
    -Manager $(Get-ADUser -Filter {employees -eq $Data.ReportsToId}) `
    -Title $Data.JobTitle `
    -MobilePhone $Data.MobilePhone 

    Move-ADObject -Identity $User.DistinguishedName -TargetPath $OU
}

$SQLConnection.Close()

#Init Google Directory Sync

#End Logging
Try{
    Stop-Transcript 
}Catch{
    Write-Host "Transcript failed to stop."
}