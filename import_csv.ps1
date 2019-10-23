<#
Description: Imports CSV to SQl

##Variables##
$Args[0]: Path of CSV
$Args[1]: Table name in Staging DB

Author: Alex Willard
Date: 10/17/2019
#>

#Globals
$SQL = @{
    SERVER = 'smdsi-wh\SQLEXPRESS'
    DB = 'Staging'
    USERNAME = 'sa'
    PASSWORD = 'Master2019!'
}

#Functions
function ToArray{
    begin{
        $Output = @()
    }process{
        $Results = $($_ | Get-Member -MemberType 'NoteProperty').Definition
        foreach($Result in $Results){
            $Output += $Result.Substring($Result.IndexOf("="), $Result.Length - $Result.IndexOf("=")).Replace("=", "")
        }
    }end{
        return, $Output
    }
}

#Setting Var
$BatchSize = 50000

#External Var
$CsvPath = $args[0]
$DestinationTable = if($args[1]){$args[1]}else{'Adp_ActiveDirectory'}

#Create CSV Object
Write-Host $CsvPath
$Csv = Import-Csv $CsvPath

#Prep SQL Connection
$ConnectionString ="server=$($SQL.SERVER);database=$($SQL.DB);User ID = $($SQL.USERNAME);Password = $($SQL.PASSWORD);" 
$BulkCopy = New-Object Data.SqlClient.SqlBulkCopy($ConnectionString, [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock)
$BulkCopy.DestinationTableName = $DestinationTable
$BulkCopy.BulkCopyTimeout = 0
$BulkCopy.BatchSize = $BatchSize

$DataTable = New-Object System.Data.DataTable


#Format Datatable
$Columns = $Csv | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name' 

foreach ($Column in $Columns) {
    $ColumnObj = New-Object System.Data.Datacolumn
    $ColumnObj.ColumnName = $Column
    $Null = $DataTable.Columns.Add($Column)

    $Map = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping ($Column, $Column)
    $BulkCopy.ColumnMappings.Add($Map) | Out-Null
}

#Insert CSV rows into Table
foreach($Line in $Csv){
    $Result = $Line | ToArray
    $Null = $DataTable.Rows.Add($Result)
}

#$DataTable.Rows | Format-Table

#Send2Server
$BulkCopy.WriteToServer($DataTable)
$DataTable.Clear()

$BulkCopy.Close(); $BulkCopy.Dispose()
$DataTable.Dispose()
