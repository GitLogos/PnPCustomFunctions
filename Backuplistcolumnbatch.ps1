
Import-Module -Name "PnP.PowerShell"
Import-Module .\CustomFunctions.psm1 -Force
. .\Variables.ps1 -Force

#source
$Site = $src.Prod
$list = $Lists.Title2

$column = $ColonneJV.Lettori
$backupList = ("Backup"+$list+$column) 

# Connect to the SharePoint site

PnPConnectionCheck -Site $Site
$colType = CheckDataType -list $list -Column $Column
$field = FieldtoChange -ColumnType $colType

$SchemaXML = (Get-PnPField -List $list -Identity $column).SchemaXml
$SchemaXML = $SchemaXML -replace 'ShowField','ShowAlways="TRUE" ShowField'


$Items = (Get-PnPListItem -List $list -Fields "ID", $Column).FieldValues
#$Items = $Items | Where-Object {$_.$Column -ne $null}

#Crea lista e colonna selezionata

New-PnPList  -Title $backupList -Template GenericList
Add-PnPFieldFromXml -List $backupList -FieldXml $SchemaXML 




# $batch = New-PnPBatch

    foreach ($currentItemName in $Items) {
        [int32]$TempID = $currentItemName.ID
        Add-PnPListItem -List $backupList -Values @{"Title" = $TempID ;$column = $currentItemName.$Column.Email} 
        }

#Invoke-PnPBatch -Batch $batch

