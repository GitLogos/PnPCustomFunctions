
Import-Module -Name "PnP.PowerShell"
Import-Module .\CustomFunctions.psm1 -Force
. .\Variables.ps1 -Force
$PSDefaultParameterValues['out-file:width'] = 20000


$beforeText = @{Email = ""
                   Title = ""
                   Nome = ""
                   Cognome = ""}

$afterText = @{Email = '' 
                    Title = ''
                    Nome = ""
                    Cognome = ""}

$targetUser = $beforeText

IF($null -ne $targetUser) {$userList = $targetUser} # $targetUser is used when replacing one user with another in a specific column, userList is used when checking if a value is part of an array (imported from CSV file in variables.ps1)

#################################################
$Site = $src.Prod
$list = $Lists.List1
$Column = $Colonne.Title2

##################################################

# Connect to the SharePoint site
#Connect-PnPOnline -Url $Site -UseWebLogin
$Connected = Get-PnPConnection
PnPConnectionCheck -Site $Site -Connected $Connected


#custom functions
$colType = CheckDataType -list $list -Column $Column
$field = FieldtoChange -ColumnType $colType
if($colType -eq "User") {$ColumnEmail = "$Column.Email"} {$ColumnEmail = $Column} #depending on column type or content of the column. Simple string types will remain the same while person column or multi-line fields with JSON content will use user Email

#Get list Items
$Items = (Get-PnPListItem -List $list -PageSize 500 -Fields "ID", $Column).FieldValues
$Deleted = $Items | Where-Object {$_.$Column -eq $null}
$Items = $Items | Where-Object {$_.$Column -ne $null }
if($colType -ne "User") {$Items = $Items | Where-Object {$_.$Column.Trim() -ne '[]' -And $_.$Column -imatch $targetUser.Email}}
if($colType -eq "User") {$Items = $Items | Where-Object {$_.$Column.Email -imatch $targetUser.Email}} # only for person column

########################################################################
#Customfunction Check if the column contains value and replace old with new

$UpdatedItems = ReplaceOldText -Items $Items -colType $colType -beforeText $beforeText -afterText $afterText -userList $userList -Column $Column

######################################################################     

$DeletedCount = $UpdatedItems |  Where-Object {$_.$Column -eq $null}
$UpdatedItems = $UpdatedItems | Where-Object {$_.$Column -ne $null}


#$PSScriptRoot logging
$logfilePath = “.\Logs\BonificaUtente\$($list+'-'+$Column+'-'+$beforetext.Title+'-'+$afterText.Title).BonificaItemsLog.txt”
$UpdatedItems | Select-Object @{n="ID";e={$_.ID}},@{n=$Column;e={$_.$Column.Email}} | Format-List | Out-File $logfilePath

######################################################################
# Output the updated array
UpdateItems -colType $colType -UpdatedItems $UpdatedItems -list $list -Column $Column

        
