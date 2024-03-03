
$PSDefaultParameterValues['out-file:width'] = 20000
Import-Module -Name "PnP.PowerShell"
Import-Module .\CustomFunctions.psm1 -Force    #-Force is used to make sure latest changes are imported
. .\Variables.ps1 -Force


$beforeText = ""
$afterText = "" 
$targetUser = $beforeText

$Site = $src.Prod

################################################

$i = 0

# Connect to the SharePoint site
$Connected = Get-PnPConnection
PnPConnectionCheck -Site $Site -Connected $Connected

#Get User from Sharepoint and all groups he is member of
$User = Get-PnPUser -Identity "i:0#.f|membership|$beforeText" 
Write-Host $User.Groups.Title
$Groups = $User.Groups
#$Groups = $Groups | Where-Object {$_.Title -NotMatch "Proprietari"} Abilita riga per mantenere accesso admin
$Groups = $Groups | Where-Object {$_.Title -match "JV-L_"} #Applica solo a gruppi lettori

#$PSScriptRoot Logs
$logfilePath = “.\Logs\ScriptGruppi\$($beforeText.Substring(0,$beforeText.IndexOf('@')) +'-'+$afterText.Substring(0,$afterText.IndexOf('@'))).Permissions.txt”
$Groups | Select-Object @{n="ID";e={$_.Id}},@{n="Title";e={$_.Title}} | Format-List | Out-File $logfilePath


#Add new user to all groups of old user
foreach ($_ in $Groups) {
    Add-PnPGroupMember -Group $_.Id -EmailAddress $afterText
    Write-Host $_.Title
}

#Check che utente sia stato aggiunto a tutti i gruppi dell'utente da rimuovere
$User2 = Get-PnPUser -Identity "i:0#.f|membership|$afterText"
$User2.Groups.Id | ForEach-Object {
    if ($User.Groups.Id -contains $_) {$i++ }
}
if($i -eq $User.Groups.Count) {$AllGroupsOk = "true"} else {
    $AllGroupsOk = "false"}


#Cancella vecchio utente da tutti i gruppi di cui è parte
  If ($AllGroupsOk -eq "true")  {
    foreach ($_ in $Groups) {
    Remove-PnPGroupMember -Group $_.Title -LoginName "i:0#.f|membership|$beforeText" 
    Write-Host $_.Title
    }
}

