#Import-Module -Name "PnP.PowerShell"
#Import-Module .\CustomFunctions.psm1 -Force


New-Module -ScriptBlock {
    function CheckDataType {
        param (
            $list,
            $Column
        )
        $ColumnType = (Get-PnPField -List $list -Identity $Column).TypeAsString
        IF( $ColumnType -eq "User") {$Type = "User"}
            else {
                if($ColumnType -eq "Note" && $null -ne ($currentItemName.$Column | Select-String -Pattern "{" -AllMatches) ) {
                    $Type = "Json"  
                } 
                else {$Type = "String" }
                }
                Return $Type
            }       
        


    function FieldtoChange {
        param (
            $ColumnType
        )
        switch ($ColumnType) {
            "User" { "Email" }
            "String" {"Email"}
            "Json" {"Name"}
            }
        
        Return $Field
        }

    



    Function PnPConnectionCheck { 
        Param( 
            $Site,
            $Connected
        )
    If ($Connected.Url -imatch $Site ) { Write-Host "Already Connected"}
    else {
        Write-Host "Connecting"
        Connect-PnPOnline -Url $Site -UseWebLogin}
        # Copia nello script: 
        #$Connected = Get-PnPConnection 
        #PnPConnectionCheck -Site $Site -Connected $Connected
    } 




    Function ReplaceOldText{

        param (
            $colType,
            $Items,
            $beforeText,
            $afterText,
            $userList,
            $Column
        )

       $FixedItems = foreach ($currentItemName in $Items) {   
            switch ($colType) {
            ######Switch Condition User Column
                    "User" { 
                        Write-Host "User" 
                    $j = $currentItemName.$Column.Count 
                    $currentItemName.$Column = $currentItemName.$Column | ForEach-Object {
                                                                        $checkLista = If($_.Email -ne $null) {$_.Email}
                                                                        if($beforeText.Email -imatch $checkLista)  {
                                                                            if($_.Email -ne $null) {
                                                                            $_=  @{Email = $afterText.Email}
                                                                            $_
                                                                        }}
                                                                        else {
                                                                            $_
                                                                            $i++
                                                                        }
                                                                        if($i -eq $j)
                                                                        {$ignoredItems += $currentItemName 
                                                                           } 
                                                                           
                                                                }
                                                                $i = 0
                                }
            ######Switch Condition String
                    "String" { 
                        Write-Host "String"
                    if($currentItemName.$Column -in $ListaUtentiEni.Email) {
                    $currentItemName.$Column -replace $beforeText.Email, $afterText.Email
                } else {
                     }}
            ######Switch Condition Json
                    "Json" {
                        $currentItemName.$Column = $currentItemName.$Column | ConvertFrom-Json 
                        Write-Host "Json"  
                        $j = $currentItemName.$Column.Count
                        $currentItemName.$Column = $currentItemName.$Column | ForEach-Object { 
                                $z++
                                $checkLista = If($_.mail+$_.Email -eq $null) {$_.Name -replace 'i:0#.f\|membership\|', ''} else {$_.mail+$_.Email}
                            if(  $userList.Email -icontains $checkLista)  
                                        {
                                            if($_.mail -ne $null) {$_.mail = $_.mail -replace $beforeText.Email, $afterText.Email}
                                            if($_.Email -ne $null) {$_.Email = $_.Email -replace $beforeText.Email, $afterText.Email}
                                            if($_.Name -ne $null) {$_.Name = $_.Name -replace $beforeText.Email, $afterText.Email}
                                            if($_.Title -ne $null) {$_.Title = $_.Title -replace $beforeText.Nome, $afterText.Nome
                                                                    $_.Title = $_.Title -replace $beforeText.Cognome, $afterText.Cognome
                                                                    }
                                        }   
                                        else { 
                                            $i++
                                        }
                            if($_ -ne $null) {
                                 $_}
                        } 
                if($z -eq 1) {
                    $currentItemName.$Column = $currentItemName.$Column | ConvertTo-Json -Compress #| Out-String  
                    $currentItemName.$Column = "[$($currentItemName.$Column)]"
                   #$currentItemName.$Column = "[{0}]" -f ($currentItemName.$Column | ConvertTo-Json -Compress | Out-String )
                    } else {
                     $currentItemName.$Column = @($currentItemName.$Column) | ConvertTo-Json -Compress | Out-String   
                    }
                if($i -eq $j)
                 {$ignoredItems += $currentItemName 
                    } 
                    $i = 0
                    $z = 0
                    
            }}
                      $currentItemName = if($currentItemName.ID -notin $ignoredItems.ID) {$currentItemName}
                      $currentItemName
            }
     Return $FixedItems 
    } 
# $UpdatedItems = ReplaceOldText -Items $Items -colType $colType -beforeText $beforeText -afterText $afterText -ListaUtentiEni $ListaUtentiEni -Column $Column



function UpdateItems {

param (
    $colType,
    $UpdatedItems,
    $list,
    $Column
)

 switch ($colType) {
        "User" {
        foreach ($currentItemName in $UpdatedItems) {
            [int32]$TempID = $currentItemName.ID 
            Set-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column.Email}  #il -Batch non funziona con multi person column
            }
        }
        "Json" {
            foreach ($currentItemName in $UpdatedItems) {
                [int32]$TempID = $currentItemName.ID 
                Set-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column.Trim()} #-Batch $batch #il -Batch non funziona con multi person column
            }
        }
        "String" {
            foreach ($currentItemName in $UpdatedItems) {
                [int32]$TempID = $currentItemName.ID 
                Set-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column} #-Batch $batch #il -Batch non funziona con multi person column
            }
        }
    }
}



function BackupItems {   ############################################ BackUP

    param (
        $colType,
        $UpdatedItems,
        $list,
        $Column
    )
    
     switch ($colType) {
            "User" {
            foreach ($currentItemName in $UpdatedItems) {
                [int32]$TempID = $currentItemName.ID 
                Add-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column.Email}  #il -Batch non funziona con multi person column
                }
            }
            "Json" {
                foreach ($currentItemName in $UpdatedItems) {
                    [int32]$TempID = $currentItemName.ID 
                    Add-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column.Trim()} #-Batch $batch #il -Batch non funziona con multi person column
                }
            }
            "String" {
                foreach ($currentItemName in $UpdatedItems) {
                    [int32]$TempID = $currentItemName.ID 
                    Add-PnPListItem -List $list -Identity $TempID  -Values @{$Column = $currentItemName.$Column} #-Batch $batch #il -Batch non funziona con multi person column
                }
            }
        }
    }


    function LogItems {
        Param (
            $colType,
            $list,
            $Column,
            $beforeText,
            $afterText
        )

    switch ($colType) {
            "User" {
    $logfilePath = “.\Logs\BonificaUtente\$($list+'-'+$Column+'-'+$beforetext.Title+'-'+$afterText.Title).BonificaItemsLog.txt”
    $UpdatedItems | Select-Object @{n="ID";e={$_.ID}},@{n=$Column;e={$_.$Column.Email}} | Format-List | Out-File $logfilePath
            }
    }
}

} -Name PnPCustomFunctions