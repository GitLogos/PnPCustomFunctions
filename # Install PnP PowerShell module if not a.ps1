# Install PnP PowerShell module if not already installed
if (-not (Get-Module -Name "PnP.PowerShell" -ListAvailable)) {
    Install-Module -Name "PnP.PowerShell" -Force -Scope CurrentUser
}

# Import the PnP PowerShell module
Import-Module -Name "PnP.PowerShell"

