# Installs MSRC PowerShell Module
if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator" ))
{
    Write-Host "You must run PowerShell 'As Administrator' to use this tool." -foregroundcolor YELLOW
    Write-Host ""
    Exit
    
}
else
{
    # create module file
    $modulePath = "C:\Program Files (x86)\WindowsPowerShell\Modules"
    $moduleName = "MsrcSecurityUpdates"
    $moduleContent = Get-Content "MsrcSecurityUpdates.psm1.txt"

    New-Item -Path $modulePath -Name $moduleName -ItemType Directory
    New-Item "$modulePath\$moduleName\$moduleName.psm1" -ItemType File
    Add-Content "$modulePath\$moduleName\$moduleName.psm1" $moduleContent
}
