<#
.Synopsis
    Gets MSRC Update Details for requested Month using supplied API
.DESCRIPTION
        Gets MSRC Update Details for requested Month using supplied API using https://www.powershellgallery.com/packages/MsrcSecurityUpdates/1.2/Content/MsrcSecurityUpdates.psm1
.EXAMPLE
    .\msrc-susan.ps1 -month march -api d76e0316202747e6a58d79d17efba49e
.EXAMPLE
    .\msrc-susan.ps1 -i      
#>
param (
    [Parameter(Mandatory=$false,Position=0)]
    [string]$month,
    [Parameter(Mandatory=$false,Position=1)]
    [string]$year,
    [Parameter(Mandatory=$false,Position=2)]
    [string]$api,
    [Parameter(Mandatory=$false,Position=3)]
    [string]$path,
    [Parameter(Mandatory=$false,Position=4)]
    [switch]$i
)
try{
Import-Module "C:\Program Files (x86)\WindowsPowerShell\Modules\MsrcSecurityUpdates\MsrcSecurityUpdates.psm1" -Global -ErrorAction Stop
}
catch{
    Write-Output "Module MsrcSecurityUpdates Not Found!"
} 
$pad = 12
$msrcAPIKeyLabel = "Using API".PadRight($pad," ")
$monLabel = "Using Month".PadRight($pad," ")
$yearLabel = "Using Year".PadRight($pad," ")
$msrcAPIKeyStored = "storedAPI.txt"
$date = [datetime]::Now
if($i)
{
    $monthOfInterest = Read-Host -Prompt "Enter Search Month eg, 2017-May"
    $msrcAPIKey = Read-Host -Prompt "Enter API Key"
}
else
{
    if(!($month))
    {
        $monthOfInterest = (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($date.Month)
    }
    else
    {
        $monthOfInterest = $month.SubString(0,3)
    }
    if(!($year))
    {
        $year = $date.Year  
    }
    if(!($api))
    {
        $storedAPI = Test-Path $msrcAPIKeyStored
        if(!($storedAPI))
        {
            $msrcAPIKey = Read-Host -Prompt "Please enter API"
            $file = New-Item $msrcAPIKeyStored -ItemType file
            Add-Content $file $msrcAPIKey
            Write-Output "API will be saved for use in future" 
        }
        else
        {
            $msrcAPIKey = Get-Content $msrcAPIKeyStored
        }
    }
    Write-Output "$monLabel : $monthOfInterest"
    Write-Output "$yearLabel : $year"
    Write-Output "$msrcAPIKeyLabel : $msrcAPIKey"
    $monthOfInterest = "$year-$monthOfInterest"  
}
if(!($path))
{
    $path = [environment]::CurrentDirectory
}
$outFile = "$path\MSRCMarSecurityUpdates-$monthOfInterest.html"
Write-Output "Result File : $outFile"
$currentYear = $date.Year
$currentMonth = (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($date.Month)
$splitMonth = $monthOfInterest.Split("-")
if(($splitMonth[1] -gt $currentMonth) -and ($splitMonth[0] -ge $currentYear))
{
    Write-Output "Date is in Future" 
}
else
{
    try{
    
        Get-MsrcCvrfDocument  -ID $monthOfInterest -ApiKey $msrcAPIKey -Verbose -ErrorAction Stop | Get-MsrcSecurityBulletinHtml -Verbose | Out-File $outFile -ErrorAction Stop
    }
    catch
    {
        $error = $_.exception.Message
        Write-Output "Error, Please try a different Year-Month"
        Write-Output "$error"
    }

}
