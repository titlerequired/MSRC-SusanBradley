<#
# MSRC-SusanBradley
.Synopsis
    Gets MSRC Update Details for requested Month using supplied API
.DESCRIPTION
Project is for using the MSRC Tool provided here, https://www.powershellgallery.com/packages/MsrcSecurityUpdates/1.2/Content/MsrcSecurityUpdates.psm1 to accept user input to generate a report.
To use, download the files. Run the Intall-MSRCUpdateTool.ps1 from an elevated PoweShell Window.
Then use the msrc-susan.ps1 file to download the reports.
.EXAMPLE
    .\msrc-susan.ps1 -month march -api d76e0316202747e6a58d79d17efba49e
.EXAMPLE
    .\msrc-susan.ps1 -i      
#>
