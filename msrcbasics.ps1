  
### Install the module from the PowerShell Gallery (must be run as Admin)
Install-Module -Name MsrcSecurityUpdates -force
Import-Module MsrcSecurityUpdates

### Download the March CVRF as an object
$cvrfDoc = Get-MsrcCvrfDocument -ID 2021-Jun

### Get the Remediations of Type 'Workaround' (0)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 0

### Get the Remediations of Type 'Mitigation' (1)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 1

### Get the Remediations of Type 'VendorFix' (2)
$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 2 

$cvrfDoc | Get-MsrcVulnerabilityReportHtml | out-file -FilePath c:\jun2021.html
