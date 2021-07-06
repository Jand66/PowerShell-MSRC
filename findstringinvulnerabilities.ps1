### Install the module from the PowerShell Gallery (must be run as Admin)
# Install-Module -Name MsrcSecurityUpdates -force

### Download the March CVRF as an object
# $cvrfDoc = Get-MsrcCvrfDocument -ID 2021-Jun

### Get the Remediations of Type 'Workaround' (0)
#$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 0

### Get the Remediations of Type 'Mitigation' (1)
#$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 1

### Get the Remediations of Type 'VendorFix' (2)
#$cvrfDoc.Vulnerability.Remediations | Where Type -EQ 2 

# $cvrfDoc | Get-MsrcVulnerabilityReportHtml | out-file -FilePath c:\jun2021.html

# Select-String -pattern "bootloader" -path C:\jun2021.html

$months = @("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov", "Dec")
$years =@("2017","2018","2019","2020","2021")
$searchstring = "boot"
foreach($year in $years)
{
    foreach($month in $months)
    {
        $docID = $year+"-"+$month
        $infile = "c:\"+ $year + $month + ".html"
        $outfile = "c:\"+ $year + $month + "_" + $searchstring + ".txt"
        #$cvrfdoc = Get-MsrcCvrfDocument -ID $docID
        #$cvrfDoc | Get-MsrcVulnerabilityReportHtml | out-file -FilePath $filename
        $infile
        $result = Select-String -pattern $searchstring -path $infile 
        if ($result -ne $null)
        { 
         $result | out-file -FilePath $outfile
         $outfile
        }
    }
}

