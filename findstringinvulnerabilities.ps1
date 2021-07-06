$months = @("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov", "Dec")
$years =@("2017","2018","2019","2020","2021")
# testvalues limited to save time
# $months = @("Jan","Feb","Mar","Apr","May","Jun")
# $years =@("2021")
$searchstring = "tpm"
$rootpath = "c:\CVEDocs\"

Function createfiles
#create inputfiles - this could take some time
{
    Param(
        [parameter(Mandatory=$true)][Array] $years, 
        [parameter(Mandatory=$true)][Array] $months, 
        [parameter(Mandatory=$true)][String] $rootpath)
    foreach($year in $years)
    {
        foreach($month in $months)
        {
            $docID = $year+"-"+$month
            $infile = $rootpath + $year + $month + ".html"
            $cvrfdoc = Get-MsrcCvrfDocument -ID $docID
            $cvrfDoc | Get-MsrcVulnerabilityReportHtml | out-file -FilePath $infile
            $infile
        }
    }
}

Function analysefiles
#create outputfiles that contain the lines where searchstring was found
{
    Param(
        [parameter(Mandatory=$true)][Array] $years, 
        [parameter(Mandatory=$true)][Array] $months, 
        [parameter(Mandatory=$true)][String] $rootpath, 
        [parameter(Mandatory=$true)][String] $searchstring
        )
    foreach($year in $years)
    {
        foreach($month in $months)
        {
            $docID = $year+"-"+$month
            $infile = $rootpath + $year + $month + ".html"
            $outfile = $rootpath + $year + $month + "_" + $searchstring + ".txt"
            $infile
            $result = Select-String -pattern $searchstring -path $infile 
            if ($result -ne $null)
            { 
             $result | out-file -FilePath $outfile
             $outfile
            }
        }
    }
}
createfiles -years $years -months $months -rootpath $rootpath
analysefiles -years $years -months $months -rootpath $rootpath -searchstring $searchstring 
