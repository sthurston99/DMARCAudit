Param(
    [Parameter(Mandatory)]
    [string]$F
)

Import-Module BurntToast
$temp = "C:\Temp\DMARC\" + $F
$fn = ($F | Select-String -Pattern '.*(?=.xml|.zip|.gz)').Matches[0].Value
$fn = ($fn | Select-String -Pattern '.*(?=.xml|.zip|.gz)').Matches[0].Value
$path = "C:\Users\Simon Thurston\Documents\DMARC\" + $fn + ".xml"
$domain = ($F | Select-String -Pattern '(?<=!)[\D\.]*(?=!)').Matches[0].Value
$service = ($F | Select-String -Pattern '^.*?(?=!)').Matches[0].Value

7z.exe x -o'C:\Users\Simon Thurston\Documents\DMARC\' $temp -aoa

$fails = Select-String -path $path -pattern fail | % {$_.LineNumber}

If ($fails -ne $null)
{
    New-BTContentBuilder | Add-BTText "DMARC Error for $domain","$service Detected a DMARC Failure on $domain on the following lines: $fails" -PassThru | Show-BTNotification
}