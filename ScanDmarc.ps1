Param(
    [Parameter(Mandatory)]
    [string]$F
)

Import-Module BurntToast
Import-Module 7Zip4Powershell
$temp = "C:\Temp\DMARC\" + $F
$out = "" + $HOME + "\Documents\DMARC\"
$fn = ($F | Select-String -Pattern '.*(?=.xml|.zip|.gz|.tar.gz)').Matches[0].Value
$fn = ($fn | Select-String -Pattern '.*(?=.xml|.zip|.gz|.tar.gz)').Matches[0].Value
$path = "" + $out + $fn + ".xml"
$domain = ($F | Select-String -Pattern '(?<=!)[\D\.]*(?=!)').Matches[0].Value
$service = ($F | Select-String -Pattern '^.*?(?=!)').Matches[0].Value

If ($F -like "*.tar.gz*")
{
    Expand-7Zip -ArchiveFileName $temp -TargetPath "C:\Temp\DMARC\"
    $temp = $temp.SubString(0,($temp.Length - 3))
}
Expand-7Zip -ArchiveFileName $temp -TargetPath $out

$fails = Select-String -path $path -pattern "(?<!soft)fail" | % {$_.LineNumber}

If ($fails -ne $null)
{
    New-BTContentBuilder | Add-BTText "DMARC Error for $domain","$service Detected a DMARC Failure on $domain on the following lines: $fails" -PassThru | Show-BTNotification
}