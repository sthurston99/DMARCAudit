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
$ID = ($F | Select-String -Pattern '(?<=!)\d*(?=\..{2,3}$)').Matches[0].Value

If ($F -like "*.tar.gz*")
{
    Expand-7Zip -ArchiveFileName $temp -TargetPath "C:\Temp\DMARC\"
    $temp = $temp.SubString(0,($temp.Length - 3))
}
Expand-7Zip -ArchiveFileName $temp -TargetPath $out

$fails = (Select-String -path $path -pattern "(?<!soft)fail").count

If ($fails -gt 0) {
    Add-Content -Path "$out\log.txt" -Value "$ID : $service reporting $domain failed $fails times"
    If ($fails -gt 5) {
        New-BTContentBuilder | Add-BTText "DMARC Error $ID for $domain","$service Detected $fails DMARC Failure(s) on $domain" -PassThru | Show-BTNotification
    }
}