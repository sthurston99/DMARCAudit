Param(
    [Parameter(Mandatory)]
    [string]$F
)

Import-Module BurntToast
$temp = "C:\Temp\DMARC\" + $F
$out = "" + $HOME + "\Documents\DMARC\"
$fn = ($F | Select-String -Pattern '.*(?=.xml|.zip|.gz)').Matches[0].Value
$fn = ($fn | Select-String -Pattern '.*(?=.xml|.zip|.gz)').Matches[0].Value
$path = "" + $out + $fn + ".xml"
$domain = ($F | Select-String -Pattern '(?<=!)[\D\.]*(?=!)').Matches[0].Value
$service = ($F | Select-String -Pattern '^.*?(?=!)').Matches[0].Value
$ID = ($F | Select-String -Pattern '(?<=!)\d+(?=!)').Matches[0].Value

$param = @('x', "-o$out", "$temp")
7z @param

$fails = (Select-String -Path $path -Pattern "(?<!soft)fail").count

If ($fails -gt 0) {
    Add-Content -Path "$out\log.txt" -Value "$ID : $service reporting $domain failed $fails times"
    If ($fails -gt 25) {
        New-BTContentBuilder | Add-BTText "DMARC Error $ID for $domain", "$service Detected $fails DMARC Failure(s) on $domain" -PassThru | Show-BTNotification
    }
} Else {
    Add-Content -Path "$out\log.txt" -Value "$ID : $service reporting $domain has no issues"
}