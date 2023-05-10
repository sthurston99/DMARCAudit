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
$ID = ($F | Select-String -Pattern '(?<=!)\d+(?=!)').Matches[0].Value

$param = @('x', "-o$out", "$temp")
7z @param

[xml]$xml = (Get-Content -Path $path)
$dkimfails = ($xml.feedback.record.auth_results.dkim.result | Where-Object { $_ -eq "fail" } | Measure-Object).count
$spffails = ($xml.feedback.record.auth_results.spf.result | Where-Object { $_ -eq "fail" } | Measure-Object).count
$fails = $dkimfails + $spffails

$service = $xml.feedback.report_metadata.org_name
$domain = $xml.feedback.policy_published.domain

If ($fails -gt 0) {
	Add-Content -Path "$out\log.txt" -Value "$ID : $service reporting $domain failed $fails times"
	If ($fails -gt $Env:FailThresh) {
		New-BTContentBuilder | Add-BTText "DMARC Error $ID for $domain", "$service Detected $fails DMARC Failure(s) on $domain" -PassThru | Show-BTNotification
	}
} Else {
	Add-Content -Path "$out\log.txt" -Value "$ID : $service reporting $domain has no issues"
}