param
(
	[int]$showLastMonths=10,
	[string]$logDir="C:\admin\scripts",
	[string]$dateFormat="MMM-yyyy",

	[bool]$showErrors=$true,
	[bool]$launchReport=$false,
	[bool]$xmlOutput=$true,
	[bool]$csvOutput=$true,
	[bool]$chartFriendly=$false,
	[bool]$dailyChecks=$true,
	[bool]$capacityChecks=$false,
	[bool]$monthlyChecks=$false,
	[bool]$tapeExpiryOnly=$false,
	[bool]$emailReport=$false,
	[string]$site="",
	[string]$smtpServer,
	[string]$from,
	[string]$fromContactName,
	[string]$toName,
	[string]$replyTo,
	[string]$toAddresses,
	[string]$siteAddress,
	[string]$contactNumber,
	[string]$logProgressHere

)
# Call Veeam modules where all of the love lives.
Import-module .\veeamModules.psm1 -force

Add-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue
Set-Variable -Name scriptName -Value $($MyInvocation.MyCommand.name) -Scope Global
Set-Variable -Name logDir -Value $logDir -Scope Global
Set-Variable -Name dateFormat -Value $dateFormat -Scope Global

if ($logProgressHere)
{
	Set-Variable -Name logfile -Value $logProgressHere -Scope Global
} else {
	Set-Variable -Name logfile -Value $($($MyInvocation.MyCommand.name) -replace '.ps1','.log') -Scope Global
}

#Get-Date | Out-File $global:logFile
logThis -msg (printLongDateTime -date (Get-date))

$global:backupSpecs=$null
$global:backupSessions=$null

#############################################################################################################################
#
# 				M	A	I	N			C	O	U	R	S	E
#
#############################################################################################################################
logThis -msg  "Being @ $(Get-date)"
$reportDate = Get-Date -Format "dd-MM-yy"
$veeamServer=Get-VBRLocalhost

if ($tapeExpiryOnly)
{
	logThis -msg  "`t-> Getting list of Missing Expired Tapes"
	$node=@{}
	$node["Name"]= $veeamServer.RealName
	$node["Report Properties"]=@{}
	$node["Report Properties"]["Report Ran on"]=$reportDate
	$node["Server Information"]=$veeamServer
	$node["Report Properties"]["Type"] = "<Customer> - Expired Tape List for Recall"
	#$list = getExpiredTapes
	$list = getExpiringTapes -daysAhead 1 | Sort-Object Barcode
	if ($list)
	{
		$node["Missing Tapes"] = $list
		$list | Out-File $global:logFile -Append
		if ($emailReport)
		{
			#$node["Email Properties"]=
			$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
			$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
			$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
			$style = $style + "TD{border: 1px solid black; padding: 5px; }"
			$style = $style + "</style>"
			$precontent = @"
				<p>Hi $toName,</p>
				<p></p>
				<p>Please find below a list of Tapes (barcodes) marked for recall. There is a total of $(($list | Measure-Object).Count) tapes in the list.</p>
				<p></p>
				<p>Please send back the below tapes to the following address: <i><b>$siteAddress</b></i></p>
				<p></p>
"@
			$postContent = @"
				<p></p>
				<p>Regards</p>
				<p></p>
				<p>$fromContactName</p>
				<p>$from | $contactNumber | $siteAddress </p>
"@
			$mail = @{
				'smtpServer' = $smtpServer;
				'from' = $from;
				'replyTo' = $replyTo;
				'toAddresses' = $toAddresses;
				'subject' = $node["Report Properties"]["Type"];
				'body' = [string]$($node.'Missing Tapes' | ConvertTo-Html -Head $style -PreContent $precontent -PostContent $postContent);
				'fromContactName' = $fromContactName
			}
			sendEmail  @mail
			$isgood=$true
		}
	} else {
		logThis -msg "`t`tThere are no expired Tapes found for recall"
		$isgood=$false
	}
} else {
	$node=@{}
	$node["Name"]= $veeamServer.RealName
	$node["Report Properties"]=@{}
	$node["Report Properties"]["Report Ran on"]=$reportDate
	$node["Server Information"]=$veeamServer
	if ($dailyChecks)
	{
		$node["Report Properties"]["Type"] = "Daily Checks"
		logThis -msg  "Running $($node["Report Properties"]["Type"])"
		$node["Reporting Period (Months)"] = $showLastMonths
		logThis -msg  "`t-> Collecting Backup Repository Information"


		logThis -msg  "`t-> Getting list of Active Backups "
		#$node["Active Backups"] = showActiveBackups

		logThis -msg  "`t-> Getting list of Missing Expired Tapes"
		$node["Expiring Missing Tapes"] = getExpiredTapes

		#logThis -msg  "`t-> Getting list of Missing Expiring Tapes"
		#$node["Missing Expired Tapes"] = getExpiredTapes -daysAhead 0

		logThis -msg  "`t-> Getting list of Tapes in drive"
		#$node["Tapes in use"] = getTapesInUse

		logThis -msg  "`t-> Getting list of Tapes used in the last backups"
		$node["Tapes to Recall"] = getExpiredTapes  -daysAhead 1

		logThis -msg  "`t-> Getting Veeam Services Status"
		$node["Veeam Status"] = Get-VeeamServices -inputObj ausydsv07
		$isgood=$true

	}

	if ($monthlyChecks)
	{
		$node["Report Properties"]["Type"] = "Monthly Checks"
		logThis -msg  "Running $($node["Report Properties"]["Type"])"
		$node["Reporting Period (Months)"] = $showLastMonths
		logThis -msg  "`t-> Collecting Backup Repository Information"
		$node["Repositories"] = Get-VeeamBackupRepositoryies -chartFriendly $chartFriendly
		logThis -msg  "`t-> Collecting Backup Sessions"
		$node["Backup Sessions"] = Get-VeeamBackupSessions -chartFriendly $chartFriendly
		logThis -msg  "`t-> Collecting Individual Backup Tasks"
		$node["Backups by Clients"] = Get-VeeamClientBackups -chartFriendly $chartFriendly

		# Get first day, last day
		$firstRecordedBackupDay = $node["Backup Sessions"]."Creation Time" | Sort-Object |  select -First 1
		$node["Report Properties"]["Sample Start Date"]=$firstRecordedBackupDay
		$lastRecordedBackupDay = $node["Backup Sessions"]."Creation Time" | Sort-Object | Select-Object -Last 1
		$thisDate=(Get-Date -Format $dateFormat)

		$node["Report Properties"]["Sample End Date"]=$lastRecordedBackupDay
		$reportingMonths = $node["Backup Sessions"].Month | ForEach-Object { get-date $_ } | Select-Object -Unique | Sort-Object | ForEach-Object { get-date $_ -format $dateFormat } | Where-Object {$_ -ne $thisDate} | Select-Object -Last $node["Reporting Period (Months)"]
		logThis -msg  "`t-> Creating Backup Jobs Summary"
		$node["Jobs Summary"] = Get-BackupJobsSummary -chartFriendly $chartFriendly -reportingMonths $reportingMonths

		logThis -msg  "`t->Client Sizes"
		$node["Client Sizes"] = Get-VeeamClientBackupsSummary-ClientInfrastructureSize -chartFriendly $chartFriendly -reportingMonths $reportingMonths

		logThis -msg  "`t->Data Change Rate"
		$node["Data Change Rate"] = Get-VeeamClientBackupsSummary-ChangeRate -chartFriendly $chartFriendly -reportingMonths $reportingMonths

		logThis -msg  "`t->Data Transfered"
		$node["Data Transfered"] = Get-VeeamClientBackupsSummary-DataIngested -chartFriendly $chartFriendly -reportingMonths $reportingMonths

		logThis -msg  "`t-> Creating Monthly Capacity Summary"
		$node["Monthly Capacity"] = Get-MonthlyBackupCapacity -chartFriendly $chartFriendly -reportingMonths $reportingMonths

		logThis -msg  "`t-> Getting list of Missing Expired Tapes"
		$node["Missing Tapes"]= getExpiredTapes
		$isgood=$true
	}
}


####################
if ($isgood)
{
	logThis -msg "Writing Report to Disks @ $logDir"
	#$prefix="$logDir\$reportDate-$($node['Name'])"
	$prefix="$logDir\$($node["Report Properties"]["Type"] -replace ' ','_')"
	if ((Test-Path -path $prefix) -ne $true)
	{
		New-Item -type directory -Path $prefix
	}
	if ($xmlOutput)
	{
		$node | Export-Clixml "$prefix\$reportDate-$($node['Name'])-$($node["Report Properties"]["Type"] -replace ' ','_').xml"
	}
	if ($csvOutput)
	{
		$exclusions="Report Properties","Name"
		$node.keys | ForEach-Object {
			#"Server Information","Repositories","Backup Sessions","Backups by Clients","Jobs Summary","Client Sizes","Data Change Rate","Data Transfered","Monthly Capacity","Missing Tapes" | ForEach-Object {
			$label=$_
			#$exclusions -notcontains "$label"
			if ($exclusions -notcontains "$label")
			{
				$node[$label]  | Export-Csv -NoTypeInformation "$prefix\$label.csv"
				#$node[$label]  | Export-Csv -NoTypeInformation "$label.csv"
			}
		}
	}
}
logThis -msg  ""
logThis -msg  "Completed @ $(Get-date)"
