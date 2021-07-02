# This is a custom health check -- still a bit of work to do here to make it a bit more friendly
Add-PSSnapin "VeeamPSSnapIn" -ErrorAction SilentlyContinue

#region User-Variables
# Report mode - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = Monthly
# Report Title
$rptTitle = "My Veeam Report"
# Append Report Mode to Report Title E.g. My Veeam Report (Last 24 Hours)
$fullTitle = $true
# Report Width in Pixels
$rptWidth = 1024
# Only show last session for each Job
$onlyLast = $false

# Location of Veeam executable (Veeam.Backup.Shell.exe)
$veeamExePath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Shell.exe"
# Location of common dll - Needed for repository function - Get-vPCRepoInfo (Veeam.Backup.Core.dll)
$veeamDllPath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Core.dll"
# vCenter server(s) - As seen in VBR server
#$vcenters = "vcenter1","vcenter2"
$vcenters = "vCenterserver.internal.domain"

# Show VMs with no successful backups within time frame ($reportMode)
$showUnprotectedVMs = $true
# Show VMs with successful backups within time frame ($reportMode)
$showProtectedVMs = $true
# To Exclude VMs from Missing and Successful Backups section add VM names to be excluded
# $excludevms = @("vm1","vm2","*_replica")
$excludeVMs = @("")
# Exclude VMs from Missing and Successful Backups section in the following (vCenter) folder(s)
# $excludeFolder =  = @("folder1","folder2","*_testonly")
$excludeFolder = @("")
# Exclude VMs from Missing and Successful Backups section in the following (vCenter) datacenter(s)
# $excludeDC =  = @("dc1","dc2","dc*")
$excludeDC = @("")
# Show Running jobs
$showRunning = $true
# Show All Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFail = $true
# Show All Successful Sessions within time frame ($reportMode)
$showSuccess = $true
# Show Repository Info
$showRepo = $true
# Show Proxy Info
$showProxy = $true
# Show Replica Target Info
$showReplica = $true
# Show Veeam Services Info (Windows Services)
$showServices = $true
# Show only Services that are NOT running
$hideRunningSvc = $false
# Show License expiry info
$showLicExp = $true

# Save output to a file - $true or $false
$saveFile = $true
# File output path and filename
$outFile = "C:\admin\scripts\test\MyVeeamReport_$(Get-Date -format MMddyyyy_hhmmss).htm"
# Launch file after creation - $true or $false
$launchFile = $true

# Email configuration
$sendEmail = $false
$emailHost = "smtp.yourserver.com"
$emailUser = ""
$emailPass = ""
$emailFrom = "MyVeeamReport@yourdomain.com"
$emailTo = "you@youremail.com"
# Send report as attachment - $true or $false
$emailAttach = $false
# Email Subject
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$fullSubject = $true

# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 10
$repoWarn = 20
# Replica Target Free Space Remaining %
$replicaCritical = 10
$replicaWarn = 20
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
#endregion

#region VersionInfo
$vPCARversion = "1.4.1"
#
# Version 1.4.1 - SM
# Fixed issue with summary counts
# Version 1.4 - SM
# Misc minor tweaks/cleanup
# Added variable for report width
# Added variable for email subject
# Added ability to show/hide all report sections
# Added Protected/Unprotected VM Count to Summary
# Added per object details for sessions w/no details
# Added proxy host name to Proxy Details
# Added repository host name to Repository Details
# Added section showing successful sessions
# Added ability to view only last session per job
# Added Cluster field for protected/unprotected VMs
# Added catch for cifs repositories greater than 4TB as erroneous data is returned
# Added % Complete for Running Jobs
# Added ability to exclude multiple (vCenter) folders from Missing and Successful Backups section
# Added ability to exclude multiple (vCenter) datacenters from Missing and Successful Backups section
# Tweaked license info for better reporting across different date formats
#
# Version 1.3 - SM
# Now supports VBR v8
# For VBR v7, use report version 1.2
# Added more flexible options to save and launch file
#
# Version 1.2 - SM
# Added option to show VMs Successfully backed up
#
# Version 1.1.4 - SM
# Misc tweaks/bug fixes
# Reconfigured HTML a bit to help with certain email clients
# Added cell coloring to highlight status
# Added $rptTitle variable to hold report title
# Added ability to send report via email as attachment
#
# Version 1.1.3 - SM
# Added Details to Sessions with Warnings or Failures
#
# Version 1.1.2 - SM
# Minor tweaks/updates
# Added Veeam version info to header
#
# Version 1.1.1 - Shawn Masterson
# Based on vPowerCLI v6 Army Report (v1.1) by Thomas McConnell
# http://www.vpowercli.co.uk/2012/01/23/vpowercli-v6-army-report/
# http://pastebin.com/6p3LrWt7
#
# Tweaked HTML header (color, title)
#
# Changed report width to 1024px
#
# Moved hard-coded path to exe/dll files to user declared variables ($veeamExePath/$veeamDllPath)
#
# Adjusted sorting on all objects
#
# Modified info group/counts
#   Modified - Total Jobs = Job Runs
#   Added - Read (GB)
#   Added - Transferred (GB)
#   Modified - Warning = Warnings
#   Modified - Failed = Failures
#   Added - Failed (last session)
#   Added - Running (currently running sessions)
#
# Modified job lines
#   Renamed Header - Sessions with Warnings or Failures
#   Fixed Write (GB) - Broke with v7
#
# Added support license renewal
#   Credit - Gavin Townsend  http://www.theagreeablecow.com/2012/09/sysadmin-modular-reporting-samreports.html
#   Original  Credit - Arne Fokkema  http://ict-freak.nl/2011/12/29/powershell-veeam-br-get-total-days-before-the-license-expires/
#
# Modified Proxy section
#   Removed Read/Write/Util - Broke in v7 - Workaround unknown
#
# Modified Services section
#   Added - $runningSvc variable to toggle displaying services that are running
#   Added - Ability to hide section if no results returned (all services are running)
#   Added - Scans proxies and repositories as well as the VBR server for services
#
# Added VMs Not Backed Up section
#   Credit - Tom Sightler - http://sightunseen.org/blog/?p=1
#   http://www.sightunseen.org/files/vm_backup_status_dev.ps1
#
# Modified $reportMode
#   Added ability to run with any number of hours (8,12,72 etc)
#	Added bits to allow for zero sessions (semi-gracefully)
#
# Added Running Jobs section
#   Added ability to toggle displaying running jobs
#
# Added catch to ensure running v7 or greater
#
#
# Version 1.1
# Added job lines as per a request on the website
#
# Version 1.0
# Clean up for release
#
# Version 0.9
# More cmdlet rewrite to improve perfomace, credit to @SethBartlett
# for practically writing the Get-vPCRepoInfo
#
# Version 0.8
# Added Read/Write stats for proxies at requests of @bsousapt
# Performance improvement of proxy tear down due to rewrite of cmdlet
# Replaced 2 other functions
# Added Warning counter, .00 to all storage returns and fetch credentials for
# remote WinLocal repos
#
# Version 0.7
# Added Utilisation(Get-vPCDailyProxyUsage) and Modes 24, 48, Weekly, and Monthly
# Minor performance tweaks

#endregion

#region NonUser-Variables
# Get the B&R Server
$vbrServer = Get-VBRLocalHost
# Get all the VI proxies in your army
$viProxyList = Get-VBRViProxy
# Get all the backup repositories
$repoList = Get-VBRBackupRepository
# Get all sessions
$allSesh = Get-VBRBackupSession
# Get all the backup sessions for mode (timeframe)
if ($reportMode -eq "Monthly") {
        $HourstoCheck = 720
} elseif ($reportMode -eq "Weekly") {
        $HourstoCheck = 168
} else {
        $HourstoCheck = $reportMode
}
$seshList = $allSesh | Where-Object {($_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck)) -and ($_.State -ne "Working")}

#Get replica jobs
$repList = Get-VBRJob | Where-Object {$_.IsReplica}

# Get session information
$totalxfer = 0
$totalRead = 0
$seshList | ForEach-Object {$totalxfer += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$seshList | ForEach-Object {$totalRead += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
If ($onlyLast) {
	$tempSeshList = $seshList
	$seshList = @()
	foreach($job in (Get-VBRJob | ? {$_.JobType -eq "Backup"}))
	{
	$seshList += $TempSeshList | Where-Object {$_.Jobname -eq $job.name} | Sort-Object CreationTime -Descending | Select-Object -First 1
	}
}
$succesSessions = @($seshList | Where-Object {$_.Result -eq "Success"})
$warningSessions = @($seshList | Where-Object {$_.Result -eq "Warning"})
$failsSessions = @($seshList | Where-Object {$_.Result -eq "Failed"})
$totalSessions = @($seshList | Where-Object {$_.Result -eq "Failed" -Or $_.Result -eq "Success" -Or $_.Result -eq "Warning"})
$runningSessions = @($allSesh | Where-Object {$_.State -eq "Working"})
$failedSessions = @($seshList | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Append Report Mode to Report Title
If ($fullTitle) {
	If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
	        $rptTitle = "$rptTitle (Last $reportMode Hrs)"
	} else {
	        $rptTitle = "$rptTitle ($reportMode)"
	}
}

# Append Report Mode to Email subject
If ($fullSubject) {
	If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
	        $emailSubject = "$emailSubject (Last $reportMode Hrs)"
	} else {
	        $emailSubject = "$emailSubject ($reportMode)"
	}
}
#endregion

#region Functions

function Get-vPCProxyInfo {
	$vPCObjAry = @()
    function Build-vPCObj {param ([PsObject]$inputObj)
            $ping = new-object system.net.networkinformation.ping
            $pinginfo = $ping.send("$($inputObj.Host.RealName)")

            if ($pinginfo.Status -eq "Success") {
                    $hostAlive = "Alive"
            } else {
                    $hostAlive = "Dead"
            }

            $vPCFuncObject = New-Object PSObject -Property @{
                    ProxyName = $inputObj.Name
                    RealName = $inputObj.Host.RealName.ToLower()
                    Disabled = $inputObj.IsDisabled
                    Status  = $hostAlive
                    IP = $pinginfo.Address
                    Responce = $pinginfo.RoundtripTime
            }
            return $vPCFuncObject
    }
    Get-VBRViProxy | ForEach-Object {$vPCObjAry = $vPCObjAry + $(Build-vPCObj $_)}
	$vPCObjAry
}

function Get-vPCRepoInfo {
[CmdletBinding()]
        param (
                [Parameter(Position=0, ValueFromPipeline=$true)]
                [PSObject[]]$Repository
                )
        Begin {
                $outputAry = @()
                [Reflection.Assembly]::LoadFile($veeamDllPath) | Out-Null
                function Build-Object {param($name, $repohost, $path, $free, $total)
                        $repoObj = New-Object -TypeName PSObject -Property @{
                                        Target = $name
										RepoHost = $repohost.ToLower()
                                        Storepath = $path
                                        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
                                        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
                                        FreePercentage = [Math]::Round(($free/$total)*100)
                                }
                        return $repoObj | Select-Object Target, RepoHost, Storepath, StorageFree, StorageTotal, FreePercentage
                }
        }
        Process {
                foreach ($r in $Repository) {
                        if ($r.GetType().Name -eq [String]) {
                                $r = Get-VBRBackupRepository -Name $r
                        }
                        if ($r.Type -eq "WinLocal") {
                                $Server = $r.GetHost()
                                $FileCommander = [Veeam.Backup.Core.CWinFileCommander]::Create($Server.Info)
                                $storage = $FileCommander.GetDrives([ref]$null) | Where-Object {$_.Name -eq $r.Path.Substring(0,3)}
                                $outputObj = Build-Object $r.Name $server.RealName $r.Path $storage.FreeSpace $storage.TotalSpace
                        }
                        elseif ($r.Type -eq "LinuxLocal") {
                                $Server = $r.GetHost()
                                $FileCommander = new-object Veeam.Backup.Core.CSshFileCommander $server.info
                                $storage = $FileCommander.FindDirInfo($r.Path)
                                $outputObj = Build-Object $r.Name $server.RealName $r.Path $storage.FreeSpace $storage.TotalSize
                        }
                        elseif ($r.Type -eq "CifsShare") {
                                $Server = $r.GetHost()
								$fso = New-Object -Com Scripting.FileSystemObject
                                $storage = $fso.GetDrive($r.Path)
								# Catch shares with > 4TB space (not calculated correctly)
								If (!($storage.TotalSize) -or (($storage.TotalSize -eq 4398046510080) -and ($storage.AvailableSpace -eq 4398046510080))){
								    $outputObj = New-Object -TypeName PSObject -Property @{
                                        Target = $r.Name
										RepoHost = $server.RealName.ToLower()
                                        Storepath = $r.Path
                                        StorageFree = "Unknown"
                                        StorageTotal = "Unknown"
                                        FreePercentage = "Unknown"
                                	}
								} Else {
                                	$outputObj = Build-Object $r.Name $server.RealName $r.Path $storage.AvailableSpace $storage.TotalSize
								}
                        }
                        $outputAry = $outputAry + $outputObj
                }
        }
        End {
                $outputAry
        }
}

function Get-vPCReplicaTarget {
[CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [PSObject[]]$InputObj
    )
    BEGIN {
		$outputAry = @()
        $dsAry = @()
        if (($Name -ne $null) -and ($InputObj -eq $null)) {
                $InputObj = Get-VBRJob -Name $Name
        }
    }
    PROCESS {
        foreach ($obj in $InputObj) {
                        if (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
                        $esxi = $obj.GetTargetHost()
                            $dtstr =  $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName
                            $objoutput = New-Object -TypeName PSObject -Property @{
                                    Target = $esxi.Name
                                    Datastore = $obj.ViReplicaTargetOptions.DatastoreName
                                    StorageFree = [Math]::Round([Decimal]$dtstr.FreeSpace/1GB,2)
                                    StorageTotal = [Math]::Round([Decimal]$dtstr.Capacity/1GB,2)
                                    FreePercentage = [Math]::Round(($dtstr.FreeSpace/$dtstr.Capacity)*100)
                            }
                            $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
                            $outputAry = $outputAry + $objoutput
                        }
                        else {
                                return
                        }
        }
    }
    END {
                $outputAry | Select-Object Target, Datastore, StorageFree, StorageTotal, FreePercentage
    }
}

function Get-VeeamVersion {
    $veeamExe = Get-Item $veeamExePath
	$VeeamVersion = $veeamExe.VersionInfo.ProductVersion
	Return $VeeamVersion
}

function Get-VeeamSupportDate {
	#Get version and license info

	$regBinary = (Get-Item 'HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication\license').GetValue('Lic1')
	$veeamLicInfo = [string]::Join($null, ($regBinary | % { [char][int]$_; }))

	if($script:VeeamVersion -like "5*"){
		$pattern = "EXPIRATION DATE\=\d{1,2}\/\d{1,2}\/\d{1,4}"
	}
	elseif($script:VeeamVersion -like "6*"){
		$pattern = "Expiration date\=\d{1,2}\/\d{1,2}\/\d{1,4}"
	}
	elseif($script:VeeamVersion -like "8*"){
		$pattern = "expiration date\=\d{1,2}\/\d{1,2}\/\d{1,4}"
	}

	# Convert Binary key
	if($script:VeeamVersion -like "5*" -OR $script:VeeamVersion -like "6*" -OR $script:VeeamVersion -like "8*"){
		$expirationDate = [regex]::matches($VeeamLicInfo, $pattern)[0].Value.Split("=")[1]
		$datearray = $expirationDate -split '/'
		$expirationDate = Get-Date -Day $datearray[0] -Month $datearray[1] -Year $datearray[2]
		$totalDaysLeft = ($expirationDate - (get-date)).Totaldays.toString().split(",")[0]
		$totalDaysLeft = [int]$totalDaysLeft
		$objoutput = New-Object -TypeName PSObject -Property @{
			ExpDate = $expirationDate.ToShortDateString()
            DaysRemain = $totalDaysLeft
        }
		$objoutput
	}
	else{
		$objoutput = New-Object -TypeName PSObject -Property @{
		    ExpDate = "Failed"
		    DaysRemain = "Failed"
		}
		$objoutput
	}
}

function Get-VeeamServers {
	$vservers=@{}
	$outputAry = @()
	$vservers.add($($script:vbrserver.realname),"VBRServer")
	foreach ($srv in $script:viProxyList) {
		If (!$vservers.ContainsKey($srv.Host.Realname)) {
		  $vservers.Add($srv.Host.Realname,"ProxyServer")
		}
	}
	foreach ($srv in $script:repoList) {
		If (!$vservers.ContainsKey($srv.gethost().Realname)) {
		  $vservers.Add($srv.gethost().Realname,"RepoServer")
		}
	}
	$vservers = $vservers.GetEnumerator() | Sort-Object Name
	foreach ($vserver in $vservers) {
		$outputAry += $vserver.Name
	}
	return $outputAry
}

function Get-VeeamServices {
    param (
	  [PSObject]$inputObj)

    $outputAry = @()
	foreach ($obj in $InputObj) {
		$output = Get-Service -computername $obj -Name "*Veeam*" -exclude "SQLAgent*" |
	    Select @{Name="Server Name"; Expression = {$obj.ToLower()}}, @{Name="Service Name"; Expression = {$_.DisplayName}}, Status
	    $outputAry = $outputAry + $output
    }
$outputAry
}

function Get-VMsBackupStatus {
    param (
	  [String]$vcenter)

	# Convert exclusion list to simple regular expression
	$excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
	$excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
	$excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"

	$outputary = @()
	$vcenterobj = Get-VBRServer -Name $vcenter
	$vmobjs = Find-VBRObject -Server $vcenterobj |
		Where-Object {$_.Type -eq "VirtualMachine" -and $_.VMFolderName -notmatch $excludefolder_regex} |
		Where-Object {$_.Name -notmatch $excludevms_regex} |
		Where-Object {$_.GetParent("Datacenter") -notmatch $excludedc_regex}


	$jobobjids = [Veeam.Backup.Core.CHierarchyObj]::GetObjectsOnHost($vcenterobj.id) | Where-Object {$_.Type -eq "Vm"}

	foreach ($vm in $vmobjs) {
		$jobobjid = ($jobobjids | Where-Object {$_.ObjectId -eq $vm.Id}).Id
		if (!$jobobjid) {
			$jobobjid = $vm.FindParent("Datacenter").Id +  + $vm.Id
		}
		$vm | Add-Member -MemberType NoteProperty "JobObjId" -Value $jobobjid
	}

	# Get a list of all VMs from vCenter and add to hash table, assume Unprotected
	$vms=@{}
	foreach ($vm in $vmobjs)  {
		if(!$vms.ContainsKey($vm.JobObjId)) {
			$vmdc = [string]$vm.GetParent("Datacenter")
			Try {$vmclus = [string]$vm.GetParent("ClusterComputeResource")} Catch {$vmclus = ""}
			$vms.Add($vm.JobObjId, @("!", $vmdc, $vmclus, $vm.Name))
		}
	}

	# Find all backup job sessions that have ended in the last x hours
	$vbrjobs = Get-VBRJob | Where-Object {$_.JobType -eq "Backup"}
	$vbrsessions = Get-VBRBackupSession | Where-Object {$_.JobType -eq "Backup" -and $_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck)}

	# Find all Successfuly backed up VMs in selected sessions (i.e. VMs not ending in failure) and update status to "Protected"
	if ($vbrsessions) {
		foreach ($session in $vbrsessions) {
			foreach ($vm in ($session.gettasksessions() | Where-Object {$_.Status -ne "Failed"} | ForEach-Object { $_ })) {
				if($vms.ContainsKey($vm.Info.ObjectId)) {
					$vms[$vm.Info.ObjectId][0]=$session.JobName
				}
			}
		}
	}
	$vms.GetEnumerator() | Sort-Object Value
}

function Get-VMsMissingBackup {
	param (
		$vms)

	$outputary = @()
	foreach ($vm in $vms) {
	  if ($vm.Value[0] -eq "!") {
	    $objoutput = New-Object -TypeName PSObject -Property @{
		Datacenter = $vm.Value[1]
		Cluster = $vm.Value[2]
		Name = $vm.Value[3]
		}
		$outputAry += $objoutput
	  }
	}
	$outputAry | Select-Object Datacenter, Cluster, Name
}

function Get-VMsSuccessBackup {
	param (
		$vms)

	$outputary = @()
	foreach ($vm in $vms) {
	  if ($vm.Value[0] -ne "!") {
	    $objoutput = New-Object -TypeName PSObject -Property @{
		Datacenter = $vm.Value[1]
		Cluster = $vm.Value[2]
		Name = $vm.Value[3]
		}
		$outputAry += $objoutput
	  }
	}
	$outputAry | Select-Object Datacenter, Cluster, Name
}

#endregion

#region Report
# Get Veeam Version
$VeeamVersion = Get-VeeamVersion

If ($VeeamVersion -lt 8) {
	Write-Host "Script requires VBR v8 or greater" -ColourScheme $global:colours.Error.Foreground
	exit
}

# HTML Stuff
$headerObj = @"
<html>
        <head>
                <title>$rptTitle</title>
                <style>
                        body {font-family: Tahoma; background-color:#fff;}
						table {font-family: Tahoma;width: $($rptWidth)px;font-size: 12px;border-collapse:collapse;}
                        <!-- table tr:nth-child(odd) td {background: #e2e2e2;} -->
						th {background-color: #cccc99;border: 1px solid #a7a9ac;border-bottom: none;}
                        td {background-color: #ffffff;border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;vertical-align: top;}
                </style>
        </head>
"@

$bodyTop = @"
        <body>
			<center>
                <table cellspacing="0" cellpadding="0">
                        <tr>
                                <td style="width: 80%;height: 45px;border: none;background-color: #003366;color: White;font-size: 24px;vertical-align: bottom;padding: 0px 0px 0px 15px;">$rptTitle</td>
                                <td style="width: 20%;height: 45px;border: none;background-color: #003366;color: White;font-size: 12px;vertical-align:text-top;text-align:right;padding: 2px 3px 2px 3px;">v$vPCARversion</td>
                        </tr>
						<tr>
                                <td style="width: 80%;height: 35px;border: none;background-color: #003366;color: White;font-size: 10px;vertical-align: bottom;padding: 0px 0px 2px 3px;">Report generated: $(Get-Date -format g)</td>
                                <td style="width: 20%;height: 35px;border: none;background-color: #003366;color: White;font-size: 10px;vertical-align:bottom;text-align:right;padding: 2px 3px 2px 3px;">Veeam v$VeeamVersion</td>
                        </tr>
                </table>
"@

$subHead01 = @"
                <table>
                        <tr>
                                <td style="height: 35px;background-color: #eeeeee;color: #003366;font-size: 16px;font-weight: bold;vertical-align: middle;padding: 5px 0 0 15px;border-top: none;border-bottom: none;">
"@

$subHead01err = @"
                <table>
                        <tr>
                                <td style="height: 35px;background-color: #FF0000;color: #003366;font-size: 16px;font-weight: bold;vertical-align: middle;padding: 5px 0 0 15px;border-top: none;border-bottom: none;">
"@

$subHead02 = @"
                                </td>
                        </tr>
                </table>
"@

$footerObj = @"
</center>
</body>
</html>
"@

#Get VM Backup Status
$vmstatus = @()
foreach ($vcenter in $vcenters) {
	$status = Get-VMsBackupStatus $vcenter
	$vmstatus += $status
}

# VMs Missing Backups
$missingVMs = @()
$missingVMs = Get-VMsMissingBackup $vmstatus

# VMs Successfuly Backed Up
$successVMs = @()
$successVMs = Get-VMsSuccessBackup $vmstatus

# Get Summary Info
$vbrMasterHash = @{
	"Coordinator" = "$((gc env:computername).ToLower())"
	"Failed" = ($failedSessions | Measure-Object).Count
	"Sessions" = ($totalSessions | Measure-Object).Count
	"Read" = $totalRead
	"Transferred" = $totalXfer
	"Successful" = ($succesSessions | Measure-Object).Count
	"Warning" = ($warningSessions | Measure-Object).Count
	"Fails" = ($failsSessions | Measure-Object).Count
	"Running" = ($runningSessions | Measure-Object).Count
	"SuccessVM" = ($successVMs | Measure-Object).Count
	"FailedVM" = ($missingVMs | Measure-Object).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash

If ($onlyLast) {
	$bodySummary =  $vbrMasterObj | Select-Object Coordinator, @{Name="Unprotected VMs"; Expression = {$_.FailedVM}},
	@{Name="Protected VMs"; Expression = {$_.SuccessVM}}, @{Name="Jobs Run"; Expression = {$_.Sessions}},
	@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
	@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
	@{Name="Warnings"; Expression = {$_.Warning}},
	@{Name="Failed"; Expression = {$_.Failed}} | ConvertTo-HTML -Fragment
} Else {
	$bodySummary =  $vbrMasterObj | Select-Object Coordinator, @{Name="Unprotected VMs"; Expression = {$_.FailedVM}},
	@{Name="Protected VMs"; Expression = {$_.SuccessVM}}, @{Name="Total Sessions"; Expression = {$_.Sessions}},
	@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
	@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
	@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
	@{Name="Failed"; Expression = {$_.Failed}} | ConvertTo-HTML -Fragment
}



# Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs) {
	If ($missingVMs -ne $null) {
		$missingVMs = $missingVMs | Sort-Object Datacenter, Cluster, Name | ConvertTo-HTML -Fragment
		$bodyMissing = $subHead01err + "VMs with No Successful Backups" + $subHead02 + $missingVMs
	}
}

# Get VMs Successfuly Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {
	If ($successVMs -ne $null) {
		$successVMs = $successVMs | Sort-Object Datacenter, Cluster, Name | ConvertTo-HTML -Fragment
		$bodySuccess = $subHead01 + "VMs with Successful Backups" + $subHead02 + $successVMs
	}
}

# Get Running Jobs
$bodyRunning = $null
if ($showRunning -eq $true) {
        if (($runningSessions | Measure-Object).count -gt 0) {
                $bodyRunning = $runningSessions | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
                @{Name="Start Time"; Expression = {$_.CreationTime}},
                @{Name="Duration (Mins)"; Expression = {[Math]::Round((New-TimeSpan $(Get-Date $_.Progress.StartTime) $(Get-Date)).TotalMinutes,2)}},
				@{Name="% Complete"; Expression = {$_.Progress.Percents}},
                @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
                @{Name="Write (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}} | ConvertTo-HTML -Fragment
                $bodyRunning = $subHead01 + "Running Jobs" + $subHead02 + $bodyRunning
        }
}

# Get Sessions with Failures or Warnings
$bodySessWF = $null
if ($showWarnFail -eq $true) {
        $sessWF = @($warningSessions + $failsSessions)
        if (($sessWF | Measure-Object).count -gt 0) {
				If ($onlyLast) {
					$headerWF = "Jobs with Warnings or Failures"
				} Else {
					$headerWF = "Sessions with Warnings or Failures"
				}
                $bodySessWF = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
				@{Name="Start Time"; Expression = {$_.CreationTime}},
                @{Name="Stop Time"; Expression = {$_.EndTime}},
				@{Name="Duration (Mins)"; Expression = {[Math]::Round($_.WorkDetails.WorkDuration.TotalMinutes,2)}},
				@{Name="Details"; Expression = {
				If ($_.GetDetails() -eq ""){($_.GetDetails()).Replace("<br />"," - ") + ($_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + $_.GetDetails()}})}
				Else {($_.GetDetails()).Replace("<br />"," - ")}}},
				Result  | ConvertTo-HTML -Fragment
                $bodySessWF = $subHead01 + $headerWF + $subHead02 + $bodySessWF
        }
}

# Get Successful Sessions
$bodySessSucc = $null
if ($showSuccess -eq $true) {
       if (($succesSessions | Measure-Object).count -gt 0) {
	   			If ($onlyLast) {
					$headerSucc = "Successful Jobs"
				} Else {
					$headerSucc = "Successful Sessions"
				}
                $bodySessSucc = $succesSessions | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
                @{Name="Start Time"; Expression = {$_.CreationTime}},
                @{Name="Stop Time"; Expression = {$_.EndTime}},
				@{Name="Duration (Mins)"; Expression = {[Math]::Round($_.WorkDetails.WorkDuration.TotalMinutes,2)}},
				Result  | ConvertTo-HTML -Fragment
                $bodySessSucc = $subHead01 + $headerSucc + $subHead02 + $bodySessSucc
        }
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
	if ($viProxyList -ne $null) {
	        $bodyProxy = Get-vPCProxyInfo | Select-Object @{Name="Proxy Name"; Expression = {$_.ProxyName}},
	        @{Name="Proxy Host"; Expression = {$_.RealName}}, Disabled, @{Name="IP Address"; Expression = {$_.IP}},
			@{Name="RT (ms)"; Expression = {$_.Responce}}, Status | Sort-Object "Proxy Host" |  ConvertTo-HTML -Fragment
	        $bodyProxy = $subHead01 + "Proxy Details" + $subHead02 + $bodyProxy
	}
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
	if ($repoList -ne $null) {
	        $bodyRepo = $repoList | Get-vPCRepoInfo | Select-Object @{Name="Repository Name"; Expression = {$_.Target}},
	        @{Name="Repository Host"; Expression = {$_.RepoHost}},
			@{Name="Path"; Expression = {$_.Storepath}}, @{Name="Free (GB)"; Expression = {$_.StorageFree}},
	        @{Name="Total (GB)"; Expression = {$_.StorageTotal}}, @{Name="Free (%)"; Expression = {$_.FreePercentage}},
			@{Name="Status"; Expression = {
				If ($_.FreePercentage -lt $repoCritical) {"Critical"}
				ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
				ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
				Else {"OK"}}} | `
			Sort "Repository Name" | ConvertTo-HTML -Fragment
	        $bodyRepo = $subHead01 + "Repository Details" + $subHead02 + $bodyRepo
	}
}

# Get Replica Target Info
$bodyReplica = $null
If ($showReplica) {
	if ($repList -ne $null) {
	        $bodyReplica = $repList | Get-vPCReplicaTarget | Select-Object @{Name="Replica Target"; Expression = {$_.Target}}, Datastore,
	        @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
	        @{Name="Free (%)"; Expression = {$_.FreePercentage}},
			@{Name="Status"; Expression = {If ($_.FreePercentage -lt $replicaCritical) {"Critical"} ElseIf ($_.FreePercentage -lt $replicaWarn) {"Warning"} Else {"OK"}}} | `
			Sort "Replica Target" | ConvertTo-HTML -Fragment
	        $bodyReplica = $subHead01 + "Replica Details" + $subHead02 + $bodyReplica
	}
}

# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
	$bodyServices = Get-VeeamServers
	$bodyServices = Get-VeeamServices $bodyServices
	If ($hideRunningSvc) {$bodyServices = $bodyServices | Where-Object {$_.Status -ne "Running"}}
	If ($bodyServices -ne $null) {
		$bodyServices = $bodyServices | Select-Object "Server Name", "Service Name", Status | Sort-Object "Server Name", "Service Name" | ConvertTo-HTML -Fragment
		$bodyServices = $subHead01 + "Veeam Services" + $subHead02 + $bodyServices
	}
}

# Get License Info
$bodyLicense = $null
If ($showLicExp) {
	$bodyLicense = Get-VeeamSupportDate | Select-Object @{Name="Expiry Date"; Expression = {$_.ExpDate}}, @{Name="Days Remaining"; Expression = {$_.DaysRemain}}, `
		@{Name="Status"; Expression = {If ($_.DaysRemain -lt $licenseCritical) {"Critical"} ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"} ElseIf ($_.DaysRemain -eq "Failed") {"Failed"} Else {"OK"}}} | `
		ConvertTo-HTML -Fragment
	$bodyLicense = $subHead01 + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}

# Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummary + $bodyMissing + $bodySuccess + $bodyRunning + $bodySessWF + $bodySessSucc + $bodyRepo + $bodyProxy + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Add color to output depending on results
#Green
$htmlOutput = $htmlOutput.Replace("<td>Running<","<td style=""background-color: Green;color: White;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<","<td style=""background-color: Green;color: White;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<","<td style=""background-color: Green;color: White;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<","<td style=""background-color: Green;color: White;"">Success<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td>Warning<","<td style=""background-color: Yellow;"">Warning<")
#Red
$htmlOutput = $htmlOutput.Replace("<td>Stopped<","<td style=""background-color: Red;color: White;"">Stopped<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<","<td style=""background-color: Red;color: White;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<","<td style=""background-color: Red;color: White;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<","<td style=""background-color: Red;color: White;"">Dead<")
#endregion

#region Output
if ($sendEmail) {
        $smtp = New-Object System.Net.Mail.SmtpClient $emailHost
        $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass);
		$msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
		$msg.Subject = $emailSubject
		If ($emailAttach) {
			$body = "Veeam Report Attached"
			$msg.Body = $body
			$tempfile = "$env:TEMP\$rptTitle.htm"
			$htmlOutput | Out-File $tempfile
			$attachment = new-object System.Net.Mail.Attachment $tempfile
      		$msg.Attachments.Add($attachment)

		} Else {
			$body = $htmlOutput
			$msg.Body = $body
			$msg.isBodyhtml = $true
		}
        $smtp.send($msg)
		If ($emailAttach) {
			$attachment.dispose()
			Remove-Item $tempfile
		}
}

If ($saveFile) {
	$htmlOutput | Out-File $outFile
	If ($launchFile) {
		Invoke-Item $outFile
	}
}
#endregion