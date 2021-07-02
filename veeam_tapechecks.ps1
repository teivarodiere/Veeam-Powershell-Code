<# v2.1
 Function: The script is generates an email containing lists of tapes marked removal and recall
			One must configure the variables below which are passed to the 'veeam_dailychecks.ps1'

 Last modified: 2018-Dec-7
  Modifications
	2018-Nov-11:
		- Added 'action' variable to tell the daily checks which function to execute
		- Added 'runtime' variable to tell the daily checks script to run the 'action' functions on specific days only
			  Some checks such as tape removal or recall are only meant to run when specified by the variable 'runtime' from the veeam_tapechecks.ps1 script

 Tips:
	* Run this script in a powershell window opened as the 'Local Administrator'
 	* You could need to execute 'Set-ExecutionPolicy -executionPolicy RemoteSigned -force:$true -Confirm:$false -Scope Localmachine -ErrorAction SilentlyContinue'
 	* 'actions' can be a combinatioin of the following 'tapesAlerts,tapesToRemove,tapesToRecall,mdailyChecks,monthlyChecks
    	Note that tapesAlerts automatically implies both tapesToRemove and tapesToRecall and generates a single email for both tape removals and recalls.
 	* 'runtime' can be a combination of 'Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday' (only use single quotes '' at the beginning and end)
#>

# enter the details required for executing
$parameters = @{
	'veeamServerName' = "veeamvbr.internal.local"; #The script queries this Veeam server
	'action' = "tapesAlerts"; #Available actions are: tapesAlerts,tapesToRemove,tapesToRecall,mdailyChecks,monthlyChecks
	'reportCuttOffTime' = "7:00AM"; # Specifies the date where all functions use as a cut off date. For example: 7AM, all tapes expiring after 24hrs from 7AM will be reported for removal and offsiteing
	'showAllCSVFields' = $false; # if set to $true, the tape reacall email will contain more tape information for troubleshooting purposes
	'emailReport' = $true; # when set to true, the output report will be emailed.
	'smtpServer' = "smtp.internal.local"; # Mailserver
	'from' = "helpdesk@internal.local"; # The from address will be set to this email address
	'fromContactName' = "ACME"; # The originator Name will be sent to this.
	'toName' = "3rd Party Tape Vault Provider Name";  # The target Name will be sent to this.
	'replyTo' = "helpdesk@internal.local";  # The reply to will be set to this email address
	'toAddresses' = "tapeprovider@acmeprovider.com.au"; # The to address
	'siteAddress' = '<Full site address>'; # it will be included in the footer part of the email
	'contactNumber' = "00 0000 0000"; # it will be included in the footer part of the email
	'emailSubject' = "ACME - Tape Recall and Removal"; # it will prefix the subject line
	'schedule' = 'Monday,Tuesday,Wednesday,Thursday,Friday'; # specifies that the 'actions' will run on the specific days set by 'runtime'
	'footerText' = @"
To remove the media for this Library:
	1.) Select "Control" from the menu on the front panel of the Library.
	2.) Then select the "Magazine" option using the front panel of the Library.
	3.) Then select the "Left" (Slots 1-22) option to open the magazine on the corresponding side.
	4.) Remove each of the magazines on the corresponding side and remove the tapes as per the above list.
	5.) Replace any "blanks or Empty Slots" with a tape that are on the recall list above.
	6.) Replace the magazines into the library.
"@ # Added additional information in the email such as tape removal instructions
}
# Call the daily checks passing the customisations
$status = .\veeam_dailychecks.ps1 @parameters
