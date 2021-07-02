<# Modifications
 This report does a few things
 - It reads in the morning report (VeeamInfrastructure.xml) produced by veeam_dailychecks.ps1 (via veeam_tapechecks.ps1)

 actions = tapesAlerts,tapesToRemove,tapesToRecall,mdailyChecks,monthlyChecks,tapesAfternoonChecks
 Note that tapesAlerts automatically implies both tapesToRemove and tapesToRecall and generates a single email for both tape removals and recalls.
 runtime = Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday
#>
$parameters = @{
			'veeamServerName' = "veamserver.internal.lan";
            # The action is set to scan for tapes that are expired or will expire between 5PM tonight and tomorrow morning
            # at the time set by the reportCuttOffTime value
			'action' = "afternoonTapeChecks";
            # Since we need to know if there are a mininmum of 'minimumTapesNeeded' tapes that have
            # 1) expired
            # 2) will expire between 5PM and 7AM(Set by tapeCuttOffTimeAhead)
            # Once the tapes are
            'tapeCuttOffTime' = "5:00PM" # Search for tapes expired after Today at this time(tapeCuttOffTime) but before tapeCuttOffTimeAhead
			'tapeCuttOffTimeAhead' = "7:00AM"; # Search for tapes expired after Today at this time(tapeCuttOffTime) but before tapeCuttOffTimeAhead
            'minimumTapesNeeded' = "Monday/3,Tuesday/3,Wednesday/3,Thursday/3,Friday5"; # Will alarm via email if not met Day/NumTapesNeeded
			'showAllCSVFields' = $false;
			'emailReport' = $true;
			'smtpServer' = "smtp.anz.hsi.local";
			'from' = "from_address@internal.lan";
			'fromContactName' = "My Company Service Desk";
			'toName' = "Iron Mountain";
			'replyTo' = "reply_address@internal.lan";
			'toAddresses' = "to_address@internal.lan";
			'siteAddress' = 'NextDC S1';
			'contactNumber' = "00 0000 0000";
            'emailSubject' = "My Company - Warning - Afternoon Tape Checks";
            'schedule' = "Monday,Tuesday,Wednesday,Thursday,Friday";
            'tapeRemovalInstructions' = @"
Instructions to remove the media from the Library:
   1.) Select "Control" from the menu on the front panel of the Library.
   2.) Then select the "Magazine" option using the front panel of the Library.
   3.) Then select the "Left" (Slots 1-11) or "Right" (Slots 12-22) option to open the magazine on the corresponding side.
   4.) Remove each of the magazines on the corresponding side and remove the tapes as per the above list.
   5.) Replace any "blanks or Empty Slots" with a tape that are on the recall list above.
   6.) Replace the magazines into the library.
"@
}
$status = .\veeam_dailychecks.ps1 @parameters
