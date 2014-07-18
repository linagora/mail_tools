#//Set fixed start and end times
$startTime = (Get-Date -Hour 00 -Minute 00 -Second 00 -Day 01 -Month 09 -Year 2008)
$endTime = (Get-Date -Hour 23 -Minute 59 -Second 59 -Day 01 -Month 09 -Year 2017)
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().name
$rooms = (gc ..\input\rooms.txt) #//This could be done with get-mailbox as well
$blankPRF = (gc ..\input\calexp.PRF) #//This file is used to create Outlook profiles on Windows
#//$olEXE = "C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
$olEXE = "C:\Program Files\Microsoft Office\OFFICE12\OUTLOOK.EXE"
$mail_tools_home = "C:\Documents and Settings\test\mail_tools\input\"

if ($args.length -gt 0)
{
	$mode_extract = $args[0]
	Write-Host "..Extraction mode" $mode_extract -ForegroundColor Yellow 
}

foreach ($a_line in $rooms)
	{
		$credentials = $a_line.Split('	')
		$room = $credentials[0]
		if ($mode_extract -eq 'exchange')
		{
			$password = $credentials[1]
		}
		Write-Host "..starting" $room "/" $password -ForegroundColor Yellow 
		#// Tester avec clone VMWare
#//		Add-MailboxPermission -Identity $room -User $currentUser -Accessrights Fullaccess -InheritanceType all -Confirm:$False > $null
#//		$homeServer = Get-Mailbox $room | select ServerName
		$newPRF = $mail_tools_home + $room + ".prf"
		$exportDirectory = "..\output\agendas\" + $room + "\"
		if (!(Test-Path $exportDirectory))
		{
			md -Path $exportDirectory
		}

		$exportFileRecurrences = $exportDirectory + $room + ".csv"
#//		$exportFileWithoutRecurrences = "..\output\" + $room + "_withoutrecurrences.csv"

#//Create a new PRF file for the current Room mailbox
		Write-Host "..creating PRF"
#//		$blankPRF | foreach { $_ -replace "%UserName%", "$room"} | foreach { $_ -replace "%homeserver%", "$($homeServer.ServerName)"} | foreach { $_ -replace "calexp", "$room"} | Set-Content $newPRF
		$blankPRF | foreach { $_ -replace "%UserName%", "$room"} | foreach { $_ -replace "calexp", "$room"} | Set-Content $newPRF

#//Start Outlook and import the PRF
		if ($mode_extract -eq 'exchange')
		{
			Write-Host "..importing PRF $newPRF using Outlook"
			& $olEXE /importprf $newPRF /profile $room
		}
		else
		{
			Write-Host "..using PRF $newPRF using Outlook"
			& $olEXE /profile $room
		}
		sleep (3)

		if ($mode_extract -eq 'exchange')
		{
			$wshell = New-Object -ComObject wscript.shell;
			$wshell.AppActivate('Mot de passe');
			sleep(1);
			$wshell.SendKeys('+{TAB}');
			$wshell.SendKeys($room);
			$wshell.SendKeys('{TAB}');
			$wshell.SendKeys($password);
			# $wshell.SendKeys('{TAB}');
			# $wshell.SendKeys('saintlo.fr');
			$wshell.SendKeys('{ENTER}');
			sleep(1);

			$wshell.SendKeys('+{TAB}');
			$wshell.SendKeys($room);
			$wshell.SendKeys('{TAB}');
			$wshell.SendKeys($password);
			$wshell.SendKeys('{ENTER}');
			sleep(1);
		}

#//Logon to current Outlook session and export calender.
		Write-Host "..attaching to Outlook session"
		$olApplication = New-Object -ComObject Outlook.Application
		$olSession = $olApplication.Session
		$olSession.Logon('$room')
		sleep(3);

		$calFolder = 9
		
		# Export with recurrences
		$calItems = $olSession.GetDefaultFolder($calFolder).Items
		$calItems.Sort("[Start]")
		# $calItems.IncludeRecurrences = $true
		$calItems.IncludeRecurrences = $false
		$dateRange = "[End] >= '{0}' AND [Start] <= '{1}'" -f $startTime.ToString("g"), $endTime.ToString("g")
		$calExport = $calItems.Restrict($dateRange)

		$l_nb_item = 0

		Write-Host "..exporting Calendar recurring patterns to CSV"
		foreach ($Contact in $calExport) {
			Write-Host "..export item "+$l_nb_item
			# $Contact.Move($tempFolder) | foreach-object {Write-Progress "Backup unique items to temp folder..." $_.FullName; $_.FullName} | Out-Null
			$Contact | Export-Csv $exportFileRecurrences".item."$l_nb_item -encoding "unicode"
			# $Contact.RequiredAttendees | Export-Csv $exportFileRecurrences".itemrequiredattendees."$l_nb_item -encoding "unicode"
			# $Contact.OptionalAttendees | Export-Csv $exportFileRecurrences".itemoptionalattendees."$l_nb_item -encoding "unicode"
			$Contact.Recipients | Export-Csv $exportFileRecurrences".itemrecipients."$l_nb_item -encoding "unicode"
			# $Contact.Parent | Export-Csv $exportFileRecurrences".parent."$l_nb_item -encoding "unicode"
			$olRecurrences = $Contact.GetRecurrencePattern()
			$l_nb_reccurences = 0
			foreach($rec in $olRecurrences)
			{
				$l_nb_exception = 0
				$rec | Export-Csv $exportFileRecurrences".recurrence."$l_nb_reccurences"."$l_nb_item -encoding "unicode"
				foreach($a_exception in $rec.Exceptions)
				{
					$a_exception | Export-Csv $exportFileRecurrences".exception."$l_nb_exception"."$l_nb_reccurences"."$l_nb_item -encoding "unicode"
					$l_nb_exception = $l_nb_exception + 1;
					$appointment_item = $a_exception.AppointmentItem
					if($appointment_item)
					{
						$appointment_item | Export-Csv $exportFileRecurrences".appointmentitem."$l_nb_exception"."$l_nb_reccurences"."$l_nb_item -encoding "unicode"
					}
				}
				$l_nb_reccurences = $l_nb_reccurences + 1;
			}
			$l_nb_item = $l_nb_item + 1
		}
		# $calExport | Export-Csv $exportFileRecurrences -encoding "unicode"

		# Export without recurrences
#		$calItemsW = $olSession.GetDefaultFolder($calFolder).Items
#		$calItemsW.Sort("[Start]")
#		$calItemsW.IncludeRecurrences = $false
#		$dateRange = "[End] >= '{0}' AND [Start] <= '{1}'" -f $startTime.ToString("g"), $endTime.ToString("g")
#		$calExportW = $calItemsW.Restrict($dateRange)
#
#		Write-Host "..exporting Calendar without including recurrences to CSV"
#		$calExportW | Export-Csv $exportFileWithoutRecurrences -encoding "unicode"

#//		SEL
#//		Write-Host "..removing permissions"
#//		Remove-MailboxPermission -Identity $room -User $currentUser -Accessrights Fullaccess -Confirm:$False
#//		Write-Host "..removing PRF"
#//		Remove-Item $newPRF

		Write-Host "..killing Outlook"
		get-process OUTLOOK | Stop-Process
		sleep (4)
	}

