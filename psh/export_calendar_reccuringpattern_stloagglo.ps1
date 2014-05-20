#//Set fixed start and end times
$startTime = (Get-Date -Hour 00 -Minute 00 -Second 00 -Day 01 -Month 09 -Year 2008)
$endTime = (Get-Date -Hour 23 -Minute 59 -Second 59 -Day 01 -Month 09 -Year 2017)
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().name
$rooms = (gc ..\input\rooms.txt) #//This could be done with get-mailbox as well
$blankPRF = (gc ..\input\calexp.PRF) #//This file is used to create Outlook profiles on Windows
#//$olEXE = "C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
$olEXE = "C:\Program Files\Microsoft Office\OFFICE11\OUTLOOK.EXE"

foreach ($room in $rooms)
	{
		Write-Host "..starting" $room -ForegroundColor Yellow 
		#// Tester avec clone VMWare
#//		Add-MailboxPermission -Identity $room -User $currentUser -Accessrights Fullaccess -InheritanceType all -Confirm:$False > $null
#//		$homeServer = Get-Mailbox $room | select ServerName
#//		$newPRF = ".\" + $room + ".prf"
		$exportFileRecurrences = "..\output\" + $room + ".csv"
		$exportFileWithoutRecurrences = "..\output\" + $room + "_withoutrecurrences.csv"

#//Create a new PRF file for the current Room mailbox
		Write-Host "..creating PRF"
#//		$blankPRF | foreach { $_ -replace "%UserName%", "$room"} | foreach { $_ -replace "%homeserver%", "$($homeServer.ServerName)"} | foreach { $_ -replace "calexp", "$room"} | Set-Content $newPRF

#//Start Outlook and import the PRF
		Write-Host "..importing PRF using Outlook"
		#// & $olEXE /importPRF $newPRF /profile $room
		& $olEXE /profile $room
		sleep (15)

#//Logon to current Outlook session and export calender.
		Write-Host "..attaching to Outlook session"
		$olApplication = New-Object -ComObject Outlook.Application
		$olSession = $olApplication.Session
		$olSession.Logon('$room')
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
			# $Contact.Move($tempFolder) | foreach-object {Write-Progress "Backup unique items to temp folder..." $_.FullName; $_.FullName} | Out-Null
			$Contact | Export-Csv $exportFileRecurrences".item."$l_nb_item -encoding "unicode"
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
		get-process OUTLOOK | Stop-Process -Force
		sleep (15)
	}

