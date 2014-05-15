#//Set fixed start and end times
$startTime = (Get-Date -Hour 00 -Minute 00 -Second 00 -Day 01 -Month 09 -Year 2008)
$endTime = (Get-Date -Hour 23 -Minute 59 -Second 59 -Day 01 -Month 09 -Year 2015)
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().name
$rooms = (gc .\rooms.txt) #//This could be done with get-mailbox as well
$blankPRF = (gc .\calexp.PRF) #//This file is used to create Outlook profiles on Windows
#//$olEXE = "C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
$olEXE = "C:\Program Files\Microsoft Office\OFFICE11\OUTLOOK.EXE"

foreach ($room in $rooms)
	{
		Write-Host "..starting" $room -ForegroundColor Yellow 
		#// Tester avec clone VMWare
#//		Add-MailboxPermission -Identity $room -User $currentUser -Accessrights Fullaccess -InheritanceType all -Confirm:$False > $null
#//		$homeServer = Get-Mailbox $room | select ServerName
#//		$newPRF = ".\" + $room + ".prf"
		$exportFile = ".\" + $room + ".csv"

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
		$calItems = $olSession.GetDefaultFolder($calFolder).Items
		$calItems.Sort("[Start]")
		$calItems.IncludeRecurrences = $true
		$dateRange = "[End] >= '{0}' AND [Start] <= '{1}'" -f $startTime.ToString("g"), $endTime.ToString("g")
		$calExport = $calItems.Restrict($dateRange)

		Write-Host "..exporting Calendar to CSV"
		# //$calExport | select Subject, StartUTC, EndUTC, Duration, Organizer | sort StartUTC -Descending | Export-Csv $exportFile
		$calExport | Export-Csv $exportFile -encoding "unicode"

#//		SEL
#//		Write-Host "..removing permissions"
#//		Remove-MailboxPermission -Identity $room -User $currentUser -Accessrights Fullaccess -Confirm:$False
#//		Write-Host "..removing PRF"
#//		Remove-Item $newPRF

		Write-Host "..killing Outlook"
		get-process OUTLOOK | Stop-Process -Force
		sleep (15)
	}

