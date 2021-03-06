﻿#########################################################
#														#
#	 Remove Duplicate Outlook Contacts by Jean Louw		#
#	 Blog http://powershellneedfulthings.blogspot.com/	#
#														#
#########################################################

$rooms = (gc ..\input\rooms.txt) #//This could be done with get-mailbox as well
$olEXE = "C:\Program Files\Microsoft Office\OFFICE12\OUTLOOK.EXE"

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
		& $olEXE /profile $room
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
		}

		$olSession = (New-Object -ComObject Outlook.Application).Session
		$olSession.Logon($room) #Outlook is the profile name
		$contactsFolder = 10
		$tempFolderName = 'temp_folder_' + (get-date -Format ddmmyyyhhmmss)
		$myContacts = $olSession.GetDefaultFolder($contactsFolder).Items
		$tempFolder = $olSession.GetDefaultFolder($contactsFolder).Folders.Add($tempFolderName)

#Write-Host "..getting unique items"
#$uniqueContacts = $myContacts | Sort FullName -Unique

#move contacts to temp contacts folder
#foreach ($Contact in $uniqueContacts) {
#$Contact.Move($tempFolder) | foreach-object {Write-Progress "Backup unique items to temp folder..." $_.FullName; $_.FullName} | Out-Null
#}

		#read default contacts again and dump to csv
		Write-Host "..export contacts to csv"
		$duplicates = $olSession.GetDefaultFolder($contactsFolder).Items
		$pathToExport = "..\output\contacts\" + $room + "\"
		$fileToExport = $pathToExport + "\contacts.csv"
		if (!(Test-Path $pathToExport))
		{
			md -Path $pathToExport
		}
		$duplicates | Export-Csv $fileToExport -encoding "Unicode"

		$item_num = 0

		foreach($rec in $duplicates)
		{
			$type_name = $rec.MessageClass
			Write-Host $type_name
			if ($type_name -eq "IPM.DistList")
			{
				Write-Host $rec.MemberCount
				$member_num = 1

				$dist_list_name = $fileToExport + "." + $item_num + "." + $rec.ConversationTopic.replace(' ', '_')
				while($member_num -le $rec.MemberCount)
				{
					$filename_to_use = $dist_list_name + "." + $member_num + ".list"
					$rec.GetMember($member_num) | Export-Csv $filename_to_use -encoding "Unicode" 
					$member_num = $member_num + 1
				}

			}
			elseif ($type_name -eq "IPM.Contact")
			{
				$filename_properties = $fileToExport +  "." + $item_num + ".properties"
				$rec.ItemProperties |  Export-Csv $filename_properties -encoding "Unicode" 
			}
			
			$item_num = $item_num + 1
		}


#delete all contacts left in default contacts folder
#Foreach ($duplicate in $duplicates){
#$duplicate.Delete() | foreach-object {Write-Progress "Deleting duplicate..." $_.FullName; $_.FullName} | Out-Null
#}
		Write-Host "..killing Outlook"
		get-process OUTLOOK | Stop-Process
		sleep (4)
	}

