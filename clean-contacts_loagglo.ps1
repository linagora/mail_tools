#########################################################
#														#
#	 Remove Duplicate Outlook Contacts by Jean Louw		#
#	 Blog http://powershellneedfulthings.blogspot.com/	#
#														#
#########################################################

$olSession = (New-Object -ComObject Outlook.Application).Session
$olSession.Logon('profil_agglolo_1') #Outlook is the profile name
$contactsFolder = 10
$tempFolderName = 'temp_folder_' + (get-date -Format ddmmyyyhhmmss)
$myContacts = $olSession.GetDefaultFolder($contactsFolder).Items
$tempFolder = $olSession.GetDefaultFolder($contactsFolder).Folders.Add($tempFolderName)

Write-Host "..getting unique items"
$uniqueContacts = $myContacts | Sort FullName -Unique

#move contacts to temp contacts folder
#foreach ($Contact in $uniqueContacts) {
#$Contact.Move($tempFolder) | foreach-object {Write-Progress "Backup unique items to temp folder..." $_.FullName; $_.FullName} | Out-Null
#}

#read default contacts again and dump to csv
Write-Host "..export duplicates to csv"
$duplicates = $olSession.GetDefaultFolder($contactsFolder).Items
$duplicates | Export-Csv duplicates.csv -encoding "Unicode"

#delete all contacts left in default contacts folder
#Foreach ($duplicate in $duplicates){
#$duplicate.Delete() | foreach-object {Write-Progress "Deleting duplicate..." $_.FullName; $_.FullName} | Out-Null
#}
