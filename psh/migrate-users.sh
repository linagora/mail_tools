#! /bin/bash

#set -ex

host_from='172.21.0.11'
host_to='127.0.0.1'
mdp_from='Mrd6u#qzoLgHe!'
mdp_to='secret'
#mdp_to='peThaef1ur7e'

date=$(date '+%d-%m-%Y à %H:%M:%S')
date_exec_deb=$(date '+%s')
> failed_user
> success_user

while read line
do
	echo "Traitement de la boite $line" >> users.log
	/usr/bin/imapsync --buffersize 8192000 --nosyncacls --subscribe --syncinternaldates --host1 $host_from --authuser1 cyrus --password1 $mdp_from --user1 $line@stif.info --authmech1 PLAIN --host2  $host_to --authuser2 cyrus --password2 $mdp_to --user2 $line --noauthmd5  --authmech2 PLAIN >> users.log
	if [ $? -eq 1 ];
	then
		echo "Le traitement de la boite $line a échoué"  >> users.log
		echo "" >> users.log
		echo $line >> failed_user
	else
		echo "Le traitement de la boite $line a réussi" >> users.log
		echo "" >> users.log
		echo $line >> success_user
	fi
		
done < liste_users

date_exec_fin=$(date '+%s')
date_exec=$(($date_exec_fin - $date_exec_deb))
h=$(($date_exec/3600))
m=$(($date_exec%3600/60))
s=$(($date_exec%60))
fails=`cat failed_user | wc -l`
success=`cat success_user | wc -l`
echo "La synchronisation des mails a duré $h heures, $m minutes et $s secondes"; >> users.log
echo "Il y a eu $success boites migrées" >> users.log
echo "Il y a eu $fails boites non-migrées" >> users.log
