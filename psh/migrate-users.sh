#! /bin/bash

#set -ex

<<<<<<< HEAD
host_from='10.0.1.14'
host_to='127.0.0.1'
WORK_FOLDER="/root/linagora/mail_tools"

date=$(date '+%d-%m-%Y à %H:%M:%S')
timestamp=$(date '+%Y%m%d%H%M%S')
date_exec_deb=$(date '+%s')
FAILED_USER_FILE=${WORK_FOLDER}"/output/failed_user"${timestamp}
SUCCESS_USER_FILE=${WORK_FOLDER}"/output/success_user"${timestamp}
LOG_FILE=${WORK_FOLDER}"/log/users_"${timestamp}".log"

> ${FAILED_USER_FILE}
> ${SUCCESS_USER_FILE}

while read user_stlo pwd_stlo
do
	LOGIN_USER=`echo ${user_stlo} | cut -d'@' -f1`

	echo "Traitement de la boite ${user_stlo}" >> ${LOG_FILE}

	/usr/bin/imapsync \
	--host1 $host_from --password1 ${pwd_stlo} --user1 ${LOGIN_USER} \
        --host2 $host_to --user2  ${user_stlo} \
	--subscribe_all \
	--exclude "Calendrier|Contacts|Journal|Notes|T\&AOI-ches|Bo\&AO4-te\ d\'envoi" \
	--exclude "Dossiers\ publics" \
	--regextrans2 "s/Brouillons/Drafts/" \
	--regextrans2 "s/\&AMk-l\&AOk-ments envoy\&AOk-s/Sent/" \
	--regextrans2 "s/\&AMk-l\&AOk-ments supprim\&AOk-s/Trash/" \
	--addheader \
	--password2 ${pwd_stlo} >> ${LOG_FILE}

	if [ $? -eq 1 ];
	then
		echo "Le traitement de la boite ${user_stlo} a échoué"  >> ${LOG_FILE}
		echo "" >> ${LOG_FILE}
		echo ${user_stlo} >> ${FAILED_USER_FILE}
	else
		echo "Le traitement de la boite ${user_stlo} a réussi" >> ${LOG_FILE}
		echo "" >> ${LOG_FILE}
		echo ${user_stlo} >> ${SUCCESS_USER_FILE}
	fi
		
done < ${WORK_FOLDER}/input/liste_users
=======
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
>>>>>>> cb3bede67fb105e0b58f8d6278472dd3f3cfea7b

date_exec_fin=$(date '+%s')
date_exec=$(($date_exec_fin - $date_exec_deb))
h=$(($date_exec/3600))
m=$(($date_exec%3600/60))
s=$(($date_exec%60))
<<<<<<< HEAD
fails=`cat ${FAILED_USER_FILE} | wc -l`
success=`cat ${SUCCESS_USER_FILE} | wc -l`
echo "La synchronisation des mails a duré $h heures, $m minutes et $s secondes"; >> ${LOG_FILE}
echo "Il y a eu $success boites migrées" >> ${LOG_FILE}
echo "Il y a eu $fails boites non-migrées" >> ${LOG_FILE}

=======
fails=`cat failed_user | wc -l`
success=`cat success_user | wc -l`
echo "La synchronisation des mails a duré $h heures, $m minutes et $s secondes"; >> users.log
echo "Il y a eu $success boites migrées" >> users.log
echo "Il y a eu $fails boites non-migrées" >> users.log
>>>>>>> cb3bede67fb105e0b58f8d6278472dd3f3cfea7b
