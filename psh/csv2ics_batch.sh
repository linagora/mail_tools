#!/bin/bash

MAIL_TOOLS_HOME="/home/linagora/mail_tools"

echo "Début des traitements ..."
cd ${MAIL_TOOLS_HOME}/psh

while read login_ad
do
	./csv2ics.sh ${login_ad}	
done < ${MAIL_TOOLS_HOME}"/input/users.txt"

echo "Fin des traitements"
