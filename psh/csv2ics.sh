#!/bin/bash

profil=$1
OBM_SCRIPTS="/home/stlo_agglo/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/agendas/${profil}"
DOMAIN_TO_PROCESS="obm.domain"

rm ${PROFILE_DATA}/*.iconv

for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e 's///' -ne '
/^".*[",]$/ {
p
}
/^".*[^"]$/ {
h
}
/^[^"].*[^"]$/ {
H
}
/^[^"].*"$/ {
H
x
s/\n/\\n/g
p
}' > $i.iconv
done

python csv2ics.py --profile agglolo_1 --data ${PROFILE_DATA} --domain ${DOMAIN_TO_PROCESS} > ~/log.txt 2>&1

