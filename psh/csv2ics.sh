#!/bin/bash

profil=$1
OBM_SCRIPTS="/home/stlo_agglo/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/agendas/${profil}"
DOMAIN_TO_PROCESS="obm.domain"

rm ${PROFILE_DATA}/*.iconv

for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e '1,2d' -e 's///' -ne '
s/\([^,]\)""\([^,]\)/\1<double-quote\/>\2/g
s/\([^"]\),\([^"]\)/\1<comma\/>\2/g
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

sed -i -e 's/<double-quote\/>/"/g' -e 's/<comma\/>/,/g' ${PROFILE_DATA}/${profil}"_agendas.ics"
