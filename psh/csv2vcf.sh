#!/bin/bash

profil=$1
OBM_SCRIPTS="/home/stlo_agglo/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/contacts/${profil}"
OUTPUT_DIRECTORY="${OBM_SCRIPTS}/output/contacts/${profil}"
DOMAIN_TO_PROCESS="obm.domain"

rm ${PROFILE_DATA}/*.iconv ${OUTPUT_DIRECTORY}/*.vcf

for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e 's///' -ne '
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

python csv2vcf.py --profile agglolo_1 --data ${PROFILE_DATA} --output ${OUTPUT_DIRECTORY} > ~/log.txt 2>&1

sed -i -e 's/<double-quote\/>/"/g' -e 's/<comma\/>/,/g' ${PROFILE_DATA}/${profil}"_contacts.vcf"
