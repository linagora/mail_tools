#!/bin/bash

profil=$1
OBM_SCRIPTS="/home/linagora/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/contacts/${profil}"
OUTPUT_DIRECTORY="${OBM_SCRIPTS}/output/contacts/${profil}"
DOMAIN_TO_PROCESS="saint-lo.fr"

rm ${PROFILE_DATA}/*.iconv ${OUTPUT_DIRECTORY}/*.vcf

# la commande tr sert a supprimer les eventuels caracteres NUL generes par l'export Outlook et non geres par la librairie csv de python
for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16LE -t UTF-8 $i | tr -d '\000' | sed -e '1,2d' > $i.iconv
done

timestamp=$(date '+%Y%m%d%H%M%S')
python csv2vcf.py --profile ${profil} --data ${PROFILE_DATA} --output ${OUTPUT_DIRECTORY} > "${OBM_SCRIPTS}/log/csv2vcf_"${timestamp}".txt" 2>&1
