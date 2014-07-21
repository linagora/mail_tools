#!/bin/bash

profil=$1
OBM_SCRIPTS="/home/linagora/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/agendas/${profil}"
DATA_TO_IMPORT="${OBM_SCRIPTS}/import/agendas/${profil}"
DOMAIN_TO_PROCESS="saint-lo.fr"

rm ${PROFILE_DATA}/*.iconv 2> /dev/null

if  ! [[ -d "${DATA_TO_IMPORT}" ]];
then
	mkdir -p ${DATA_TO_IMPORT}
fi

for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e '1,2d' > $i.iconv
done

timestamp=$(date '+%Y%m%d%H%M%S')
python csv2ics.py --profile ${profil} --data ${PROFILE_DATA} --domain ${DOMAIN_TO_PROCESS} --import_dir ${DATA_TO_IMPORT} > "${OBM_SCRIPTS}/log/csv2ics_"${timestamp}".txt" 2>&1
