#!/bin/bash

profil=$1
OBM_SCRIPTS="/root/linagora/mail_tools"
PROFILE_DATA="${OBM_SCRIPTS}/output/agendas/${profil}"
DOMAIN_TO_PROCESS="saint-lo.fr"

rm ${PROFILE_DATA}/*.iconv

for i in `find ${PROFILE_DATA} -type f`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e '1,2d' \
-e '# Traitement des double-quotes situees en milieu de champs
s/\([^,]\)""\([^,]\)/\1<double-quote\/>\2/g
# Traitement des virgules situees en milieu de champs
#	en debut de ligne
s/^"\([^"]*\),\([^"]*\)"/\1<comma\/>\2/g
#	en fin de ligne
s/"\([^"]*\),\([^"]*\)"$/\1<comma\/>\2/g
#	en milieu de ligne
s/,"\([^"]*\),\([^"]*\)",/\1<comma\/>\2/g
# Traitement des retours-chariots a l interieur du champ
# 	un champ avec un retour-chariot au milieu commence par une double quote et se termine par un retour chariot
/"[^"]*$/ {
h
}
# 	cas d une ligne avec plusieurs retour chariot
# 	la ligne n inclut pas de double quote 
/^[^"]*$/ {
H
}
# 	ou le champ ne commence pas par une double quote
/^[^"]*",/ {
H
x
s/\n/\\n/g
p
}
# 	ou la ligne se termine par une double quote
/^[^"]*"$/ {
H
x
s/\n/\\n/g
p
}

##/^".*[",]$/ {
##p
##}
##/^".*[^"]$/ {
##h
##}
##/^[^"].*[^"]$/ {
##H
##}
##/^[^"].*"$/ {
##H
##x
##s/\n/\\n/g
##p
##}
' \
-e 's///' \
> $i.iconv
done

timestamp=$(date '+%Y%m%d%H%M%S')
python csv2ics.py --profile ${profil} --data ${PROFILE_DATA} --domain ${DOMAIN_TO_PROCESS} > "${OBM_SCRIPTS}/log/csv2ics_"${timestamp}".txt" 2>&1

sed -i -e 's/<double-quote\/>/"/g' -e 's/<comma\/>/,/g' ${PROFILE_DATA}/${profil}"_agendas.ics"
