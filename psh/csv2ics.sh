#!/bin/bash

# iconv -f UTF-16 -t UTF-8 profil_agglolo_1_withrecurrences.csv > profil_agglolo_1_withrecurrences.csv.iconv

cd ../output

for i in `ls`
do
	iconv -f UTF-16 -t UTF-8 $i > $i.iconv
done

python csv2ics.py

