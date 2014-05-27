#!/bin/bash

profil=$1

rm ../output/${profil}/*.iconv
cd ../output/${profil}

for i in `ls`
do
	iconv -f UTF-16 -t UTF-8 $i | sed -e 's///' -ne '
/^".*"$/ {
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

cd -
python csv2ics.py > ~/log.txt 2>&1

