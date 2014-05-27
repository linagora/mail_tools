# -*- coding: utf-8 -*-

import glob

class VCard:
	def process_contact(a_contact_line, l_line_number):
		logger.debug('process_contact')
		line_len=len(a_contact_line)
		a_vcf_card = 'BEGIN:VCARD'

		if (l_line_number > 2) and (line_len > 1):
			contact_arr = a_contact_line.split(',')
			# Tel [45]
			a_vcf_card = a_vcf_card + '\nEMAIL;TYPE=PREF,INTERNET:'+contact_arr[58].strip('"')
		a_vcf_card = a_vcf_card + '\nEND:VCARD'

# main
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(format=FORMAT)
logger = logging.getLogger('tcpserver')
logger.setLevel('DEBUG')

data_directory='../output'
vcf_directory='../output/vcf'
files_contacts_arr = glob.glob(data_directory + '/' + profile_to_process + '_contacts.csv')
f_contacts_output = open(vcf_directory + '/' + profile_to_process + '_contacts.vcf', 'w')

a_vcard = VCard()
for a_contact_file in files_contacts_arr:
	f_contacts = open(a_contact_file, 'r')
	line_number=0

	for a_contact_line in f_contacts:
		line_number=line_number+1
		a_vcard.process_contact(a_contact_line, line_number)
	
