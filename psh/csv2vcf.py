# -*- coding: utf-8 -*-

import glob
import logging, argparse

class VCard:
	def process_contact(self, a_contact_line):
		logger.debug('process_contact')
		line_len=len(a_contact_line)
		a_vcf_card = 'BEGIN:VCARD'
		contact_arr = a_contact_line.split('","')
		# Tel [45]
		a_vcf_card = a_vcf_card + '\nVERSION:3.0'
		a_vcf_card = a_vcf_card + '\nN:'+contact_arr[27].strip('"')
		a_vcf_card = a_vcf_card + '\nFN:'+contact_arr[27].strip('"')
		a_vcf_card = a_vcf_card + '\nORG:'+contact_arr[53].strip('"')
		a_vcf_card = a_vcf_card + '\nTITLE:'+contact_arr[134].strip('"')
		for a_work_tel in contact_arr[45].strip('"').split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=WORK,VOICE:'+a_work_tel
		for a_home_tel in contact_arr[86].strip('"').split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=HOME,VOICE:'+a_home_tel
		for a_work_address in contact_arr[36].strip('"').split(';'):
			a_vcf_card = a_vcf_card + '\nADR;TYPE=WORK:;;'+a_work_address
			a_vcf_card = a_vcf_card + '\nLABEL;TYPE=WORK:;;'+a_work_address
		for a_home_address in contact_arr[78].strip('"').split(';'):
			a_vcf_card = a_vcf_card + '\nADR;TYPE=HOME:;;'+a_home_address
			a_vcf_card = a_vcf_card + '\nLABEL;TYPE=HOME:;;'+a_home_address
		a_vcf_card = a_vcf_card + '\nEMAIL;TYPE=PREF,INTERNET:'+contact_arr[57].strip('"')
		a_vcf_card = a_vcf_card + '\nEND:VCARD\n'
		return a_vcf_card

# main
if __name__ == '__main__':
	FORMAT = '%(asctime)-15s %(message)s'
	logging.basicConfig(format=FORMAT)
	logger = logging.getLogger('tcpserver')
	logger.setLevel('DEBUG')

	parser = argparse.ArgumentParser(description='CSV to VCF')
	parser.add_argument('--profile')
	parser.add_argument('--data')
	parser.add_argument('--output')
	args = parser.parse_args()

	profile_to_process=args.profile
	data_directory=args.data
	output_directory=args.output
	files_contacts_arr = glob.glob(data_directory + '/' + profile_to_process + '_contacts.csv.iconv')
	f_contacts_output = open(output_directory + '/' + profile_to_process + '_contacts.vcf', 'w')

	a_vcard = VCard()
	for a_contact_file in files_contacts_arr:
		logger.debug('a_contact_file:'+a_contact_file)
		f_contacts = open(a_contact_file, 'r')

		for a_contact_line in f_contacts:
			a_vcard_line = a_vcard.process_contact(a_contact_line)
			f_contacts_output.write(a_vcard_line)
		
