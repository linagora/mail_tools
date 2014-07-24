# -*- coding: utf-8 -*-

import glob
import csv
import codecs
import logging, argparse
import logging.handlers, logging.config

class VCard:
	def process_contact(self, a_contact_line):
		logger.debug('process_contact')
		line_len=len(a_contact_line)
		a_vcf_card = 'BEGIN:VCARD'
		contact_arr = a_contact_line
		logger.debug('a_contact_line:'+'#'.join(a_contact_line))
		# Tel [45]
		a_vcf_card = a_vcf_card + '\nVERSION:3.0'
		a_vcf_card = a_vcf_card + '\nN:'+contact_arr[27]
		if contact_arr[69].find(', ') != -1:
			a_vcf_card = a_vcf_card + '\nFN:'+contact_arr[69].split(', ')[0]+';'+contact_arr[69].split(', ')[1]
		a_vcf_card = a_vcf_card + '\nROLE:'+contact_arr[90]
		a_vcf_card = a_vcf_card + '\nX-OBM-COMPANY:'+contact_arr[53]
		a_vcf_card = a_vcf_card + '\nTITLE:'+contact_arr[134]
		for a_work_tel in contact_arr[45].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=WORK,VOICE:'+a_work_tel
		for a_home_tel in contact_arr[86].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=HOME,VOICE:'+a_home_tel
		for a_cell_tel in contact_arr[109].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=CELL,VOICE:'+a_cell_tel
		for a_work_fax in contact_arr[43].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=WORK,FAX:'+a_work_fax
		a_vcf_card = a_vcf_card + '\nX-OBM-COMMENT:'+contact_arr[7]
		a_vcf_card = a_vcf_card + '\nADR;TYPE=WORK:'+contact_arr[39]+';;'+contact_arr[42]+';'+contact_arr[37]+';;'+contact_arr[39]+';'+contact_arr[38]
		a_vcf_card = a_vcf_card + '\nLABEL;TYPE=WORK:'+contact_arr[39]+';;'+contact_arr[42]+';'+contact_arr[37]+';;'+contact_arr[39]+';'+contact_arr[80]
		a_vcf_card = a_vcf_card + '\nADR;TYPE=HOME:'+contact_arr[81]+';;'+contact_arr[84]+';'+contact_arr[79]+';;'+contact_arr[81]+';'+contact_arr[80]
		a_vcf_card = a_vcf_card + '\nLABEL;TYPE=HOME:'+contact_arr[81]+';;'+contact_arr[84]+';'+contact_arr[79]+';;'+contact_arr[81]+';'+contact_arr[38]
		a_address = contact_arr[59]
		a_vcf_card = a_vcf_card + '\nEMAIL;TYPE=PREF,INTERNET:'+a_address[a_address.rfind('(')+1:a_address.rfind(')')]
		a_vcf_card = a_vcf_card + '\nEND:VCARD\n'
		return a_vcf_card

# main
if __name__ == '__main__':
	FORMAT = '%(asctime)-15s %(message)s'
	logging.basicConfig(format=FORMAT)
	# logging.basicConfig(filename='/home/linagora/mail_tools/log/csv2vcf.log', level=logging.DEBUG)
	logger = logging.getLogger('csv2vcf')
	logger.setLevel(logging.DEBUG)

	a_handler = logging.handlers.TimedRotatingFileHandler(filename='/home/linagora/mail_tools/log/csv2vcf.log', when='D')
	logger.addHandler(a_handler)

	parser = argparse.ArgumentParser(description='CSV to VCF')
	parser.add_argument('--profile')
	parser.add_argument('--data')
	parser.add_argument('--output')
	parser.add_argument('--import_dir')
	args = parser.parse_args()

	logger.info('Début des traitements ...')

	profile_to_process=args.profile
	data_directory=args.data
	output_directory=args.output
	import_directory=args.import_dir
	files_contacts_arr = glob.glob(data_directory + '/' + 'contacts.csv.iconv')
	f_contacts_output = open(import_directory + '/' + profile_to_process + '_contacts.vcf', 'w')

	logger.debug(data_directory + '/' + profile_to_process + '_contacts.csv.iconv')
	a_vcard = VCard()
	for a_contact_file in files_contacts_arr:
		logger.debug('a_contact_file:'+a_contact_file)
		f_contacts = csv.reader(open(a_contact_file, 'rb'), doublequote=True)

		for a_contact_line in f_contacts:
			logger.debug('a_contact_line:'+a_contact_line[0])
			a_vcard_line = a_vcard.process_contact(a_contact_line).replace('\r\n', ' ')
			f_contacts_output.write(a_vcard_line)
	logger.info('Fin des traitements ...')
