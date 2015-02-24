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
		for (i, a_field) in enumerate(contact_arr):
			contact_arr[i] = a_field.replace('\r\n', ' ').replace('\n', ' ')
		logger.debug('a_contact_line:'+'#'.join(a_contact_line))
		# Tel [45]
		a_vcf_card = a_vcf_card + '\nVERSION:3.0'
		a_vcf_card = a_vcf_card + '\nN:'+contact_arr[27].replace('\n', ' ')
		if contact_arr[69].find(', ') != -1:
			a_vcf_card = a_vcf_card + '\nFN:'+contact_arr[69].split(', ')[0].replace('\n', ' ')+';'+contact_arr[69].split(', ')[1].replace('\n', ' ')
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

	def load_vcf_dict(self, a_vcf_dict, a_item_num):
		logger.debug('load_vcf_dict')
		files_properties_arr = glob.glob(data_directory + '/' + 'contacts.csv.' + str(a_item_num) + '.properties.iconv')
		for a_properties_file in files_properties_arr:
			f_properties = csv.reader(open(a_properties_file, 'rb'), doublequote=True)

			for a_property_line in f_properties:
				a_vcf_dict[a_property_line[5]] = a_property_line[9]

	def process_contact_without_version(self, a_contact_line, a_item_num):
		logger.debug('process_contact_without_version')
		line_len=len(a_contact_line)
		a_vcf_card = 'BEGIN:VCARD'
		contact_arr = a_contact_line
		for (i, a_field) in enumerate(contact_arr):
			contact_arr[i] = a_field.replace('\n', ' ').replace('\r\n', ' ')
		logger.debug('a_contact_line:'+'#'.join(a_contact_line))
		# Tel [45]
		a_vcf_card = a_vcf_card + '\nVERSION:3.0'
		a_vcf_card = a_vcf_card + '\nN:'+contact_arr[27].replace('\n', ' ')
		vcf_dict = dict()
		self.load_vcf_dict(vcf_dict, a_item_num)
		for a_key in vcf_dict.keys():
			logger.debug('Key:'+a_key+'@'+vcf_dict[a_key])

		if vcf_dict['FileAs'].find(', ') != -1:
			a_vcf_card = a_vcf_card + '\nFN:'+vcf_dict['FileAs'].split(', ')[0].replace('\n', ' ')+';'+vcf_dict['FileAs'].split(', ')[1].replace('\n', ' ')
		a_vcf_card = a_vcf_card + '\nROLE:'+vcf_dict['JobTitle']
		a_vcf_card = a_vcf_card + '\nX-OBM-COMPANY:'+vcf_dict['CompanyName']
		a_vcf_card = a_vcf_card + '\nTITLE:'+vcf_dict['Title']
		for a_work_tel in vcf_dict['BusinessTelephoneNumber'].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=WORK,VOICE:'+a_work_tel
		for a_home_tel in vcf_dict['HomeTelephoneNumber'].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=HOME,VOICE:'+a_home_tel
		for a_cell_tel in vcf_dict['MobileTelephoneNumber'].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=CELL,VOICE:'+a_cell_tel
		for a_work_fax in vcf_dict['BusinessFaxNumber'].split(';'):
			a_vcf_card = a_vcf_card + '\nTEL;TYPE=WORK,FAX:'+a_work_fax
		a_vcf_card = a_vcf_card + '\nX-OBM-COMMENT:'+vcf_dict['Body']+' '+vcf_dict['FullName']+' '+vcf_dict['Department']
		a_vcf_card = a_vcf_card + '\nADR;TYPE=WORK:'+vcf_dict['BusinessAddressPostalCode']+';;'+vcf_dict['BusinessAddressStreet']+';'+vcf_dict['BusinessAddressCity']+';;'+vcf_dict['BusinessAddressPostalCode']+';'+vcf_dict['BusinessAddressCountry']
		a_vcf_card = a_vcf_card + '\nLABEL;TYPE=WORK:'+vcf_dict['BusinessAddressPostalCode']+';;'+vcf_dict['BusinessAddressStreet']+';'+vcf_dict['BusinessAddressCity']+';;'+vcf_dict['BusinessAddressPostalCode']+';'+vcf_dict['HomeAddressCountry']
		a_vcf_card = a_vcf_card + '\nADR;TYPE=HOME:'+vcf_dict['HomeAddressPostalCode']+';;'+vcf_dict['HomeAddressStreet']+';'+vcf_dict['HomeAddressCity']+';;'+vcf_dict['HomeAddressPostalCode']+';'+vcf_dict['HomeAddressCountry']
		a_vcf_card = a_vcf_card + '\nLABEL;TYPE=HOME:'+vcf_dict['HomeAddressPostalCode']+';;'+vcf_dict['HomeAddressStreet']+';'+vcf_dict['HomeAddressCity']+';;'+vcf_dict['HomeAddressPostalCode']+';'+vcf_dict['BusinessAddressCountry']
		a_address = vcf_dict['Email1DisplayName']
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

		item_num = 0
		for a_contact_line in f_contacts:
			logger.debug('a_contact_line:'+a_contact_line[0])
			# the number of fields indicates which function to call
			# different users may have different number of records in the extraction
			if a_contact_line[19] == 'IPM.Contact' and len(a_contact_line) == 169:
				a_vcard_line = a_vcard.process_contact(a_contact_line).replace('\r\n', ' ').replace('\r', '\\n')
				f_contacts_output.write(a_vcard_line)
			elif a_contact_line[19] == 'IPM.Contact' and len(a_contact_line) == 54:
				a_vcard_line = a_vcard.process_contact_without_version(a_contact_line, item_num).replace('\r\n', ' ').replace('\r', '\\n')
				f_contacts_output.write(a_vcard_line)
			elif a_contact_line[19] == 'IPM.DistList':
				logger.debug('IPM.DistList')
				# a_vcard_line = a_vcard.process_distlist(a_contact_line, item_num).replace('\r\n', ' ')
				# f_contacts_output.write(a_vcard_line)
			item_num = item_num + 1
	logger.info('Fin des traitements ...')
