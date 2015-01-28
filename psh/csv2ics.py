# -*- coding: utf-8 -*-

from icalendar import Calendar, Event, vCalAddress, vText
import pytz
import csv
from datetime import datetime
import logging
import logging.handlers, logging.config
import argparse
import tempfile, os

class NameEmpty(Exception):
	def __init__(self, value):
		self.value = value

	def __str__(self):
		return repr(self.value)

class EmailEmpty(Exception):
	def __init__(self, value):
		self.value = value

	def __str__(self):
		return repr(self.value)

class ExportToCsv:
	def load_virtual_alias(self, a_virtual_alias_dict):
		logger.debug('load_virtual_alias')
		f_virtual_alias = open('/etc/postfix/virtual_alias', 'r')
		for alias_line in f_virtual_alias:
			logger.debug('alias_line:'+alias_line)
			alias_line_arr = alias_line.rstrip('\n').split('\t')
			logger.debug('alias_line_arr[0]:'+alias_line_arr[0])
			logger.debug('alias_line_arr[1]:'+alias_line_arr[1])
			a_virtual_alias_dict[alias_line_arr[1]] = alias_line_arr[0]
		f_virtual_alias.close()
		
	def split_outlook_date(self, a_outlook_date):
		logger.debug('split_outlook_date:'+a_outlook_date)
		date_without_quote = a_outlook_date
		day_part=date_without_quote[:date_without_quote.find(' ')]
		logger.debug('daypart:%s' % day_part)
		hour_part=date_without_quote[date_without_quote.find(' ')+1:]
		logger.debug('hourpart:%s' % hour_part)

		a_day_arr=day_part.split('/')
		a_hour_arr=hour_part.split(':')

		return int(a_day_arr[0]), int(a_day_arr[1]), int(a_day_arr[2]), int(a_hour_arr[0]), int(a_hour_arr[1]), int(a_hour_arr[2]) 

	def deal_event(self, a_event_arr, a_tzinfo, a_recipient_dict, a_virtual_alias_dict, a_email_organizer):
		logger.debug('deal_event')
		event = Event()

		# summary - Body
		event.add('summary', a_event_arr[11])

		# dtstamp - Start
		a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(a_event_arr[56])
		message = 'day:%s month:%s year:%s hour:%s minute:%s second:%s' % (a_day, a_month, a_year, a_hour, a_minute, a_second)
		logger.debug(message)
		event.add('dtstart', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

		a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(a_event_arr[33])
		event.add('dtend', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

		# event.add('dtstamp', datetime(2005, 4, 4, 0, 10, 0, tzinfo=pytz.utc))

		organizer_field = a_event_arr[43]
		organizer=self.get_cal_address(organizer_field, a_recipient_dict, a_virtual_alias_dict, a_email_organizer)
		organizer.params['role'] = vText('CHAIR')

		event['organizer'] = organizer

		event_location = a_event_arr[36].replace('\r\n', '\n')
		# print event_location
		event['location'] = vText(event_location)

		#event['uid'] = '20050115T101010/27346262376@mxm.dk'
		event['uid'] = a_event_arr[10]

		event_importance = a_event_arr[16].replace('\r\n', '\n')
		# print event_importance
		event.add('priority', event_importance)
		event.add('description', a_event_arr[7].replace('\r\n', '\n'))

		required_attendees = []
		if a_event_arr[52]  != '':
			required_attendees = a_event_arr[52].split('; ')
		logger.debug('required_attendees:' + '#'.join(required_attendees))

		for a_required in required_attendees:
			logger.debug('for a_required:' + a_required)
			a_cal_required_attendee = self.get_cal_address(a_required, a_recipient_dict, a_virtual_alias_dict)
			if a_cal_required_attendee is not None:
				a_cal_required_attendee.params['ROLE'] = vText('REQ-PARTICIPANT')
				# TO FINISH
				# a_cal_required_attendee.params['PARTSTAT'] = vText('ACCEPTED')
				a_cal_required_attendee.params['CUTYPE'] = vText('INDIVIDUAL')
				# TO FINISH
				event.add('attendee', a_cal_required_attendee, encode=0)

		optional_attendees = []
		if a_event_arr[42]  != '':
			optional_attendees = a_event_arr[42].split('; ')
		logger.debug('optional_attendees:' + '#'.join(optional_attendees))

		for a_optional in optional_attendees:
			logger.debug('for a_optional:' + a_optional)
			a_cal_optional_attendee = self.get_cal_address(a_optional, a_recipient_dict, a_virtual_alias_dict)
			a_cal_optional_attendee.params['ROLE'] = vText('OPT-PARTICIPANT')
			# TO FINISH
			# a_cal_optional_attendee.params['PARTSTAT'] = vText('ACCEPTED')
			a_cal_optional_attendee.params['CUTYPE'] = vText('INDIVIDUAL')
			# TO FINISH
			event.add('attendee', a_cal_optional_attendee, encode=0)

		return event

	def process_dayofweek_mask(self, recurrence_dayofweek, week_day):
		list_of_days = []
		for i in sorted(week_day,cmp=lambda x,y: cmp(int(x), int(y)), reverse=True):
			logger.debug('i:'+i)
			logger.debug('recurrence_dayofweek:%d' % (recurrence_dayofweek))
			if recurrence_dayofweek >= int(i):
				list_of_days.append(week_day[i])
				recurrence_dayofweek = recurrence_dayofweek - int(i)
		return list_of_days

	def process_appointment(self, a_appointment_line, a_event, a_moved_event_list, a_tzinfo, a_recipient_dict, a_virtual_alias_dict, email_organizer, a_orig_time_exception):
		if len(a_appointment_line) > 1:
			a_appointment_arr = a_appointment_line
			a_moved_event = self.deal_event(a_appointment_arr, a_tzinfo, a_recipient_dict, a_virtual_alias_dict, email_organizer)
			# a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(a_event_arr[33])
			a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(a_orig_time_exception)
			a_moved_event.add('recurrence-id', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))
			a_moved_event_list.append(a_moved_event)

	def create_moved_event(self, l_event, a_recurrence_number, a_item_number, a_tzinfo, a_recipient_dict, a_virtual_alias_dict, a_orig_time_exception, email_organizer):
		a_moved_event_list = []
		
		files_appointement_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.appointmentitem.*.' + recurrence_number + '.' + item_number + '.iconv')
		for file_appointment in files_appointement_arr:
			f_appointment = csv.reader(open(file_appointment, 'rb'), doublequote=True)
			for appointment_line in f_appointment:
				self.process_appointment(appointment_line, l_event, a_moved_event_list, a_tzinfo, a_recipient_dict, a_virtual_alias_dict, email_organizer, a_orig_time_exception)
			# f_appointment.close()
		return a_moved_event_list[0]

	def process_exception(self, exception_line, l_event, l_tzinfo, a_date_list_exc, moved_events_list, a_recurrence_number, a_item_number, a_recipient_dict, a_virtual_alias_dict, email_organizer):
		logger.debug('process_exception')
		logger.debug(exception_line)
		line_len=len(exception_line)

		if (line_len > 1):
			exception_arr = exception_line
			is_deleted = exception_arr[5]
			a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(exception_arr[6])
			if  (is_deleted == "True"):
				logger.debug('is_deleted:'+','.join(exception_line))
				a_date_list_exc.append(datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(l_tzinfo)))
			else:
				logger.debug('is_not_deleted:'+','.join(exception_line))
				orig_time_exception = exception_arr[6]
				a_moved_event = self.create_moved_event(l_event, a_recurrence_number, a_item_number, l_tzinfo, a_recipient_dict, a_virtual_alias_dict, orig_time_exception,  email_organizer)
				moved_events_list.append(a_moved_event)

	def process_recurrence(self, recurrence_line, l_event, a_tzinfo):
		logger.debug('process_recurrence')
		line_len=len(recurrence_line)

		if (line_len > 1):
			recurrence_arr = recurrence_line
			recurrence_type = recurrence_arr[16]
			recurrence_noenddate = recurrence_arr[12]
			recurrence_interval = recurrence_arr[10]
			recurrence_instance = recurrence_arr[9]
			recurrence_dayofweek = recurrence_arr[5]
			recurrence_dayofmonth = recurrence_arr[4]
			recurrence_monthofyear = recurrence_arr[11]
			logger.debug('recurrence_type:'+recurrence_type)
			logger.debug('recurrence_interval:'+recurrence_interval)
			logger.debug('recurrence_dayofweek:'+recurrence_dayofweek)

			week_day = {'1': 'SU','2': 'MO','4': 'TU','8': 'WE','16': 'TH','32': 'FR','64': 'SA',}
			weekday_list = self.process_dayofweek_mask(int(recurrence_dayofweek), week_day)
			instance_dict = {'1' : '+1','2' : '2','3' : '3','4' : '4','5' : '-1',}
			# week_day = {'1': '0','2': '1','4': '2','8': '3','16': '4','5': '5','64': '6',}

			rec_dict = dict()
			if recurrence_type == '0':
				rec_dict = {'freq': 'daily',}
			elif recurrence_type == '1':
				rec_dict = {'freq': 'weekly', 'byday': weekday_list, }
				if  recurrence_interval != '1':
					logger.debug('week_day:'+','.join(weekday_list))
					rec_dict['interval'] = recurrence_arr[10]
			elif recurrence_type == '2':
				rec_dict = {'freq': 'monthly', 'bymonthday': recurrence_dayofmonth, }
				if  recurrence_interval != '1':
					logger.debug('week_day:'+','.join(weekday_list))
					rec_dict['interval'] = recurrence_arr[10]
			elif recurrence_type == '3':
				logger.debug('rrule added')
				rec_dict = {'freq': 'monthly', 'wkst': instance_dict[recurrence_instance] + weekday_list[0], }
				if  recurrence_interval != '1':
					logger.debug('week_day:'+','.join(weekday_list))
					rec_dict['interval'] =  recurrence_arr[10]
			elif recurrence_type == '4':
				logger.debug('rrule added')
			elif recurrence_type == '5':
				logger.debug('rrule added')
				rec_dict = {'freq': 'yearly', 'bymonth': recurrence_monthofyear, 'bymonthday': recurrence_dayofmonth, }
			elif recurrence_type == '6':
				logger.debug('rrule added')
				rec_dict = {'freq': 'yearly', 'wkst': instance_dict[recurrence_instance] + weekday_list[0], 'bymonth': recurrence_monthofyear}

			logger.debug('recurrence_noenddate:' + recurrence_noenddate)
			if recurrence_noenddate == 'False':
				a_day, a_month, a_year, a_hour, a_minute, a_second = self.split_outlook_date(recurrence_arr[14])
				rec_dict['until'] = datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo))

			l_event.add('rrule', rec_dict)

	def process_item(self, outlook_line, l_event, a_recipient_dict, a_virtual_alias_dict, a_email_organizer):
		logger.debug('process_item')
		line_len=len(outlook_line)

		if (line_len > 1):
			event_arr = outlook_line
			l_event = self.deal_event(event_arr, l_tzinfo, a_recipient_dict, a_virtual_alias_dict, a_email_organizer)
			is_recurring = event_arr[35]

	#		is_recurring = event_arr[35]
	#		event_conversation_index = event_arr[10]
	#		print is_recurring
	#		print event_conversation_index
	#		if is_recurring == '"True"':
	#			if event_conversation_index in recurring_events_dict:
	#				# evenement modifie
	#				recurrence_state = event_arr[45]
	#				if recurrence_state == '"3"':
	#					a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(event_arr[56])
	#					l_event.add('exdate', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(l_tzinfo)))
	#			else:
	#				l_event = deal_event(event_arr, l_tzinfo)
		return l_event, is_recurring

	def get_cal_address(self, a_address, a_recipient_dict, a_virtual_alias_dict, e_mail=None):
		logger.debug('get_cal_address')
		logger.debug('a_address:'+a_address)
		logger.debug('a_virtual_alias_dict keys:<'+'#'.join(a_virtual_alias_dict.keys())+'>')
		# l_attendee_email = a_address[a_address.rfind('('):a_address.rfind(')')]

		if (a_address == ''):
			l_attendee_cn = e_mail.split('@')[0]
		elif (a_address.rfind('(') == -1):
			l_attendee_cn = a_address
		elif a_address != '':
			# -1 is here not to get the space before (
			l_attendee_cn = a_address[:a_address.rfind('(')-1]

		logger.debug('l_attendee_cn:<'+l_attendee_cn+'>')
		logger.debug('recipient_dict keys:<'+'#'.join(a_recipient_dict.keys())+'>')
		a_part_stat = ''
		l_attendee_email = ''

		if l_attendee_cn == '':
			raise NameEmpty(a_address)
		if l_attendee_cn in a_recipient_dict:
			# l_attendee_email = a_recipient_dict[l_attendee_cn][0]
			# si le mail inclut le caractere @, c'est un mail qui sera garde
			# sinon c'est un DN. Le traitement recupere le CN en fin de chaine puis ajoute le domaine '@saint-lo.fr'
			l_attendee_email_to_process = a_recipient_dict[l_attendee_cn][0]
			logger.debug('l_attendee_email_to_process:<'+l_attendee_email_to_process+'>')

			# si le mail n'est pas dans la liste des recipients, renvoie None
			if l_attendee_email_to_process == '':
				return None

			if (l_attendee_email_to_process.find('@') == -1):
				# recuperation du CN en fin de chaine
				if (l_attendee_email_to_process.rfind('CN=') != -1):
					l_attendee_email_logindomain = l_attendee_email_to_process[l_attendee_email_to_process.rfind('CN=')+3:]+'@'+domain_to_process
					l_attendee_email_logindomain_low = l_attendee_email_logindomain.lower()
				else:
					l_attendee_email_logindomain_low = l_attendee_email_to_process.lower() + '@' + domain_to_process
			else:
				l_attendee_email_logindomain_low = l_attendee_email_to_process.lower()
		
			if l_attendee_email_logindomain_low in a_virtual_alias_dict:
				l_attendee_email = a_virtual_alias_dict[l_attendee_email_logindomain_low]
			else:
				l_attendee_email = l_attendee_email_logindomain_low

			if (a_recipient_dict[l_attendee_cn][1] == '0') or (a_recipient_dict[l_attendee_cn][1] == '3') or (a_recipient_dict[l_attendee_cn][1] == '1'):
				a_part_stat = 'ACCEPTED'
			elif (a_recipient_dict[l_attendee_cn][1] == '4'):
				a_part_stat = 'DECLINED'
			elif (a_recipient_dict[l_attendee_cn][1] == '2'):
				a_part_stat = 'TENTATIVE'
		else:
			l_attendee_email = e_mail
			a_part_stat = 'ACCEPTED'

		l_attendee = None
		if l_attendee_email is not None:
			l_attendee = vCalAddress('MAILTO:'+l_attendee_email.strip("'"))
			l_attendee.params['cn'] = vText(l_attendee_cn)
			l_attendee.params['PARTSTAT'] = vText(a_part_stat)

		return l_attendee

	def process_recipient(self, a_recipient_line, a_recipient_dict):
		logger.debug('process_recipient:'+','.join(a_recipient_line))

		if (len(a_recipient_line) > 1):
			recipient_arr = a_recipient_line
			for a_element in recipient_arr:
				logger.debug('a_element:'+a_element)

			a_address = recipient_arr[11].replace('\r\n', '\n')
			logger.debug('a_address:'+a_address)
			if (a_address.rfind('(') == -1):
				l_attendee_cn = a_address
			else:
				# -1 is here not to get the space before (
				l_attendee_cn = a_address[:a_address.rfind('(')-1]

			# for each name, stores mail, meetingresponsestatus
			recipient_mail = recipient_arr[4]
			a_recipient_dict[l_attendee_cn] = [ recipient_mail, recipient_arr[10],]

# main
if __name__ == '__main__':
	FORMAT = '%(asctime)-15s %(message)s'
	logging.basicConfig(format=FORMAT)
	# logging.basicConfig(filename='/home/linagora/mail_tools/log/csv2ics.log', level=logging.DEBUG)
	# loggin.config.fileConfig('/home/linagora/mail_tools/conf/logging.conf')
	logger = logging.getLogger('csv2ics')
	logger.setLevel(logging.DEBUG)

	a_handler = logging.handlers.TimedRotatingFileHandler(filename='/home/linagora/mail_tools/log/csv2ics.log', when='D')
	logger.addHandler(a_handler)

	logger.info('Début des traitements')
	parser = argparse.ArgumentParser(description='CSV to ICS')
	parser.add_argument('--profile')
	parser.add_argument('--data')
	parser.add_argument('--domain')
	parser.add_argument('--import_dir')
	args = parser.parse_args()

	cal = Calendar()
	cal.add('prodid', '-//My calendar product//mxm.dk//')
	cal.add('version', '2.0')

	profile_to_process=args.profile
	domain_to_process=args.domain
	data_directory=args.data
	import_directory=args.import_dir
	line_number = 0
	l_tzinfo="Europe/Paris"
	a_export_to_csv = ExportToCsv()
	virtual_alias_dict = dict()
	a_export_to_csv.load_virtual_alias(virtual_alias_dict)
	logger.debug('virtual_alias_dict keys:<'+'#'.join(virtual_alias_dict.keys())+'>')

	import glob
	from os.path import basename

	files_items_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.item.*.iconv')

	for a_item_file in files_items_arr:
		try:
			file_name = basename(a_item_file)
			item_number = file_name.split('.')[3]
			l_event = Event()
			is_recurring = "False"
			moved_events = []
			recipient_dict = dict()

			logger.debug('path to recipients:'+data_directory + '/' + profile_to_process + '.csv.itemrecipients.' + item_number + '.iconv')
			files_recipients_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.itemrecipients.' + item_number + '.iconv')
			for a_recipient_file in files_recipients_arr:
				logger.debug('recipients handled:' + a_recipient_file)
				f_recipient = csv.reader(open(a_recipient_file, 'rb'), doublequote=True)
				for recipient_line in f_recipient:
					a_export_to_csv.process_recipient(recipient_line, recipient_dict)
				# if len(recipient_dict) == 0:
					# raise EmailEmpty("No recipients found in "+a_recipient_file)

			logger.debug('recipient_dict keys1:<'+'#'.join(recipient_dict.keys())+'>')
			logger.debug('file handled:' + a_item_file)
			f_item = csv.reader(open(a_item_file, 'rb'), doublequote=True)
			for outlook_line in f_item:
				l_event, is_recurring = a_export_to_csv.process_item(outlook_line, l_event, recipient_dict, virtual_alias_dict, profile_to_process + '@' + domain_to_process)
			# f_item.close()

			if is_recurring == "True":
				files_recurrences_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.recurrence.*.' + item_number + '.iconv')
				for a_recurrence_file in files_recurrences_arr:
					# print a_recurrence_file
					file_rec_name = basename(a_recurrence_file)
					recurrence_number = file_rec_name.split('.')[3]

					f_recurrence = csv.reader(open(a_recurrence_file, 'rb'), doublequote=True)
					for recurrence_line in f_recurrence:
						a_export_to_csv.process_recurrence(recurrence_line, l_event, l_tzinfo)
					# f_recurrence.close()

					files_exceptions_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.exception.*.' + recurrence_number + '.' + item_number + '.iconv')
					a_date_list_exc = []
					for a_exception_file in files_exceptions_arr:
						logger.debug(a_exception_file)

						f_exception = csv.reader(open(a_exception_file, 'rb'), doublequote=True)
						for exception_line in f_exception:
							a_export_to_csv.process_exception(exception_line, l_event, l_tzinfo, a_date_list_exc, moved_events, recurrence_number, item_number, recipient_dict, virtual_alias_dict, profile_to_process + '@' + domain_to_process)
							logger.debug('a_date_list_exc len:%s' % len(a_date_list_exc))
						# f_exception.close()

					logger.debug('a_date_list_exc final len:%s' % len(a_date_list_exc))
					if len(a_date_list_exc) != 0:
						logger.debug('adding a exdate')
						l_event.add('exdate', a_date_list_exc)

			cal.add_component(l_event)

			for a_event in moved_events:
				cal.add_component(a_event)

		except EmailEmpty as e:
			logger.error('Exception caught')

		except NameEmpty as e:
			logger.error('Exception caught')

	f = open(os.path.join(import_directory, profile_to_process + '_agendas.ics'), 'wb')
	f.write(cal.to_ical())
	f.close()
	logger.info('Fin des traitements')

