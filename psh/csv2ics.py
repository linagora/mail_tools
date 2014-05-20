# -*- coding: utf-8 -*-

from icalendar import Calendar, Event, vCalAddress, vText
import pytz
from datetime import datetime
import logging

def split_outlook_date(a_outlook_date):
	date_without_quote = a_outlook_date.strip('"')
	day_part=date_without_quote[:date_without_quote.find(' ')]
	logger.debug('daypart:%s' % day_part)
	hour_part=date_without_quote[date_without_quote.find(' ')+1:]
	logger.debug('hourpart:%s' % hour_part)

	a_day_arr=day_part.split('/')
	a_hour_arr=hour_part.split(':')

	return int(a_day_arr[0]), int(a_day_arr[1]), int(a_day_arr[2]), int(a_hour_arr[0]), int(a_hour_arr[1]), int(a_hour_arr[2]) 

def deal_event(a_event_arr, a_tzinfo):
	# print event_arr
	# print "deal event"
	event = Event()

	# summary - Body
	event.add('summary', a_event_arr[11])

	# dtstamp - Start
	a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(a_event_arr[56])
	message = 'day:%s month:%s year:%s hour:%s minute:%s second:%s' % (a_day, a_month, a_year, a_hour, a_minute, a_second)
	logger.debug(message)
	event.add('dtstart', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

	a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(a_event_arr[33])
	event.add('dtend', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

	# event.add('dtstamp', datetime(2005, 4, 4, 0, 10, 0, tzinfo=pytz.utc))

	organizer_field = a_event_arr[43].strip('"')
	organizer=get_cal_address(organizer_field)
	organizer.params['role'] = vText('CHAIR')

	event['organizer'] = organizer

	event_location = a_event_arr[36][1:len(a_event_arr[36])-1]
	# print event_location
	event['location'] = vText(event_location)

	#event['uid'] = '20050115T101010/27346262376@mxm.dk'
	event['uid'] = a_event_arr[10][1:len(a_event_arr[10])-1]

	event_importance = a_event_arr[16][1:len(a_event_arr[16])-1]
	# print event_importance
	event.add('priority', event_importance)

	required_attendees = a_event_arr[52][1:len(a_event_arr[52])-1].split(';')
	# print a_event_arr[52]
	# print required_attendees
	for a_required in required_attendees:
		a_cal_required_attendee = get_cal_address(a_required)
		a_cal_required_attendee.params['ROLE'] = vText('REQ-PARTICIPANT')
		event.add('attendee', a_cal_required_attendee, encode=0)

	return event

def process_dayofweek_mask(recurrence_dayofweek, week_day):
	list_of_days = []
	for i in sorted(week_day,cmp=lambda x,y: cmp(int(x), int(y)), reverse=True):
		logger.debug('i:'+i)
		logger.debug('recurrence_dayofweek:%d' % (recurrence_dayofweek))
		if recurrence_dayofweek >= int(i):
			list_of_days.append(week_day[i])
			recurrence_dayofweek = recurrence_dayofweek - int(i)
	return list_of_days
			


def process_recurrence(recurrence_line, l_event, l_line_number):
	logger.debug('process_recurrence')
	line_len=len(recurrence_line)

	if (l_line_number > 2) and (line_len > 1):
		recurrence_arr = recurrence_line.split(',')
		recurrence_type = recurrence_arr[16].strip('"')
		recurrence_interval = recurrence_arr[10].strip('"')
		recurrence_instance = recurrence_arr[9].strip('"')
		recurrence_dayofweek = recurrence_arr[5].strip('"')
		recurrence_dayofmonth = recurrence_arr[4].strip('"')
		recurrence_monthofyear = recurrence_arr[11].strip('"')
		logger.debug('recurrence_type:'+recurrence_type)
		logger.debug('recurrence_interval:'+recurrence_interval)
		logger.debug('recurrence_dayofweek:'+recurrence_dayofweek)

		week_day = {'1': 'SU','2': 'MO','4': 'TU','8': 'WE','16': 'TH','32': 'FR','64': 'SA',}
		weekday_list = process_dayofweek_mask(int(recurrence_dayofweek), week_day)
		instance_dict = {'1' : '+1','2' : '2','3' : '3','4' : '4','5' : '-1',}
		# week_day = {'1': '0','2': '1','4': '2','8': '3','16': '4','5': '5','64': '6',}

		if recurrence_type == '0':
			l_event.add('rrule', {'freq': 'daily',})
		elif recurrence_type == '1':
			if  recurrence_interval == '1':
				logger.debug('week_day:'+','.join(weekday_list))
				l_event.add('rrule', {'freq': 'weekly', 'byday': weekday_list, })
			else:
				l_event.add('rrule', {'freq': 'weekly', 'interval': recurrence_arr[10].strip('"'), 'byday': weekday_list, })
		elif recurrence_type == '2':
			if  recurrence_interval == '1':
				logger.debug('week_day:'+','.join(weekday_list))
				l_event.add('rrule', {'freq': 'monthly', 'bymonthday': recurrence_dayofmonth, })
			else:
				l_event.add('rrule', {'freq': 'monthly', 'interval': recurrence_arr[10].strip('"'), 'bymonthday': recurrence_dayofmonth, })
		elif recurrence_type == '3':
			logger.debug('rrule added')
			if  recurrence_interval == '1':
				logger.debug('week_day:'+','.join(weekday_list))
				l_event.add('rrule', {'freq': 'monthly', 'byday': instance_dict[recurrence_instance] + weekday_list[0], })
			else:
				l_event.add('rrule', {'freq': 'monthly', 'interval': recurrence_arr[10].strip('"'), 'byday': instance_dict[recurrence_instance] + weekday_list[0], })
		elif recurrence_type == '4':
			logger.debug('rrule added')
		elif recurrence_type == '5':
			logger.debug('rrule added')
			l_event.add('rrule', {'freq': 'yearly', 'bymonth': recurrence_monthofyear, 'bymonthday': recurrence_dayofmonth, })
		elif recurrence_type == '6':
			logger.debug('rrule added')
			if  recurrence_interval == '1':
				logger.debug('week_day:'+','.join(weekday_list))
				l_event.add('rrule', {'freq': 'yearly', 'byday': instance_dict[recurrence_instance] + weekday_list[0], })
			else:
				l_event.add('rrule', {'freq': 'yearly', 'interval': recurrence_arr[10].strip('"'), 'byday': instance_dict[recurrence_instance] + weekday_list[0], })

def process_item(outlook_line, l_event, l_line_number):
	logger.debug('process_item')
	line_len=len(outlook_line)

	if (l_line_number > 2) and (line_len > 1):
		event_arr = outlook_line.split(',')
		l_event = deal_event(event_arr, l_tzinfo)

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
	return l_event

def get_cal_address(a_address):
	l_attendee_email = a_address[a_address.rfind('('):a_address.rfind(')')]
	l_attendee = vCalAddress('MAILTO:'+l_attendee_email)

	l_attendee_cn = a_address[:a_address.rfind('(')]
	l_attendee.params['cn'] = vText(l_attendee_cn)
	return l_attendee

# main
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(format=FORMAT)
logger = logging.getLogger('tcpserver')
logger.setLevel('DEBUG')
logger.warning('Protocol problem: %s', 'connection reset')

cal = Calendar()
cal.add('prodid', '-//My calendar product//mxm.dk//')
cal.add('version', '2.0')

profile_to_process='profil_agglolo_1'
data_directory='../output'
line_number = 0
l_tzinfo="Europe/Paris"

import glob
from os.path import basename

files_items_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.item.*.iconv')

for a_item_file in files_items_arr:
	file_name = basename(a_item_file)
	item_number = file_name.split('.')[3]
	l_event = Event()

	logger.debug('file handled:' + a_item_file)
	f_item = open(a_item_file, 'r')
	line_number=0
	for outlook_line in f_item:
		line_number=line_number+1
		l_event = process_item(outlook_line, l_event, line_number)

		files_recurrences_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.recurrence.*.' + item_number + '.iconv')
		# print file_name
		# print item_number

		for a_recurrence_file in files_recurrences_arr:
			# print a_recurrence_file
			file_rec_name = basename(a_recurrence_file)
			recurrence_number = file_rec_name.split('.')[3]

			f_recurrence = open(a_recurrence_file, 'r')
			line_rec_number=0
			for recurrence_line in f_recurrence:
				line_rec_number=line_rec_number+1
				process_recurrence(recurrence_line, l_event, line_rec_number)
				files_exceptions_arr = glob.glob(data_directory + '/' + profile_to_process + '.csv.exception.*.' + recurrence_number + '.' + item_number + '.iconv')
				for a_exception_file in files_exceptions_arr:
					print a_exception_file
			f_recurrence.close()

	f_item.close()
	cal.add_component(l_event)




import tempfile, os
# directory = tempfile.mkdtemp()
f = open(os.path.join('/home/stlo_agglo/mail_tools/psh', 'example.ics'), 'wb')
f.write(cal.to_ical())
f.close()
