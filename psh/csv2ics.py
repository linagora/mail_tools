# -*- coding: utf-8 -*-

from icalendar import Calendar, Event, vCalAddress, vText
import pytz
from datetime import datetime

def split_outlook_date(a_outlook_date):
	date_without_quote = a_outlook_date.strip('"')
	day_part=date_without_quote[:date_without_quote.find(' ')-1]
	hour_part=date_without_quote[date_without_quote.find(' ')+1:]

	a_day_arr=day_part.split('/')
	a_hour_arr=hour_part.split(':')

	return int(a_day_arr[0]), int(a_day_arr[1]), int(a_day_arr[2]), int(a_hour_arr[0]), int(a_hour_arr[1]), int(a_hour_arr[2]) 

def deal_event(a_event_arr, a_tzinfo):
	# print event_arr
	print "deal event"
	event = Event()

	# summary - Body
	event.add('summary', event_arr[11])

	# dtstamp - Start
	a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(event_arr[56])
	event.add('dtstart', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

	a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(event_arr[33])
	event.add('dtend', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(a_tzinfo)))

	# event.add('dtstamp', datetime(2005, 4, 4, 0, 10, 0, tzinfo=pytz.utc))

	organizer_field = event_arr[43]
	organizer=get_cal_address(organizer_field)
	organizer.params['role'] = vText('CHAIR')

	event['organizer'] = organizer

	event_location = event_arr[36][1:len(event_arr[36])-1]
	# print event_location
	event['location'] = vText(event_location)

	#event['uid'] = '20050115T101010/27346262376@mxm.dk'
	event['uid'] = event_arr[10][1:len(event_arr[10])-1]

	event_importance = event_arr[16][1:len(event_arr[16])-1]
	# print event_importance
	event.add('priority', event_importance)

	required_attendees = event_arr[52][1:len(event_arr[52])-1].split(';')
	# print event_arr[52]
	# print required_attendees
	for a_required in required_attendees:
		a_cal_required_attendee = get_cal_address(a_required)
		a_cal_required_attendee.params['ROLE'] = vText('REQ-PARTICIPANT')
		event.add('attendee', a_cal_required_attendee, encode=0)

	is_recurring = event_arr[35]
	print "deal event:"+is_recurring
	if is_recurring == '"True"':
		print "rrule"
		event.add('rrule', {'freq': 'daily'})

	return event

def get_cal_address(a_address):
	l_attendee_email = a_address[a_address.rfind('('):a_address.rfind(')')]
	l_attendee = vCalAddress('MAILTO:'+l_attendee_email)

	l_attendee_cn = a_address[1:a_address.rfind('(')]
	l_attendee.params['cn'] = vText(l_attendee_cn)
	return l_attendee

cal = Calendar()
cal.add('prodid', '-//My calendar product//mxm.dk//')
cal.add('version', '2.0')

f = open('../output/profil_agglolo_1_withrecurrences_0.5.csv.iconv', 'r')
recurring_events_dict = dict()
line_number = 0
l_tzinfo="Europe/Paris"

for outlook_line in f:
	# print len(outlook_line)
	line_len=len(outlook_line)
	line_number=line_number+1
	l_event = Event()

	if (line_number > 2) and (line_len > 1):
		event_arr = outlook_line.split(',')

		is_recurring = event_arr[35]
		event_conversation_index = event_arr[10]
		print is_recurring
		print event_conversation_index
		if is_recurring == '"True"':
			if event_conversation_index in recurring_events_dict:
				# evenement modifie
				recurrence_state = event_arr[45]
				if recurrence_state == '"3"':
					l_event = recurring_events_dict[event_conversation_index]
					a_day, a_month, a_year, a_hour, a_minute, a_second = split_outlook_date(event_arr[56])
					l_event.add('exdate', datetime(a_year, a_month, a_day, a_hour, a_minute, a_second, tzinfo=pytz.timezone(l_tzinfo)))
			else:
				l_event = deal_event(event_arr, l_tzinfo)
				recurring_events_dict[event_conversation_index] = l_event
		else:
			l_event = deal_event(event_arr, l_tzinfo)
			recurring_events_dict[event_conversation_index] = l_event

		cal.add_component(l_event)

f.close()

import tempfile, os
# directory = tempfile.mkdtemp()
f = open(os.path.join('/home/stlo_agglo/mail_tools/psh', 'example.ics'), 'wb')
f.write(cal.to_ical())
f.close()
