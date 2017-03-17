# Build time: 3/13/2017-3/16/2017

# = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
# SCRIPT FOR CONVERTING OSP DATABASE TABLES (SUBCONTRACTS, INVOICES) 
# FOR UPLOAD INTO NEW SAP-BASED TOOL
# = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

# INPUTS:
# input_subs.txt, input_invs.txt
# OUTPUTS:
# output_subs.txt, output_subs_details.txt, output_invs.txt, output_invs_details.txt
# OPTIONAL FILES:
# subs_include.txt, subs_exclude.txt
# (if provided subs_include, only those subs will be processed, and
#  if proviced subs_exclude, all subs excluding those will be processed)

# Before exporting data from OSP database, run these two queries to remove line breaks:
# UPDATE Subcontracts SET Comm1 = Replace(Replace(Nz([Comm1],""),Chr(10),"; "),Chr(13),"; ");
# UPDATE Invoices SET Comm1 = Replace(Replace(Nz([Comm1],""),Chr(10),"; "),Chr(13),"; ");

# Rules for formatting source data when exporting from Access database:
# 1. Use the pipe character to delimit fields
# 2. Don't use quotes around text
# 3. Don't use headers
# 4. Use Unicode UTF-8 encoding
# 5. Use double digits for months and days, and 4 digits for year

import os
from datetime import datetime, timedelta, date
from time import sleep
from operator import itemgetter # used for sorting lists

print('-' * 51)
print('Program started', ' '*14, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
print('-' * 51)

# clean out any old files
try: os.remove('log.txt')
except OSError: pass

try: os.remove('output_subs.txt')
except OSError: pass

try: os.remove('output_subs_details.txt')
except OSError: pass

try: os.remove('output_invs.txt')
except OSError: pass

try: os.remove('output_invs_details.txt')
except OSError: pass

if not os.path.exists('input_subs.txt') or not os.path.exists('input_invs.txt'):
	print('Cannot find input files input_subs.txt and/or input_invs.txt')
	print('-' * 50); print('Program terminated early'); print('-' * 50)
	exit()

sleep(1)

# open needed files
input_subs = open('input_subs.txt', encoding='utf8')
input_invs = open('input_invs.txt', encoding='utf8')
log_file = open('log.txt', 'a', encoding='utf8')
log_file.write('----- Start: ' + str(datetime.now()) + ' -----\n')
outfile_subs = open('output_subs.txt', 'a', encoding='utf8')
outfile_subs_details = open('output_subs_details.txt', 'a', encoding='utf8')
outfile_invs = open('output_invs.txt', 'a', encoding='utf8')
outfile_invs_details = open('output_invs_details.txt', 'a', encoding='utf8')
if os.path.exists('subs_include.txt'):
	subs_include = open('subs_include.txt', 'r', encoding='utf8')
else:
	subs_include = None
if os.path.exists('subs_exclude.txt'):
	subs_exclude = open('subs_exclude.txt', 'r', encoding='utf8')
else:
	subs_exclude = None

budget_items = [
	('691641', '697141'), # salary
	('691642', '697142'), # fringe
	('691645', '697145'), # supplies
	('691650', '697150'), # travel
	('691643', '697143'), # consulting
	('691648', '697148'), # odc
	('691658', '697158'), # idc
	('691647', '697147'), # equipment
	('691659', '697159')  # misc
]

def padded_text(text, taken, total=50):
	'''
	Accepts a string and returns the same string padded with leading
	white so the length of the entire line is equal to (or no less than)
	the given total length.
	
	This is to make printed statements look better on the console.
	'''
	padding = ''
	text = str(text)
	text_len = len(text)
	if taken + text_len > total:
		return text
	else:
		padding_len = total - (taken + text_len)
		padding = ' ' * padding_len
		return padding + text

def text_to_float(text):
	'''
	Accepts a text that represents a dollar amount, cleans it up 
	and returns the float value or 0.
	'''
	text = text.strip().strip('$')
	if '(' in text:
		text = text.strip('(').strip(')').strip('$')
		text = '-' + text
	if text:
		return float(text)
	else:
		return 0

def text_to_date(text):
	'''
	Accepts a text that represents a date, cleans it up 
	and returns the date object.
	'''
	text = text.strip()[0:10]
	return datetime.strptime(text, '%m/%d/%Y')

def check_budget_periods(sub):
	'''
	Accepts a sub and returns a two-element tuple where the first
	value is the number of valid budget periods found for this sub,
	and the second value is either None (if the first value = 0) or
	the same or updated sub record (updated if found and fixed any
	missing start dates).
	'''
	valid_periods = 0
	# initial_sub = sub
	sub = sub.strip()
	sub = sub.split('|')
	per1_start = sub[24].strip()
	per1_end = sub[30].strip()
	if not per1_start	or not per1_end:
		return 0, None
	else:
		valid_periods	+= 1
		# check remaining periods
		for i in range(0, 5):
				start = sub[25 + i].strip()
				end = sub[31 + i].strip()
				if start and end:
					valid_periods	+= 1
				elif not start and end:
					# fix the start date using a prior period's end date + 1 day:
					prior_end = sub[30 + i][0:10]
					# print('sub:', sub[0])
					prior_end = datetime.strptime(prior_end, '%m/%d/%Y')
					new_start = prior_end	+ timedelta(days = 1)
					new_start	= new_start.strftime('%m/%d/%Y %H:%M:%S')
					sub[25 + i] = new_start
					valid_periods	+= 1
				else:
					# missing an end date (and possibly missing a start date)
					# I can stop checking further periods but let's do Nate
					# a favor and check further periods and see if they seem to exist
					while i < 4:
						i += 1
						start = sub[25 + i].strip()
						end = sub[31 + i].strip()
						if start or end:
							# an odd period possibly exists - let Nate know about this sub
							log_file.write('Sub '+sub[1]+' may have periods ignored by tool\n')
					break
	sub.append(str(valid_periods))
	sub = '|'.join(sub)
	return(valid_periods, sub)

def get_gl_bucket(gl_break, prior_exp, this_amt):
	'''
	Determines which GL bucket this_amt belongs to:
	# 1 = 6916xx, 2 = 6971xx, or 3=both
	'''
	if gl_break >= prior_exp + this_amt:
		gl_bucket = 1
	elif gl_break <= prior_exp:
		gl_bucket = 2
	else:
		gl_bucket = 3
	return gl_bucket

# read all subs into a variable
to_print = 'Reading subaward records...'
print(to_print, end='')
raw_subs = input_subs.readlines()
count = len(raw_subs)
print(padded_text(count, len(to_print)))

# read all invoices into a variable
to_print = 'Reading invoice records...'
print(to_print, end='')
raw_invs = input_invs.readlines()
count = len(raw_invs)
print(padded_text(count, len(to_print)))

# check that sub records have correct number of fields (96)
to_print = 'Checking subawards...'
print(to_print, end='')
for sub in raw_subs:
	sub = sub.strip()
	sub = sub.split('|')
	if len(sub) != 96:
		print(padded_text('Errors found', len(to_print)))
		print('Sub id', sub[0], 'has odd number of fields:', len(sub))
		print('-' * 51); print('Program terminated early'); print('-' * 51)
		exit()
print(padded_text('OK', len(to_print)))

# check that inv records have correct number of fields (35)
to_print = 'Checking invoices...'
print(to_print, end='')
for inv in raw_invs:
	inv = inv.strip()
	inv = inv.split('|')
	if len(inv) != 35:
		print(padded_text('Errors found', len(to_print)))
		print('Invoice id', inv[0], 'has odd number of fields: ', len(inv))
		print('-' * 51); print('Program terminated early'); print('-' * 51)
		exit()
print(padded_text('OK', len(to_print)))

# cleanup subs that are outside the 2000000-3999999 range
to_print = 'Removing unneeded subs...'
print(to_print, end='')
i = 0
cleaner_subs = []
remaining_sub_ids = []
for sub in raw_subs:
	raw_sub = sub # remember the raw version
	sub = sub.strip()
	sub = sub.split('|')
	wbse_first_char = int(sub[1][0:1])
	if wbse_first_char < 2 or wbse_first_char > 3:
		i += 1
	else:
		remaining_sub_ids.append(sub[0])
		cleaner_subs.append(raw_sub)
print(padded_text(i, len(to_print)))

# cleanup
subs = cleaner_subs
del cleaner_subs
del raw_subs

to_print = 'Removing unneeded invoices...'
print(to_print, end='')
i = 0
clean_invs = []
for inv in raw_invs:
	raw_inv = inv.strip()
	inv = inv.strip()
	inv = inv.split('|')
	sub_id = inv[1]
	if sub_id in remaining_sub_ids:
		clean_invs.append(raw_inv)
	else:
		i += 1
print(padded_text(i, len(to_print)))

# cleanup
invs = clean_invs
del clean_invs
del raw_invs
del remaining_sub_ids

to_print = 'Fixing period start dates...'
print(to_print, end='')
i = 0
num_periods = 0
clean_subs = []
dropped_subs = []
for sub in subs:
	valid_periods, updated_sub = check_budget_periods(sub)
	if valid_periods:
		clean_subs.append(updated_sub)
		i += 1
	else:
		dropped_subs.append(sub)
print(padded_text(i, len(to_print	)))
to_print = 'Subs dropped for not having any periods...'
print(to_print, end='')
print(padded_text(len(dropped_subs), len(to_print)))

# cleanup
subs = clean_subs
del clean_subs
del dropped_subs

# remove inactive subs (end date < 7/1/2015)
to_print = 'Removing inactive subs...'
print(to_print, end='')
i = 0
clean_subs = []
remaining_sub_ids = []
for sub in subs:
	raw_sub = sub # remember the raw version
	sub = sub.strip()
	sub = sub.split('|')
	index = 29 + int(sub[96])
	last_per_end = sub[index][0:10]
	last_per_end = datetime.strptime(last_per_end, '%m/%d/%Y')
	if last_per_end < datetime(year=2015, month=7, day=1):
		i += 1
	else:
		remaining_sub_ids.append(sub[0])
		clean_subs.append(raw_sub)
print(padded_text(i, len(to_print)))

# cleanup
subs = clean_subs
del clean_subs

to_print = 'Removing inactive invoices...'
print(to_print, end='')
i = 0
clean_invs = []
for inv in invs:
	raw_inv = inv.strip()
	inv = inv.strip()
	inv = inv.split('|')
	sub_id = inv[1]
	if sub_id in remaining_sub_ids:
		clean_invs.append(raw_inv)
	else:
		i += 1
print(padded_text(i, len(to_print)))

# cleanup
invs = clean_invs
del clean_invs
del remaining_sub_ids

# check for include/exclude list
# if include list exists, only include those subs (and their invs) in the output
# if exclude list exists, excluding those subs
if subs_include:
	subs_include = subs_include.readlines()
	clean_subs_include = []
	for sub in subs_include:
		sub = sub.strip()
		clean_subs_include.append(sub)
	# cleanup
	subs_include = clean_subs_include
	del clean_subs_include

	to_print = 'Subs to include exclusively...'
	print(to_print, end='')
	print(padded_text(len(subs_include), len(to_print)))

	clean_subs = []
	sub_ids = []
	for sub in subs:
		sub = sub.split('|')
		if sub[1] in subs_include:
			sub_ids.append(sub[0])
			sub = '|'.join(sub)
			clean_subs.append(sub)
	# cleanup
	subs = clean_subs
	del clean_subs

	clean_invs = []
	for inv in invs:
		inv = inv.split('|')
		if inv[1] in sub_ids:
			inv = '|'.join(inv)
			clean_invs.append(inv)
	# cleanup
	invs = clean_invs
	del clean_invs

elif subs_exclude:
	subs_exclude = subs_exclude.readlines()
	clean_subs_exclude = []
	for sub in subs_exclude:
		sub = sub.strip()
		clean_subs_exclude.append(sub)
	# cleanup
	subs_exclude = clean_subs_exclude
	del clean_subs_exclude

	to_print = 'Subs to exclude...'
	print(to_print, end='')
	print(padded_text(len(subs_exclude), len(to_print)))

	clean_subs = []
	sub_ids = []
	for sub in subs:
		sub = sub.split('|')
		if not sub[1] in subs_exclude:
			sub_ids.append(sub[0])
			sub = '|'.join(sub)
			clean_subs.append(sub)
	# cleanup
	subs = clean_subs
	del clean_subs

	clean_invs = []
	for inv in invs:
		inv = inv.split('|')
		if inv[1] in sub_ids:
			inv = '|'.join(inv)
			clean_invs.append(inv)
	# cleanup
	invs = clean_invs
	del clean_invs

to_print ='Sorting subs by wbse value...'
print(to_print, end='')
temp_subs = []
for sub in subs:
	sub = sub.split('|')
	temp_subs.append(sub)
temp_subs.sort(key=itemgetter(1))  # sort by wbse/fund code
subs = []
for sub in temp_subs:
	sub = '|'.join(sub)
	subs.append(sub)
print(padded_text('OK', len(to_print)))

to_print = 'Number of subs to convert:'; print(to_print, end='')
print(padded_text(len(subs), len(to_print)))
to_print = 'Number of invoices to convert:'; print(to_print, end='')
print(padded_text(len(invs), len(to_print)))

output2_count = 0
output3_count = 0
output4_count = 0
# output1_count is just the length of subs list

# write file headers
header_subs = 'WBSE|State|Country|Subaward Number|FFATA|Final Invoice Due|G/L Break|Prior Year WBSE|OSP Notes|IDC Default|Received Date|Subrecipient PI Name|Manual Prior Exp|Type of Subaward|Type of Payment|Invoice Requirements|Equipment|Budgetary Changes|Budget Restrictions|Special T&C'
header_subs_details = 'WBSE|Fiscal Period|Fiscal Year|Budget Period Start|Budget Period End|Amount|Category|IDC Rate'
header_invs = 'WBSE|Invoice #|AP Check Request #|Received Date|Final|Treat as Final|Initially Accurate|Vendor|Wire or Draft|Notes|Start Date|End Date|OSP Invoice Type|IDC Rate'
header_invs_details = 'WBSE|Invoice Number|Amount|Cost Element'
outfile_subs.write(header_subs + '\n')
outfile_subs_details.write(header_subs_details + '\n')
outfile_invs.write(header_invs + '\n')
outfile_invs_details.write(header_invs_details + '\n')

for sub in subs:

	out_subs = []
	sub = sub.strip()
	sub = sub.split('|')

	out_subs.append(sub[1].strip())      # wbse
	out_subs.append(sub[20].strip())     # state
	out_subs.append(sub[21].strip())     # country
	out_subs.append(sub[5].strip())      # subaward number
	out_subs.append(sub[6].strip())      # ffata
	out_subs.append(sub[9].strip())      # final invoice due
	gl_break = text_to_float(sub[14])
	out_subs.append(str(gl_break))       # gl break
	out_subs.append(sub[11].strip())     # prior year wbse
	out_subs.append(sub[10].strip())     # osp notes
	out_subs.append('') 						     # skip row
	out_subs.append('') 						     # skip row
	out_subs.append('') 						     # skip row
	prior_exp = text_to_float(sub[13])
	out_subs.append(str(prior_exp))      # manual prior exp

	out_sub = '|'.join(out_subs)
	outfile_subs.write(out_sub + '\n')

	num_periods = int(sub[96])
	first_period = 10 - num_periods

	for i in range(0, num_periods):

		wbse = sub[1].strip()
		fisc_per = str(first_period + i)
		fisc_yr = str(2017)
		start = sub[24 + i].strip()[0:10]
		end = sub[30 + i].strip()[0:10]
		idc_rate = sub[72 + i]
		idc_rate = text_to_float(idc_rate)
		idc_rate_reformatted = str(round((idc_rate * 100), 2))

		line_items = []

		# salary
		salary = sub[36 + i]
		salary = text_to_float(salary)
		if salary:
			salary = str(salary)
			category = '693541'
			line_items.append((salary, category))
		
		# fringe
		fringe = sub[42 + i]
		fringe = text_to_float(fringe)
		if fringe:
			fringe = str(fringe)
			category = '693542'
			line_items.append((fringe, category))

		# supplies
		supplies = sub[48 + i]
		supplies = text_to_float(supplies)
		if supplies:
			supplies = str(supplies)
			category = '693545'
			line_items.append((supplies, category))

		# travel
		travel = sub[54 + i]
		travel = text_to_float(travel)
		if travel:
			travel = str(travel)
			category = '693550'
			line_items.append((travel, category))

		# consulting
		consulting = sub[60 + i]
		consulting = text_to_float(consulting)
		if consulting:
			consulting = str(consulting)
			category = '693543'
			line_items.append((consulting, category))

		# odc (other direct cost)
		odc = sub[66 + i]
		odc = text_to_float(odc)
		if odc:
			odc = str(odc)
			category = '693548'
			line_items.append((odc, category))

		# idc (indirect cost)
		if idc_rate:
			direct_cost =  sum([float(item[0]) for item in line_items])
			adj_amt = sub[78 + i]
			adj_amt = text_to_float(adj_amt)
			if direct_cost:
				idc = direct_cost * idc_rate + adj_amt
				idc = str(round(idc, 2))
				category = '693558'
				line_items.append((idc, category))

		# equipment
		equipment = sub[84 + i]
		equipment = text_to_float(equipment)
		if equipment:
			equipment = str(equipment)
			category = '693547'
			line_items.append((equipment, category))

		# misc
		misc = sub[90 + i]
		misc = text_to_float(misc)
		if misc:
			misc = str(misc)
			category = '693559'
			line_items.append((misc, category))

		# write to file
		for amount, category in line_items:
			sub_detail = [wbse, fisc_per, fisc_yr, start, end, amount, category, idc_rate_reformatted]
			sub_detail = '|'.join(sub_detail)
			outfile_subs_details.write(sub_detail + '\n')
			output2_count += 1

	# collect invoices for the given sub
	sub_invs = []
	sub_id = sub[0]
	for inv in invs:
		inv = inv.strip()
		inv = inv.split('|')
		if inv[1] == sub_id:
			inv = '|'.join(inv)
			sub_invs.append(inv)

	if sub_invs:

		# convert a couple elements to float/date for sorting on them later
		new_sub_invs = []
		for inv in sub_invs:
			inv = inv.split('|')
			inv[0] = text_to_float(inv[0])   # id
			inv[24] = text_to_date(inv[24])  # end date
			new_sub_invs.append(inv)
		# cleanup
		sub_invs = new_sub_invs
		del new_sub_invs

		# sort invoices by end date, then by ID
		sub_invs.sort(key=itemgetter(0))   # secondary sort (by ID)
		sub_invs.sort(key=itemgetter(24))  # primary sort (by end date)

		for inv in sub_invs:

			# collect all line items (salary - misc)
			line_items = []
			for i in range(25,35):
				item = inv[i]
				item = text_to_float(item)
				line_items.append(item if item else 0)

			direct_cost = sum(line_items[0:6])
			idc = direct_cost * line_items[6] + line_items[7]
			total_inv_amount = direct_cost + idc + sum(line_items[8:])

			if not total_inv_amount:
				log_file.write('Inv ' + str(inv[0]) + ' skipped because total amt = 0' + '\n')
				continue

			line_items[6] = idc # replace idc_rate with idc amount
			line_items.pop(7)   # remove adj item

			out_invs = []

			out_invs.append(sub[1].strip())         # wbse
			out_invs.append(inv[2].strip())         # invoice number
			out_invs.append(inv[3].strip())         # ap check req number
			out_invs.append(inv[6].strip()[0:10])   # received date
			out_invs.append(inv[10].strip())        # final
			out_invs.append('')                     # skip row (treat as final)
			out_invs.append(inv[11].strip())        # initially accurate
			out_invs.append('')                     # skip row (vendor)
			out_invs.append('')                     # skip row (wire draft)
			out_invs.append(inv[9].strip())         # notes
			out_invs.append(inv[23].strip()[0:10])  # start date
			end_date = inv[24].strftime('%m/%d/%Y')
			out_invs.append(end_date)               # end date
			out_invs.append('DB') 						      # skip row (osp invoice type)
			idcr = text_to_float(inv[31])
			idcr_reformatted = round((idcr * 100), 2)
			out_invs.append(str(idcr_reformatted))  # idc rate

			out_inv = '|'.join(out_invs)
			outfile_invs.write(out_inv + '\n')
			output3_count += 1

			wbse = sub[1].strip()
			inv_num = inv[2].strip()
			
			# gl_bucket: 1 = 6916xx, 2 = 6971xx, or 3=both

			for index, amount in enumerate(line_items):
				if amount:
					amount = round(amount, 2)
					gl_bucket = get_gl_bucket(gl_break, prior_exp, amount)

					if gl_bucket == 1 or gl_bucket == 2:
						cost_elem = budget_items[index][gl_bucket - 1]
						inv_detail = [wbse, inv_num, str(amount), cost_elem]
						inv_detail = '|'.join(inv_detail)
						outfile_invs_details.write(inv_detail + '\n')
						output4_count += 1
					else:
						amount_1 = round((gl_break - prior_exp), 2)
						amount_2 = round((amount - amount_1), 2)
						# gl bucket 1
						cost_elem = budget_items[index][0]  
						inv_detail = [wbse, inv_num, str(amount_1), cost_elem]
						inv_detail = '|'.join(inv_detail)
						outfile_invs_details.write(inv_detail + '\n')						
						# gl bucket 2
						cost_elem = budget_items[index][1]
						inv_detail = [wbse, inv_num, str(amount_2), cost_elem]
						inv_detail = '|'.join(inv_detail)
						outfile_invs_details.write(inv_detail + '\n')
						output4_count += 2

					prior_exp += amount

to_print = 'Output records created - subs:'; print(to_print, end='')
print(padded_text(len(subs), len(to_print)))
to_print = 'Output records created - subs_details:'; print(to_print, end='')
print(padded_text(output2_count, len(to_print)))
to_print = 'Output records created - invs:'; print(to_print, end='')
print(padded_text(output3_count, len(to_print)))
to_print = 'Output records created - invs_details:'; print(to_print, end='')
print(padded_text(output4_count, len(to_print)))

log_file.write('----- End:   ' + str(datetime.now()) + ' -----\n')

print('-' * 51)
print('Program completed', ' '*12, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
print('-' * 51)
