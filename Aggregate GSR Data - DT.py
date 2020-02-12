# Zachary Kirshbaum
# 3/16/17

# Aggregates DT Physio data to one wide-format
# file for each scoring parameter

from __future__ import division

import collections
import csv
import os
import re
import sys
import zipfile

from xml.etree.ElementTree import iterparse

# add one subject's data to aggregate data
def addSubject(subject_filename, aggregate_data, variables, phases):
	filename = subject_filename.split('/')[-1]
	filename = re.findall("\d+", filename)

	subject_id = filename[0]

	# add to data
	if (subject_id in aggregate_data.keys()):
		subject = aggregate_data.get(subject_id)
	# create new data
	else:
		subject = collections.OrderedDict()

		subject['SEA ID'] = subject_id
		subject['Task Completed'] = 0

		for key in sorted(phases.keys()):
			subject[phases.get(key)] = 0

	translated_data = translateXML(subject_filename)

	# check completion
	first_row = True
	for row in translated_data:
		if (not first_row):
			if (row.values()[0] != ''):
				try:
					stim_label = phases.get(int(row.get('H')))
				except:
					pass
				if ("." not in row.values() and ". " not in row.values() and '="."' not in row.values()):
					subject[stim_label] += 1
					subject['Task Completed'] = 1
		else:
			first_row = False
	
	# record values
	for letter in sorted(variables.keys()):
		first_row = True
		previous_stim_label = ''
		stim_times = {}
		for row in translated_data:
			if (not first_row):
				if (not row.get('A') == '.'):
					try:
						stim_label = phases.get(int(row.get('H')))
						value = row.get(letter)
						if (value == "." or value == ". "):
							value = '="."'
						if (not stim_label == previous_stim_label):
							trial = 1
							completion_checked = False
							for stim_time in sorted(stim_times):
								subject[previous_stim_label + str(trial) + '_' + variables.get(letter)] = stim_times.get(stim_time)
								trial += 1
							stim_times = {}
							previous_stim_label = stim_label
						stim_times[float(row.get('A'))] = value
					except:
						pass
			else:
				first_row = False
		trial = 1

		for stim_time in sorted(stim_times):
			subject[previous_stim_label + str(trial) + '_' + variables.get(letter)] = stim_times.get(stim_time)
			trial += 1

	aggregate_data[subject_id] = subject
	
# writes all data to a master file
def writeAggregateData(subjects, output_file):
	output_file = open(output_file, 'wb')
	csv_writer = csv.writer(output_file)
	labels = subjects.get(subjects.keys()[0]).keys()
	for subj in subjects.values():
		if (len(subj.keys()) > len(labels)):
			labels = subj.keys()
	csv_writer.writerow(labels)

	for subject_name in subjects:
		subject = subjects.get(subject_name)
		subject_row = []
	
		for label in labels:
			subject_row.append(subject.get(label))
	
		csv_writer.writerow(subject_row)

	output_file.close()
	
# take data from Excel xml file
def translateXML(xml_file):
	z = zipfile.ZipFile(xml_file)
	strings = [el.text for e, el in iterparse(z.open('xl/sharedStrings.xml')) if el.tag.endswith('}t')]
	rows = []
	row = collections.OrderedDict()
	value = ''

	for e, el in iterparse(z.open('xl/worksheets/sheet2.xml')):
	    if el.tag.endswith('}v'): # <v>84</v>
	        value = el.text
	    if el.tag.endswith('}c'): # <c r="A3" t="s"><v>84</v></c>
	        if el.attrib.get('t') == 's':
	            value = strings[int(value)]
	        letter = el.attrib['r'] # AZ22
	        while letter[-1].isdigit():
	            letter = letter[:-1]
	        row[letter] = value
	        value = ''
	    if el.tag.endswith('}row'):
	        rows.append(row)
	        row = collections.OrderedDict()

	z.close()

	return rows

variables = {
	# 'A':'Stim Time'
	'B':'SCL'
	,'C':'Latency'
	,'D':'SCRAmplitude'
	,'E':'SCRRiseTime'
	,'F':'SCRSize'
	,'G':'SCROnset'
	# ,'H':'StimLabel'
}

phases_acq = {
	2:'A Marker CS+', # (11 trials)
	4:'B Marker CS-', # (11 trials)
	8:'Noise Alone', # (14 trials)
	128:'Air Puff' # (9 trials)
}

phases_ext = {
	2:'A Marker CS+', # (12 trials)
	4:'B Marker CS-', # (12 trials)
	8:'Noise Alone', # (15 trials)
	128:'Air Puff' # (0 trials)
}

subjects_P1 = collections.OrderedDict()
subjects_P2 = collections.OrderedDict()

script_directory = os.path.dirname(os.path.realpath(__file__))
param_dirs = os.listdir(script_directory)
p1_dir = ''
p2_dir = ''
for item in param_dirs:
        if ('PARAMETER 1' in item.upper()):
                p1_dir = item
        elif ('PARAMETER 2' in item.upper()):
                p2_dir = item

p1_contents = os.listdir(script_directory + '/' + p1_dir)
p1_excel_folder = ''

for item in p1_contents:
        if ('EXCEL' in item.upper() and 'OUTPUTS' in item.upper()):
                p1_excel_folder = item

p2_contents = os.listdir(script_directory + '/' + p2_dir)
p2_excel_folder = ''

for item in p2_contents:
        if ('EXCEL' in item.upper() and 'OUTPUTS' in item.upper()):
                p2_excel_folder = item

GSR_files_P1 = [f for f in os.listdir(script_directory + '/' + p1_dir + '/' + p1_excel_folder + '/') if f.endswith('.xlsx')]
print GSR_files_P1
GSR_files_P2 = [f for f in os.listdir(script_directory + '/' + p2_dir + '/' + p2_excel_folder + '/') if f.endswith('.xlsx')]

if (len(GSR_files_P1) == 0 and len(GSR_files_P2) == 0):
        print('No subject files found')
        print('Press enter to close...')
        raw_input()
        sys.exit()

for f in sorted(GSR_files_P1, key=lambda s:s.lower()):
        try:
                addSubject(script_directory + '/' + p1_dir + '/' + p1_excel_folder + '/' + f, subjects_P1, variables, phases)
        except:
                pass

for f in sorted(GSR_files_P2, key=lambda s:s.lower()):
        try:
                addSubject(script_directory + '/' + p2_dir + '/' + p2_excel_folder + '/' + f, subjects_P2, variables, phases)
        except:
                pass

writeAggregateData(subjects_P1, script_directory + '/' + p1_dir + '/GSR_Wide_Aggregate_P1.csv')
writeAggregateData(subjects_P2, script_directory + '/' + p2_dir + '/GSR_Wide_Aggregate_P2.csv')