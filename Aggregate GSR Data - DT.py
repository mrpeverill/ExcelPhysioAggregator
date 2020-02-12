# Zachary Kirshbaum
# 3/16/17
# Converted to Python 3 and updated - Matthew Peverill 2/12/2020

# Aggregates DT Physio data to one wide-format
# file for each scoring parameter

from __future__ import division
import glob
import pathlib
import collections
import csv
import os
import re
import sys
import zipfile

from xml.etree.ElementTree import iterparse

######## CUSTOMIZATION
# I've tried to put everything you'll need to tweak for a new environment or dataset here:

script_directory = os.path.dirname(os.path.realpath(__file__))
p1_excel_folder = os.path.join(script_directory,"SCORING PARAMETER 1","EXCEL OUTPUTS") # This should be a list of directories pointing to parameter 1 excel files
p2_excel_folder = os.path.join(script_directory,"SCORING PARAMETER 2","EXCEL OUTPUTS") # This should be a list of directories pointing to parameter 2 excel files

#######


# add one subject's data to aggregate data
def addSubject(target_file, aggregate_data, variables, phases):
    
    match = re.search(r'/\d\d\d\d',target_file.name)
    if match:
        subject_id = match.group(1)
    else:
        print("No subject number found in %s" % target_file.name)
        raise ValueError

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

    translated_data = translateXML(target_file)

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

phases_gen = {
    2:'A Marker CS+', # (11 trials)
    4:'B Marker CS-', # (11 trials)
    8:'Noise Alone', # (14 trials)
    128:'Air Puff' # (9 trials)
}

# phases_acq = {
    # 2:'A Marker CS+', # (11 trials)
    # 4:'B Marker CS-', # (11 trials)
    # 8:'Noise Alone', # (14 trials)
    # 128:'Air Puff' # (9 trials)
# }

# phases_ext = {
    # 2:'A Marker CS+', # (12 trials)
    # 4:'B Marker CS-', # (12 trials)
    # 8:'Noise Alone', # (15 trials)
    # 128:'Air Puff' # (0 trials)
# }

subjects_P1 = collections.OrderedDict()
subjects_P2 = collections.OrderedDict()

print(p1_excel_folder)

GSR_files_P1 = list(pathlib.Path(p1_excel_folder).glob('*.xlsx'))
print(GSR_files_P1)
GSR_files_P2 = list(pathlib.Path(p2_excel_folder).glob('*.xlsx'))

if (len(GSR_files_P1) == 0 and len(GSR_files_P2) == 0):
        print('No subject files found')
        print('Press enter to close...')
        input()
        sys.exit()

for f in sorted(GSR_files_P1, key=lambda s:s.name.lower()): # We want to sort by the lower case translation of the filename.
    try:
        addSubject(f, subjects_P1, variables, phases_gen)
    except ValueError:
        pass

for f in sorted(GSR_files_P2, key=lambda s:s.name.lower()):
    try:
        addSubject(f, subjects_P2, variables, phases_gen)
    except ValueError:
        pass


writeAggregateData(subjects_P1, script_directory + '/' + p1_dir + '/GSR_Wide_Aggregate_P1.csv')
writeAggregateData(subjects_P2, script_directory + '/' + p2_dir + '/GSR_Wide_Aggregate_P2.csv')