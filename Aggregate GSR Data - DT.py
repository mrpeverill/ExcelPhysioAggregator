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

# ignore_phases is set below and says which phases should be ignored by the program
# the match regex might need to be changed if the subject number is no longer just 4 digits.
#######

class Error(Exception):
   """Base class for other exceptions"""
   pass

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

class NoIDinFileError(Error):
   """Raised when the subject ID was not found"""
   pass
   
# add one subject's data to aggregate data
def addSubject(target_file, aggregate_data, variables, phases):
    print('Reading %s' % target_file.name)
    match = re.search('\d\d\d\d',target_file.name) # SUBJECTID REGEX. \d\d\d\d is 4 digits
    #match = re.search('5002',target_file.name) # SUBJECTID REGEX. \d\d\d\d is 4 digits
    if match is not None:
        subject_id = match.group()
    else:
        #print("No subject number found in file %s" % target_file.name)
        raise NoIDinFileError

    # Load the subject from aggregate OR make a new entry
    if (subject_id in aggregate_data.keys()):
        subject = aggregate_data.get(subject_id)
    else:
        subject = collections.OrderedDict()

        subject['ID'] = subject_id
        subject['Task Completed'] = 0

        for key in sorted(phases.keys()):
            subject[phases.get(key)+'_total'] = 0
            subject[phases.get(key)+'_matched'] = 0

    translated_data = translateXML(target_file) # This creates a list of OrderedDict objects (one for each row). Keys are column letter. Not in numpy probably because the tables aren't the same in edited v unedited xlsx?
    #Filter the data we want to work with: 
    filtered_data = [r for r in translated_data[1:] if ( 
                                                            'A' in r and 'H' in r and # both keys should exist
                                                            isfloat(r['A']) and isfloat(r['H']) and int(r['H']) in phases  #values should be convertable to floats. Ignore ignored phases.
                                                       )]

    for row in filtered_data: # Go through row by row.
        stim_label = phases[int(row['H'])]
        subject[stim_label + '_total'] += 1 # Count the total number of trials recorded
        if ("." not in row.values() and ". " not in row.values() and '="."' not in row.values()): # Count the trial complete if there are no missing values.
                subject[stim_label + '_matched'] += 1 # Count the number of valid trials per stimulus
                subject['Task Completed'] = 1 # Mark completed if we have at least one valid row?
        trial = 2+(len([i for i in subject.keys() if stim_label in i])//len(variables)) # 2 matches will be trial counts, then we want to know how many sets of variables there are for each label.
        for letter in sorted(variables.keys()):
            stim_times = {}
            value = row[letter].strip()
            #print("Writing %s , trial %s, value %s" % (stim_label, trial, value))
            stim_times[float(row['A'])] = value
            for stim_time in sorted(stim_times):
                    subject[stim_label + str(trial) + '_' + variables[letter]] = stim_times[stim_time]
        stim_times[float(row['A'])] = value
    
    aggregate_data[subject_id] = subject
    
# writes all data to a master file
def writeAggregateData(subjects, output_file):
    #print(subjects)
    output_file = open(output_file, 'w',newline='')
    csv_writer = csv.writer(output_file)
    
    firstsubject=list(subjects)[0]
    labels = subjects[firstsubject].keys() # Get the keys from the first subject.
    
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

GSR_files_P1 = list(pathlib.Path(p1_excel_folder).glob('*.xlsx'))
GSR_files_P2 = list(pathlib.Path(p2_excel_folder).glob('*.xlsx'))

if (len(GSR_files_P1) == 0 and len(GSR_files_P2) == 0):
        print('No subject files found')
        print('Press enter to close...')
        input()
        sys.exit()

for f in sorted(GSR_files_P1, key=lambda s:s.name.lower()): # We want to sort by the lower case translation of the filename.
    try:
        addSubject(f, subjects_P1, variables, phases_gen)
    except NoIDinFileError:
        pass

for f in sorted(GSR_files_P2, key=lambda s:s.name.lower()):
    try:
        addSubject(f, subjects_P2, variables, phases_gen)
    except NoIdinFileError:
        pass


writeAggregateData(subjects_P1, 'GSR_Wide_Aggregate_P1_test.csv')
writeAggregateData(subjects_P2, 'GSR_Wide_Aggregate_P2_test.csv')