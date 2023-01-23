import pandas as pd
import os
import subprocess
from tqdm import tqdm

path_to_data = '/Users/colmconnaughton/Warwick/DSSG/msforms/'
filename = 'applications.xlsx'
path_to_md_output = '/Users/colmconnaughton/Warwick/DSSG/msforms/md/'
path_to_pdf_output = '/Users/colmconnaughton/Warwick/DSSG/msforms/pdfs/'

data = pd.read_excel(path_to_data+filename, index_col=0)

titles = list(data.columns)
tags = ['START_TIME', 'COMPLETE_TIME', 'EMAIL', 'NAME', 'FORENAME', 'SURNAME', 'PREFERRED_NAME',
        'EMAIL2', 'PHONE', 'ADDRESS', 'COUNTRY', 'SHORT_INTRO', 'GENDER', 'DSSG_AT', 'UNIVERSITY', 'UNI_COUNTRY',
        'MAJOR', 'LEVEL', 'YEAR', 'GRADUATION', 'TRANSCRIPTS',
        'PROGRAMMING', 'CS-ALGORITHMS', 'STATS','ML','SOCSCI','REAL_WORLD','EXPT_DESIGN','ETL', 'VISUALISATION', 'PYTHON',
        'R','C','OTHER_PROG','PROGRAMMING_EXP',
        'SQL', 'GIS', 'SAS', 'SPSS', 'STATA', 'DATA_ANALYSIS_EXP',
        'REGRESSION', 'DECISION_TREES', 'SVMS', 'RANDOM_FORESTS', 'NN', 'TIME_SERIES', 'UNSUPERVISED_MODELS',
        'SEMISUPERVISED_MODELS', 'GRAPHICAL_MODELS', 'ML_EXP',
        'CAUSAL', 'MATCHING', 'INSTRUMENTAL', 'REGRESSION_DISCONTINUITY', 'NATURAL_EXPTS', 'SOCSCI_EXP',
        'TEXT_FILES', 'RELATIONAL_DBS', 'NLP', 'GRAPH_DB', 'MULTIMEDIA', 'SENSORS', 'BIG_DATA', 'GEOSPATIAL', 'DATA_EXP',
        'PROJECTS', 'GITHUB', 'MOTIVATION', 'TEAMWORK_EXP', 'FUTURE',
        'HOW_HEAR', 'PREVIOUS_APPL', 'ANYTHING_ELSE', 'CV', 'AGREE_TCS'
         ]

fields = {tags[k]:item for k,item in enumerate(titles[0:len(tags)])}


def write_section(sectags):
  for k in sectags:
        f.write("**"+fields[k]+"** : ")
        f.write(str(data.iloc[id][fields[k]]))
        f.write('\n\n')

failures = []
for id in tqdm(range(len(data))):
#for id in tqdm(range(5)):
#for id in tqdm([401]):
    # Open separate file for each application
    md_filename = path_to_md_output+str(id+1)+'.md'
    f = open(md_filename, "w")

    # Write sections in human-readable markdown
    f.write("## Personal details:\n\n")
    A = ['FORENAME', 'SURNAME', 'PREFERRED_NAME',
        'EMAIL2', 'PHONE', 'ADDRESS', 'COUNTRY', 'SHORT_INTRO', 'GENDER', 'DSSG_AT', 'UNIVERSITY', 'UNI_COUNTRY',
        'MAJOR', 'LEVEL', 'YEAR', 'GRADUATION']
    write_section(A)

    # Hide the URL for CV in a link since most of them are long
    link = str(data.iloc[id][fields['CV']])
    f.write("**"+fields['CV']+"** : ")
    if link != 'nan':
        f.write("[click here](")
        f.write(link)
        f.write(')\n\n')
    else:
        f.write('nan\n\n')
    # Hide the URL for transcripts in a link since most of them are long
    link = str(data.iloc[id][fields['TRANSCRIPTS']])
    f.write("**"+fields['TRANSCRIPTS']+"** : ")
    if link != 'nan':
        f.write("[click here](")
        f.write(link)
        f.write(')\n\n')
    else:
        f.write('nan\n\n')



    f.write("## Expertise:\n\n")
    A = ['PROGRAMMING', 'CS-ALGORITHMS', 'STATS','ML','SOCSCI','REAL_WORLD','EXPT_DESIGN','ETL', 'VISUALISATION']
    write_section(A)

    A = ['PROGRAMMING_EXP', 'PYTHON','R','C','OTHER_PROG']
    f.write("## Programming:\n\n")
    write_section(A)

    f.write("## Statistics software:\n\n")
    A = ['DATA_ANALYSIS_EXP', 'SQL', 'GIS', 'SAS', 'SPSS', 'STATA']
    write_section(A)

    f.write("## Machine learning:\n\n")
    A = ['ML_EXP','REGRESSION', 'DECISION_TREES', 'SVMS', 'RANDOM_FORESTS', 'NN', 'TIME_SERIES', 'UNSUPERVISED_MODELS',
        'SEMISUPERVISED_MODELS', 'GRAPHICAL_MODELS']
    write_section(A)

    f.write("## Social science:\n\n")
    A = [ 'SOCSCI_EXP', 'CAUSAL', 'MATCHING', 'INSTRUMENTAL', 'REGRESSION_DISCONTINUITY', 'NATURAL_EXPTS']
    write_section(A)

    f.write("## Data handling:\n\n")
    A = ['DATA_EXP', 'TEXT_FILES', 'RELATIONAL_DBS', 'NLP', 'GRAPH_DB', 'MULTIMEDIA', 'SENSORS', 'BIG_DATA', 'GEOSPATIAL']
    write_section(A)

    f.write("## Interests, teamwork and motivation:\n\n")
    A = ['PROJECTS', 'GITHUB', 'MOTIVATION', 'TEAMWORK_EXP', 'FUTURE', 'ANYTHING_ELSE']
    write_section(A)

    f.close()

    # Use pandoc to convert .md file to .pdf
    pdf_filename = path_to_pdf_output+str(id+1)+'.pdf'
    #cmd =  "/opt/homebrew/bin/pandoc "+md_filename +" --variable mainfont=\"Helvetica\" --pdf-engine=xelatex -o "+pdf_filename
    #res = os.system(cmd)
    cmd =  ["/opt/homebrew/bin/pandoc", md_filename,  "--variable", "mainfont=\"Helvetica\"", "--pdf-engine=xelatex" ,"-o", pdf_filename]
    p = subprocess.Popen(cmd, stderr=open(os.devnull, 'wb'))
    output, err = p.communicate()
    res = p.wait()


    if res != 0:
        print("Failed to create "+pdf_filename)
        failures.append(id + 1)

print("Conversion failed for ", failures)
