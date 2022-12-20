
import pandas as pd
import os
import subprocess
from tqdm import tqdm

path_to_data = '/Users/colmconnaughton/Warwick/DSSG/msforms/'
filename = 'partner_EOIs.xlsx'

data = pd.read_excel(path_to_data+filename, index_col=0)

titles = list(data.columns)

tags = ['START_TIME', 'COMPLETE_TIME', 'EMAIL', 'NAME', 'FORENAME', 'SURNAME',
        'EMAIL2', 'ADDRESS', 'ROLE', 'ORGANISATION', 'WWW', 'PHONE', 'GEOG_SCOPE',
        'ABOUT','PROJECTS', 'INTERNAL', 'DATA','ANYTHING_ELSE', 'AGREE_TCS'
         ]

fields = {tags[k]:item for k,item in enumerate(titles[0:len(tags)])}

shortnames = [item.strip().replace(' ','_') for item in list(data[fields['ORGANISATION']])]

path_to_md_output = '/Users/colmconnaughton/Warwick/DSSG/msforms/projects-md/'
path_to_pdf_output = '/Users/colmconnaughton/Warwick/DSSG/msforms/projects-pdfs/'

def write_section(sectags):
  for k in sectags:
        f.write("**"+fields[k]+"** : ")
        f.write(str(data.iloc[id][fields[k]]))
        f.write('\n\n')

failures = []
for id in tqdm(range(len(data))):
    md_filename = path_to_md_output+str(id+1)+'-'+shortnames[id]+'.md'
    f = open(md_filename, "w")
    f.write('### '+ str(data.iloc[id][fields['ORGANISATION']])+'\n\n')
    A = ['FORENAME', 'SURNAME',
        'EMAIL2', 'ADDRESS', 'ROLE', 'ORGANISATION', 'WWW', 'PHONE', 'GEOG_SCOPE',
        'ABOUT','PROJECTS', 'INTERNAL', 'DATA','ANYTHING_ELSE']
    write_section(A)
    f.close()

    # Use pandoc to convert .md file to .pdf
    pdf_filename = path_to_pdf_output+str(id+1)+'-'+shortnames[id]+'.pdf'
    cmd =  ["/opt/homebrew/bin/pandoc", md_filename,  "--variable", "mainfont=\"Helvetica\"", "--pdf-engine=xelatex" ,"-o", pdf_filename]
    p = subprocess.Popen(cmd, stderr=open(os.devnull, 'wb'))
    output, err = p.communicate()
    res = p.wait()


    if res != 0:
        print("Failed to create "+pdf_filename)
        failures.append(id + 1)

print("Conversion failed for ", failures)
