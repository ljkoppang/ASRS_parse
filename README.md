#ASRS Parse
##Takes pdf test result file, extracts summary information, and creates 2 .docx files containing 2 tables and a summary for use ASD assessment.

Psychologists who assess children for Autism Spectrum Disorder, run several different standard assessments, then summarize them in a final report.
This project takes the results from the ASRS (Autism Spectrum Rating Scales) in pdf form, extracts summary information, and creates 2 tables and 
a summary paragraph in docx format to be added to the final report, saving time in the final report creation process.

pip insstall:
#### %pip install pypdf
#### %pip install pdfplumber
#### %pip install pandas
#### %pip install python-docx

#### %pip install pyinflect spacy
#### %python -m spacy download en_core_web_sm

#### %pip install pdfminer.six==20221105
#### %pip install pdfplumber==0.9.0

To do List:
* Format docx tables with Times New Roman size 12 font
* Add padding to table top and bottom
* Add shading to the title and total rows of the appendix table.
* To the add_comment function, add code to confirm entire comment is in present tense.
