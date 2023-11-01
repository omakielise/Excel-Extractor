# Excel-Extractor

The Excel Extractor is a customized natural-text information extraction tool created in Microsoft Excel to extract information on 12 attributes of interest from narratives recorded in the Violent Death Reporting System. While the tool was designed to extract key information from natural text narratives describing fatal drug overdoses, its approach can be generalized to other domains. To guarantee ease-of-use and platform compatibility, the Excel Extractor was created exclusively with native Excel functions (such as MID, FIND, and SUBSTITUTE). It has no macros, VBA scripts, or connections to external programs, i.e., software security concerns are allayed as there is no custom imperative programming language code, and it will work on any offline machine, as long as the machine has base Microsoft Excel installed.

The tool is packaged into seven different files to prevent freezing or crashing; the layout and procedure to run each of the main files is exactly the same. The seven files and attributes programmed are:

EE1:
* Where was the V found?
* Who found the V?
* Who last saw the V alive?
* Was the V experiencing homelessness at time of death?

EE2:
* Did the V have a history of substance use?
* What substances?
    * Alcohol
    * Heroin
    * Crack/Cocaine
* Did the V have a previous overdose?
    
EE3: Did the V have any acute or chronic medical conditions?
* Cardiovasular Disease
* Endocrine Disease
* Gastrointestinal Disease
* Infectious Disease
* Mental Health Disorder
* Neurologic Disease
* Obesity
* Respiratory Disease
* Substance Use Disorder

EE4: Mental Health Disorder
* ADHD
* Anxiety
* Bipolar Disorder
* Depression
* PTSD
* Schizophrenia
* Suicidal ideation or suicide attempt
* Other

EE5: Did the V have any contact with the healthcare system? 
* Recieving care for cancer
* Recieving care for cardiovascular disease
* Has a current prescription medication
* Recieving care for endocrine disease
* Recieving care for gastrointestinal disease
* Recieving care for infectious disease
* Recieving care for mental health
* Recieving care for musculoskeletal disease
* Recieving care for neurologic disease
* Recieving care for respiratory disease
* Recieving care for substance use, including rehab
* Other (i.e. recent emercengy or trauma care)

EE6: Did the V experience any stressful life events?
* Related to criminal or legal proceedings
* Related to employment
* Related to a family member (i.e. death in the family)
* Related to a significant other (i.e. recent divorce)
* Related to finances
* Related to housing
* Related to health
* Other

EE7: What points prior to the overdose offered potential opportunities for intervention?
* During an encounter with a (specific) healthcare provider (excluding rehab)
* Following contact with the healthcare system at large
* During or upon completing of rehab
* During legal court proceedings
* Upon release from prison

The Excel Extractor follows three steps to extract data. First, an indicator dictionary is used to identify the narratives’ relevant sections. Next, a keyword dictionary is used to extract the desired information from the relevant sections. Lastly, extracted information is then mapped to label values. To achive these steps, each attribute has six associated worksheets:

1. Indicator and keyword term dictionary
2. CME indicator array
3. LE indicator array
4. CME keyword array
5. LE keyword array
6. Extracted data

The user must paste the narratives they want processed into the empty cells on the 'main' spreadsheet under the columns labeled ‘CME’ and ‘LE’ (medical examiner and law enforcement narratives respectively) in the selected main.xlsx file. Each narrative is processed independently, resulting in two labels for each entry. Because the medical examiner’s typically has more detail regarding the information of interest, labels resulting from that narrative were given priority such that the law enforcement label will only be used in the case that the medical examiner narrative resulted in no keyword matches. Once this is done, the narratives will be automatically processed and the columns corresponding to the extracted information populated. 

If the user wants to modify a particular dictionary or the way the Excel Extractor works for a specific question, they need to open the corresponding main.xlsx file, and then unhide the question-specific spreadsheets. Once unhidden, the user will be able to add/remove terms to the dictionary, as well as modify the way the information is extracted.
