from pyresparser import ResumeParser
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from docx import Document
from geotext import GeoText
from datetime import date

import os
import pandas as pd
import numpy as np
import spacy
import pyodbc 
import re
import nltk
import io
import string
import json
from os import path 

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

#Connecting with the db
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                    "Server=192.168.6.15;"
                    "Database=4DHRIS;"
                    "uid=sa;"
                    "pwd=arch_sql@dxb15;"
                    "Trusted_Connection=no;")

#Fn for extracting text from .docx
def convertDocxToText(path):
    document = Document(path)
    return "\n".join([para.text for para in document.paragraphs])

#Fn for extracting text from .pdf extention
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as fh:
        # iterate over all pages of PDF document
        for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
            # creating a resoure manager
            resource_manager = PDFResourceManager()
            
            # create a file handle
            fake_file_handle = io.StringIO()
            
            # creating a text converter object
            converter = TextConverter(
                                resource_manager, 
                                fake_file_handle, 
                                codec='utf-8', 
                                laparams=LAParams()
                        )

            # creating a page interpreter
            page_interpreter = PDFPageInterpreter(
                                resource_manager, 
                                converter
                            )

            # process current page
            page_interpreter.process_page(page)
            
            # extract text
            text = fake_file_handle.getvalue()
            yield text

            # close open handles
            converter.close()
            fake_file_handle.close()

#Extracting the File Type
def extractText (file_Path):
    text = ""
    #Extracting Distict skills from Db
    df = pd.read_sql_query('select DISTINCT SkillName from dbo.CV_skills', cnxn)
    filename, file_extension = os.path.splitext(file_Path)
    if file_extension == '.pdf':
        for page in extract_text_from_pdf(file_Path):
            text += ' ' + page
    elif file_extension == '.docx':
        text = convertDocxToText(file_Path)

    #Extracting the fields using python library - Method 1
    data = ResumeParser(file_Path).get_extracted_data()
    print("*******************************")
    print(data)
    print("*******************************")

    return text,data,df

#Extracting the skills field
def extract_skills(resume_text,df):
    # load pre-trained model
    nlp = spacy.load('en_core_web_sm')
    nlp_text = nlp(resume_text)

    # removing stop words and implementing word tokenization
    tokens = [token.text for token in nlp_text if not token.is_stop]
    
    # extract values
    skills = list(df.SkillName.values)
    skills = [x.lower() for x in skills]
    
    skillset = []
    
    # check for one-grams
    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)
    
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in skills:
            skillset.append(token)
    
    return [i.capitalize() for i in set([i.lower() for i in skillset])]


#Fn for extracting Educational fields from text
def extract_edu(resume_text):
    # Extracting Distinct education from Db
    edu = pd.read_sql_query('select DISTINCT EducationType, SchoolName, DegreeType,DegreeName, DegreeYear, DegreeMajor, NormalizedGPA  from dbo.CV_Education', cnxn)

    # load pre-trained model
    nlp = spacy.load('en_core_web_sm')
    nlp_text = nlp(resume_text)

    # removing stop words and implementing word tokenization
    tokens = [token.text for token in nlp_text if not token.is_stop]
    
    # extract EducationType values
    eduType = list(edu.EducationType.unique())
    eduType = [x.lower() for x in eduType]
    
    eduTypeList = []
    
    # check for one-grams
    for token in tokens:
        if token.lower() in eduType:
            eduTypeList.append(token)
    
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in eduType:
            eduTypeList.append(token)
    
    college = ['college', 'community', 'professional', 'university', 'university']
    EducationType = ""
    for i in set([i.lower() for i in eduTypeList]):
        if i in college:
            EducationType = i.capitalize()
            print("EducationType: ",EducationType)
            break
            
    
    # extract SchoolName values
    schoolName = list(edu.SchoolName.unique())
    schoolName = [x.lower() for x in schoolName]
    
    schoolNameList = []    
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in schoolName:
            schoolNameList.append(token)
    
    SchoolName = ''
    for i in set(np.unique([i.lower() for i in schoolNameList])):
        SchoolName = SchoolName + i.capitalize() + " "        
    print("SchoolName: ",SchoolName)
    
    # extract DegreeType values
    degreeType = list(edu.DegreeType.unique())
    degreeType = [x for x in degreeType]
    
    degreeTypeList = []
    # check for one-grams
    for token in tokens:
        if token.lower() in degreeType:
            degreeTypeList.append(token)
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in degreeType:
            degreeTypeList.append(token)
            
    DegreeType = ''
    for i in set(np.unique([i.lower() for i in degreeTypeList])):
        DegreeType = DegreeType + i.capitalize() + " "        
    print("DegreeType: ",DegreeType)
    
    # extract DegreeName values
    degreeName = list(edu.DegreeName.unique())
    degreeName = [x for x in degreeName]
    
    degreeNameList = []
    # check for one-grams
    for token in tokens:
        if token.lower() in degreeName:
            degreeNameList.append(token)
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in degreeName:
            degreeNameList.append(token)
            
    DegreeName = ''
    for i in set(np.unique([i.lower() for i in degreeNameList])):
        DegreeName = DegreeName + i.capitalize() + ","
    print("DegreeName: ",DegreeName)
    

    # extract NormalizedGPA values
    normalizedGPA = list(edu.NormalizedGPA.unique())
    normalizedGPA = [x for x in normalizedGPA]
    
    normalizedGPAList = []
    # check for one-grams
    for token in tokens:
        if token.lower() in normalizedGPA:
            normalizedGPAList.append(token)
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in normalizedGPA:
            normalizedGPAList.append(token)

    print("NormalizedGPA: ",normalizedGPAList)
    return EducationType, SchoolName, DegreeType, DegreeName, normalizedGPAList


#Function to extract the education year
def get_degree_year(text):
    degreeYearExtract = ""
    # load pre-trained model
    nlp = spacy.load('en_core_web_sm')
    nlp_text = nlp(text)

    # removing stop words and implementing word tokenization
    tokens = [token.text for token in nlp_text if not token.is_stop]
    
    # check for bi-grams and tri-grams
    for token in nlp_text.noun_chunks:
        token = token.text.lower()
        degreeYearSearch = re.search("[1-3][0-9]{3} - [1-3][0-9]{3}", token)
        if degreeYearSearch :
            degreeYearExtract = degreeYearSearch.group()
            break
    return degreeYearExtract

def employmentExperience(data):
    # Extracting the experience details from text
    exp = data['experience']
    ExecutiveSummary = data['experience']
    print("ExecutiveSummary : ", ExecutiveSummary)
    if exp:
        exp_text = ""
        for ele in exp: 
            exp_text += ele 
            places = GeoText(ele)
            if places.cities:
                OrgLocation = places.cities[0]
            else:
                OrgLocation = ""
            
        exp_year = re.search("[A-Za-z]{4} [1-3][0-9]{3} to [pP]resent" or "[1-3][0-9]{3}" or 
                            "[A-Za-z]{3} [1-3][0-9]{3}" or "[A-Za-z]{3} [1-3][0-9]{3} - [pP]resent" or 
                            "[1-3][0-9]{3} - [1-3][0-9]{3}" or "[A-Za-z]{3}. [1-3][0-9]{3}" or 
                            "[A-Za-z]{3}. [1-3][0-9]{3} - [A-Za-z]{3}. [1-3][0-9]{3}" or
                            "[A-Za-z]{4} [1-3][0-9]{3} - [tT]ill [dD]ate" or
                            "[A-Za-z]{3}. [1-3][0-9]{3} - [tT]ill [dD]ate", exp_text)
        if exp_year:
            Dates = exp_year.group()
        else:
            Dates = ''
    else:
        Dates = ''
        OrgLocation = ""
        
    print("OrgLocation :",OrgLocation)   
    print("Date: ",Dates)
    if Dates:
        empDates = Dates.split(' ')
        empDates = [i.lower() for i in empDates]
        if 'present' in empDates:
            expEndDate = str(date.today())
            for i in empDates:
                startYear = re.search("[1-3][0-9]{3}",i)
                if startYear:
                    startYear = startYear.group()
                    break
            for i in empDates:
                startMonth=re.search("[Jj][uU][nN]" or "[Jj][aA][nN]" or "[Ff][eE][Bb]"
                                or "[Mm][aA][rR]" or "[Aa][pP][rR]" or "[mM][aA][yY]"
                                or "[jJ][uU][lL]" or "[aA][uU][gG]" or "[sS][eE][pP]"
                                or "[oO][cC][tT]" or "[Nn][oO][vV]" or "[dD][cC][mM]",i)
                if startMonth:
                    startMonth = startMonth.group()
                    break
                else:
                    startMonth = "jan"
            monthDict = {"jan" : "01",
                        "feb" : "02",
                        "mar" : "03",
                        "apr" : "04",
                        "may" : "05",
                        "jun" : "06",
                        "jul" : "07",
                        "aug" : "08",
                        "sep" : "09",
                        "oct" : "10",
                        "nov" : "11",
                        "dec" : "12"}
            for key in monthDict:
                month_key = re.match(key,startMonth)
                if month_key:
                    month_key = month_key.group()
                    break
            startMonth = monthDict[month_key]
        expStartDate = str(startYear + "-" + startMonth + "- " + "01")

    else:
        expEndDate = ""
        expStartDate = ""

    print("End Date:",expEndDate)
    print("Start Date:",expStartDate)
    return ExecutiveSummary, OrgLocation, expEndDate, expStartDate

#Creating a dict including all extracted values
def extractResume(file_Path):
    
    #Fn call for text extraction
    text, data, df = extractText (file_Path)
    PersonName = data["name"]
    Mobile = data["mobile_number"]
    EmailAddress = data["email"]
    PositionName = data["designation"]
    OrganizationName = data["company_names"]
    if data["total_experience"] == 0:
        AvgTotalExp = 0
    else:
        AvgTotalExp = data ["total_experience"] * 12

    #Fn call for Skill extraction
    skills_extract = pd.DataFrame()
    skills_data = pd.DataFrame()
    skills_extract["skills"] = extract_skills(text,df)
    skills_data["skills"] = data["skills"]
    SkillName = ((skills_extract["skills"] + skills_data["skills"]).unique())
    SkillName = SkillName.tolist()


    #Fn call for education details
    EducationType, SchoolName, DegreeType, DegreeName, normalizedGPAList = extract_edu(text)


    #Fn call for getting degree year details
    Degree_Year = get_degree_year(text)
    degreeDates = ""
    DegreeYear = ""
    degreeStartDate = ""
    degreeEndDate = ""
    if Degree_Year: 
        degreeDates = Degree_Year.split('-')
        degreeDates[0] = degreeDates[0].replace(" ","")
        degreeDates[1] = degreeDates[1].replace(" ","")
        degreeStartDate = degreeDates[0] + "-01-01"
        degreeEndDate = degreeDates[1] + "-01-01"
        DegreeYear = degreeDates[1]    
    print("StartDate :", degreeStartDate)
    print("EndDate :", degreeEndDate)
    print("Degree Year :", DegreeYear)


    #Fn call for employement experience
    ExecutiveSummary, OrgLocation, expEndDate, expStartDate = employmentExperience(data)

    cv_data = {
        "BasicInfo":{
            "PersonName" : PersonName,
            "Mobile" : Mobile,
            "EmailAddress" : EmailAddress
        },
        "EmploymentHistory":{
            "PositionName" : PositionName,
            "OrganizationName" : OrganizationName,
            "AvgTotalExp" : AvgTotalExp,
            "SkillName" : SkillName,
            "ExecutiveSummary" : ExecutiveSummary,
            "OrgLocation" : OrgLocation,
            "expEndDate" : expEndDate,
            "expStartDate" : expStartDate
        },
        "EducationHistory":{
            "EducationType" : EducationType,
            "SchoolName" : SchoolName,
            "DegreeType" : DegreeType,
            "DegreeName" : DegreeName,
            "normalizedGPAList" : normalizedGPAList,
            "Degree_Year" : Degree_Year,
            "degreeStartDate" : degreeStartDate,
            "degreeEndDate" : degreeEndDate
            }  
    }

    print(cv_data)
    return cv_data