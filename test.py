import xlsxwriter
import urllib.request
from stackapi import StackAPI
from random import randint
from bs4 import BeautifulSoup

tags = [
    "SAP",
    "JavaScript",
    "Oracle",
    "Network",
    "HDMI"
    ]

cis = [
    "SAP Materials Management",
    "WEBSERVER",
    "(empty)",
    "nyc rac nas200",
    "*BETH-IBM"
    ]

categories = [
    "Software",
    "Inquiry / Help",
    "Database",
    "Network",
    "Hardware"
    ]

resolution_code = [
    "Solved (Work Around)",
    "Solved (Permanently)",
    "Solved Remotely (Work Around)",
    "Solved Remotely (Permanently)",
    "Not Solved (Not Reproducible)",
    "Not Solved (Too Costly)",
    "Closed/Resolved by Caller"
]

def get_answer(url):
    content = urllib.request.urlopen(url)
    soup = BeautifulSoup(content)
    answer = soup.find_all('div',attrs={'class':'accepted-answer'})
    if len(answer)>1:
        answer = soup.find_all('div',attrs={'class':'accepted-answer'})
        accepted_answer = answer[0].contents[1].find(
                "div",attrs = {'class' :'post-text'}
            )
        return accepted_answer.text
    return 'Not answered'

SITE = StackAPI('stackoverflow')
for i,tag in enumerate(tags):
    workbook = xlsxwriter.Workbook(tag+'.xlsx')
    worksheet = workbook.add_worksheet()
    questions = SITE.fetch('questions',min=20,tagged=tag, sort='votes')

    row = 0 
    for item in questions['items']:
        if(item['is_answered']):
            worksheet.write(row,0,item['title'] )
            worksheet.write(row,1,cis[i] )
            worksheet.write(row,2,categories[i] )
            worksheet.write(row,3,get_answer(item['link']) )
            row += 1


    workbook.close()


