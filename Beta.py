from jira.client import JIRA
import simplejson as json
import win32com.client
import xlrd
from pathlib import Path
#Server
JIRA_SERVER={'server':#Your Server
}

#Credentials
USER=#Your Username
PW=#Your Password

#JIRA Custom variables
TC_COMPLEXITY='customfield_2007036'
PRIORITY='customfield_2006121'
EPIC_NAME='customfield_2003424'
#Max Results to be Shown
MAX_RESULTS=1

#File Destination

file=Path(
    #Your Excel Sheet to Read/Write in/From
).resolve()

#Use Excel
excel = win32com.client.Dispatch("Excel.Application")
#Open Workbook (EXCEL)
wb = excel.Workbooks.Open(file)
ws=wb.Worksheets(#Sheet name will retreive/Write in
)
config=wb.Worksheets(#Sheet name where you can put Custom JQls
)

#Query
JQL= config.Range("F4").Value
CUSTOM_JQL=config.Range("F5").Value
#Try to Connect
try:
    jira=JIRA(options=JIRA_SERVER,basic_auth=(USER,PW))
    print('Login Successful!')
except:
    print('Login Failed!')

def create_TestCase():
    wb = xlrd.open_workbook(str(file),"r")
    
    ws = wb.sheet_by_name(#Sheet name will retreive/Write in
    )
    #Loop to create the TestCases on JIRA
    for rownum in range(4,ws.nrows):
        row_values = ws.row_values(rownum)  #Put Cell value in var
        sumary=row_values[4]
        comp=row_values[5]
        pri=int(row_values[6])              #Parse PRIORITY to Int
        epic=row_values[7]
        descr=row_values[8]
        status=row_values[8]
        labelsList=row_values[10].split(',') #Convert the Labels to Array
        rep=str(row_values[11])
        assign=str(row_values[12])
        
        #Assign ID for PRIORITY based on cell Value    
        if pri == 1:
            priID='2008812'
        elif pri==2:
            priID='2008813'
        else:
            priID='2008814'    

        if  comp=="3 - Medium":
            compID='2010282'
        elif comp=="4 - High":
            compID='2010283'
        elif comp=="2 - Low":
            compID='2010281'
        elif comp=="1 - Very low":
            compID='2010280'
        elif comp=="None":
            compID='null'
        
        
        if status=="In Testing":
            statusID=51
        elif status=="Test Blocked":
            statusID=61
        elif status=="Planned":
            statusID=81
        elif status=="Failed":
            statusID=91
        elif status=="Passed":
            statusID=101
        elif status=="Test Case Defined":
            statusID=111
        elif status=="In Definition":
            statusID=121
        elif status=="Result in Verification":
            statusID=131
        elif status=="Closed":
            statusID=141
        else:
            statusID=""

        #Dictionary to be Uploaded
        if comp=="None":
            issue_dict={
                'project':{'id':2007066},
                'summary':sumary,
                'customfield_2006121':{'value':pri,'id':priID},
                #'customfield_2003424':{'name':epic},
                'description':descr,
                #'transition': {'id':81},
                'labels':labelsList,
                #'reporter':{'displayName':rep,'key':''},
                #'assignee':{'displayName':assign,'key':''},
                'issuetype':{'name':'Test Case'}
            }
        else:
            issue_dict={
                'project':{'id':2007066},
                'summary':sumary,
                'customfield_2007036':{'value':comp,'id':compID},
                'customfield_2006121':{'value':pri,'id':priID},
                #'customfield_2003424':{'name':epic},
                'description':descr,
                #'transition': {'id':81},
                'labels':labelsList,
                #'reporter':{'displayName':rep,'key':''},
                #'assignee':{'displayName':assign,'key':''},
                'issuetype':{'name':'Test Case'}
            }
        
        #Create TestCase
        jira.create_issue(fields=issue_dict)

    print("Finished Creating")


def retreive_TestCases():
    #Use Excel
    excel = win32com.client.Dispatch('Excel.Application')

    #Open Workbook (EXCEL)
    wb = excel.Workbooks.Open(file)
    ws=wb.Worksheets('Retreive')
    getData=jira.search_issues(JQL,maxResults=MAX_RESULTS)
    
    for idx,issue in enumerate(getData):
        idxL=0
        rownum=idx+5
        
        key=issue.key
        summary=issue.fields.summary
        complexity=str(issue.fields.customfield_2007036)
        priority=str(issue.fields.customfield_2006121)
        epicLink=str(issue.fields.customfield_2003423)

        if epicLink!="None":
            epicName= get_epicName(epicLink)
        else:
            epicName="None"

        status=str(issue.fields.status)
        labels= issue.fields.labels
        reporter=issue.fields.reporter.displayName
        assignee=str(issue.fields.assignee)
        if assignee!="None":
            assignee=issue.fields.assignee.displayName
        else:
            assignee="Unassigned"
        description=issue.fields.description


        ws.Range("D%d" %rownum).Value=key                                 
        ws.Range("E%d" %rownum).Value=summary                      
        ws.Range("F%d" %rownum).Value=complexity                                  
        ws.Range("G%d" %rownum).Value=priority
        ws.Range("H%d" %rownum).Value=epicName
        ws.Range("I%d" %rownum).Value= description
        ws.Range("J%d" %rownum).Value= status
        
        while idxL < len(labels):
            StrLabels=','.join(labels)
            ws.Range("K%d" %rownum).Value=StrLabels                      
            idxL += 1
        
        ws.Range("L%d" %rownum).Value= reporter
        ws.Range("M%d" %rownum).Value= assignee                      
        
    wb.Save()

    print("Finished Retreiving!")   


def get_epicName(epicID):
    epic=jira.issue(epicID)
    getEpicName=epic.fields.summary
    return getEpicName


def delete_Rows():
    ws.Range("D5:M1000").EntireRow.Delete()



if __name__=='__main__':
    delete_Rows()
    retreive_TestCases()
    #copyToUploadSheet()
    create_TestCase()
    delete_Rows()
    retreive_TestCases()
