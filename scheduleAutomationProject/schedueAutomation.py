import win32com.client as win32
import pandas as pd
import datetime

projectFile = win32.Dispatch('MSProject.Application')
projectFile.FileOpen('C:/Users/iamto/Documents/scheduleAutomationProject/input.mpp')
projectFile.Visible = 1

rvtInput=['waw','wae','fdn']

project=projectFile.ActiveProject

headers=["Name","Duration","Start","Finish","Predecessors"]
asPlannedSchedule=pd.DataFrame(columns=headers)

taskCollection = project.Tasks

for t in taskCollection:

    temp=[]

    for head in headers:
        temp.append(t.GetField(projectFile.FieldNameToFieldConstant(head)))  

    df=pd.DataFrame([temp],columns=headers)
    asPlannedSchedule = asPlannedSchedule.append(df,ignore_index=True)             


inc=pd.DataFrame(columns=headers)
incIndex=[]


for name in asPlannedSchedule.Name:

    flag=0

    for r in rvtInput:
        if(r==name):
            flag=1

    if flag == 0:
        valMask= asPlannedSchedule.Name == name
        ind=asPlannedSchedule.Name[valMask].index
        incIndex.append(ind[0])


inc=inc.append(asPlannedSchedule.loc[incIndex])
inc["Finish"] = pd.to_datetime(inc["Finish"],dayfirst=True)
inc["Start"] = pd.to_datetime(inc["Start"],dayfirst=True)


def predEffect(pString,inc):
    
    pList=[]
    returnTuple=()
    
    for c in pString:
        if c!=',':
            for i in incIndex:
                if((int(c)-1)==i):
                    #return True,i
                    pList.append(int(c)-1)
                    
    if len(pList)==1:
        returnTuple= True,pList[0]

    elif len(pList)>1:
        lastFinish = inc.Finish[pList[0]]
        fIndex=pList[0]

        for l in pList:
            if lastFinish < inc.Finish[l]:
                fIndex,lastFinish=l,inc.Finish[l]

        returnTuple= True,fIndex          

    else:
        returnTuple = False,-1  

    return returnTuple                             
    

def durationStringToDays(strDate):

    dur=''

    for c in strDate:
        if(c==' '):
            break
        dur+=c

    return (int(dur)-1)        


def tdateUpdate(row,tdate,inc,asPlannedSchedule):
        
    duration=durationStringToDays(inc.Duration[row])

    inc.Start[row]=tdate
    inc.Finish[row]=inc.Start[row] + pd.Timedelta(duration,unit='days')
        
    stringStart=inc.Start[row].strftime("%a %d-%m-%y")
    stringFinish=inc.Finish[row].strftime("%a %d-%m-%y")

    asPlannedSchedule.Start[row]=stringStart
    asPlannedSchedule.Finish[row]=stringFinish


def predUpdateDate(row,inc,asPlannedSchedule):

        duration=durationStringToDays(inc.Duration[row[0]])

        inc.Start[row[0]]=inc.Finish[row[1]] +pd.Timedelta(1,unit='days')
        inc.Finish[row[0]]=inc.Start[row[0]]+ pd.Timedelta(duration,unit='days')
    
        stringStart=inc.Start[row[0]].strftime("%a %d-%m-%y")
        stringFinish=inc.Finish[row[0]].strftime("%a %d-%m-%y")

        asPlannedSchedule.Start[row[0]]=stringStart
        asPlannedSchedule.Finish[row[0]]=stringFinish

    
def updater(tdate_str,inc,asPlannedSchedule):

    tdate=datetime.datetime.strptime(tdate_str,'%Y-%m-%d')

    affBypred=[]

    for i in incIndex:

        if(inc.Finish[i] <= tdate):
            tdateUpdate(i,tdate,inc,asPlannedSchedule)
            flag,prd = predEffect(inc.Predecessors[i],inc)

            if flag:
                affBypred=affBypred +[[i,prd]]
                predUpdateDate([i,prd],inc,asPlannedSchedule)

    #print(affBypred)

#print(asPlannedSchedule)    
updater('2021-07-13',inc,asPlannedSchedule)
#print(inc) 
#print(asPlannedSchedule)

outputFile = win32.Dispatch('MSProject.Application')
outputFile.FileOpen('C:/Users/iamto/Documents/scheduleAutomationProject/output.mpp')
outputFile.Visible = 1

output=outputFile.ActiveProject

#headers=["Name","Duration","Start","Finish","Predecessors"]
#asPlannedSchedule=pd.DataFrame(columns=headers)

outputTaskCollection = output.Tasks

#print(outputTaskCollection(2).Name)
#print(outputTaskCollection.count)

asPlannedSchedule.rename(index=lambda s: s + 1, inplace=True)

for t in range(1,outputTaskCollection.count+1):
    #for i in asPlannedSchedule.index:
        #temp.append(t.GetField(projectFile.FieldNameToFieldConstant(head)))
        #print(t,asPlannedSchedule.Name[t])
        outputTaskCollection(t).SetField(outputFile.FieldNameToFieldConstant("Name"),asPlannedSchedule.Name[t])
        #print(t,outputTaskCollection(t).Name)
        outputTaskCollection(t).SetField(outputFile.FieldNameToFieldConstant("Duration"),asPlannedSchedule.Duration[t])
        outputTaskCollection(t).SetField(outputFile.FieldNameToFieldConstant("Start"),asPlannedSchedule.Start[t])
        outputTaskCollection(t).SetField(outputFile.FieldNameToFieldConstant("Finish"),asPlannedSchedule.Finish[t])
        if (asPlannedSchedule.Predecessors[t] != ''):
            #print('')
            #print(type(asPlannedSchedule.Predecessors[t-1]))    
            outputTaskCollection(t).SetField(outputFile.FieldNameToFieldConstant("Predecessors"),asPlannedSchedule.Predecessors[t])
        #else:
            #print("none")
#for t in range(1,outputTaskCollection.count+1)
#projectField = FieldNameToFieldConstant("TestEntProjText", pjProject)              