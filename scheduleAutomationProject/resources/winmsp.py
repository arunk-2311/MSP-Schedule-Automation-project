import win32com.client as win32
import pandas as pd

file= win32.Dispatch('MSProject.Application')

file.Visible =1

file.FileOpen('C:/Project2.mpp')

project=file.ActiveProject


'''
This function will return a pandas dataframe containing the task from an the supplied 
win32com referenced microsoft project "project".

Pandas dataframes are great for manipulating table of data and can readily output excel spreadsheets
which is our ultimate aim.


'''


def create_project_data_frame(project, UniqueIDs_to_Ignore, mspApplication, headers):
    '''
    Keywords:  
        * project = reference to the win32com microsoft project object.
        * UniqueIDs_to_Ignore = this is a list of MSProject unique IDs of tasks we wish to exclude from the 
        dataframe.  This feature is added as occassionally I come across projects where the PM
        doesnt want to remove a task even though it is no longer valid, prefereing to keep it
        in for future reference.
        * mspApplication = reference to the win32com MSProject applications being used.
        * headers = list of headers to output in the pandas dataframe
    '''
    
    # Create reference to empty Pandas dataframe with headers
    projectDataFrame = pd.DataFrame(columns=headers)
    
    # reference to initially empty list to be used to collect summary tasks
    summary_tasks_to_task = []
    
    # reference to collection of Tasks in MS Project using the Project "Tasks" collection object
    task_collection= project.Tasks

    
    # iterate through the collection of tasks
    for t in task_collection:
        # determine is the task is not a summary task and is not a task to be ignored
        if (not t.Summary) & ~(t.UniqueID in UniqueIDs_to_Ignore):  # i.e. it is a task line not a Summary Task
            # find dependent task
            dep = []  # an empty list to add dependent task id
            for d in t.TaskDependencies:
                if int(d.From) != t.UniqueID:  # a task can have multiple references to itself, not sure why, but this removes them
                    #print(type(d))
                    dep.append(str(d.From) + "-" + str(d.From.Name))

            # collect resource names        
            res = []  # an empty list to add resources
            for r in t.Assignments:
                print(type(r))
                res.append(r.ResourceName)    
            
            ''' 
            it is not good practic but it is possible to have project tasks at the top level (outline level 1)
            So this if statement catches those occurances and empties summary_tasks_to_task list 
            '''
            if t.OutlineLevel ==1:
                summary_tasks_to_task = []
            # create a temporary "temp" list variable holding the entries for the dataframe row.
            sum_task = ">".join(summary_tasks_to_task)
            dependencies = [", ".join(dep)]
            resources = [", ".join(res)]

            temp = [t.UniqueID, sum_task]
            
            # iterate over the list of headers (excluding UniqueID and SummaryTask as these are covered)
            # add the value of the header to the list
            for head_title in headers:
                if head_title != "UniqueID" and head_title != "SummaryTask":
                    '''
                    note that dependencies and resources have been created by iterating over their
                    respective collection objects and are therefore not found via Task.GetField
                    '''
                    if head_title == "Predecessors":
                        temp = temp+dependencies
                    elif head_title=="Resource Names":
                        temp = temp+resources
                    
                    #Other headers can, in the main, be found by access the field value of the header
                    #So it is important that the name of header is correct.  We can make a function
                    #to check whether the headers list contains valid headers
                        
                    else:
                        temp = temp+[t.GetField(mspApplication.FieldNameToFieldConstant(head_title))]

            # Append the task to pandas dataframe            
            projectDataFrame = projectDataFrame.append(pd.Series(temp, index=headers), ignore_index=True)

        elif t.Summary & (t.OutlineLevel > len(summary_tasks_to_task)):
            '''
            if tasks is a summary task and its outline level is greater than number of summary tasks in the list
            that is it is lower level summary task
            then add that summary task to the list summary_tasks_to_task
            '''
            summary_tasks_to_task.append(t.Name)

        else:
            while not len(summary_tasks_to_task) == t.OutlineLevel - 1:
                '''
                if tasks is a summary task and its outline level is less than number of summary tasks in the list
                summaryTasksToTask then remove last summary task from list and add new summary task to the list
                (Basically we have gone up a summary task level)
                '''
                summary_tasks_to_task.pop()
            summary_tasks_to_task.append(t.Name)

    # finally, set the index of the dataframe to the unique MS Project Task ID
    projectDataFrame = projectDataFrame.set_index("UniqueID")
    
    # set the type of the Finish and Start columns to datatime types    
    projectDataFrame["Finish"] = pd.to_datetime(projectDataFrame["Finish"],dayfirst=True)
    projectDataFrame["Start"] = pd.to_datetime(projectDataFrame["Start"],dayfirst=True)
    
    #return projectDataFrame

#task_collection= project.Tasks
    for t in task_collection:
        t.SetField(mspApplication.FieldNameToFieldConstant("newc"),"hi")
        print(t.GetField(mspApplication.FieldNameToFieldConstant("newc")))
        #print(type(t.GetField(mspApplication.FieldNameToFieldConstant("Start"))))
        
        #t.SetField(mspApplication.FieldNameToFieldConstant("newc"))="hi"
   
    return projectDataFrame

const_header = ["UniqueID", "SummaryTask", "Name", "Start", "Finish", "% Complete"]
additional_header=["Resource Names", "Notes", "Predecessors", "Text1","newc"]
headers=const_header+additional_header
ignoreID = [44]
frame = create_project_data_frame(project, ignoreID, file, headers)

#print(frame)