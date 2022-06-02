import win32com.client as win32

projectFile = win32.Dispatch('MSProject.Application')
projectFile.FileOpen('C:/Users/iamto/Documents/scheduleAutomationProject/input.mpp')
projectFile.Visible = 1