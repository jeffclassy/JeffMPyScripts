#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
#192204204200203146135135202185207134191193204192205186205203189202187199198204189198204134187199197135194189190190187196185203203209135162189190190165168209171187202193200204203135164193187189198203189188133171187202193200204203135162189190190165183164193187189198203189188125138136155196185203192125138136164199191193187183174137134137134200209
import sys
import clr

clr.AddReference('RevitAPIUI')
from  Autodesk.Revit.UI import TaskDialog, TaskDialogIcon, TaskDialogCommonButtons

clr.AddReference("Microsoft.Office.Interop.Outlook")
from System.Runtime.InteropServices import Marshal

#get script
import os
clr.AddReference('DynamoRevitDS')
import Dynamo 
scriptname=(Dynamo.Applications.DynamoRevit()).RevitDynamoModel.CurrentWorkspace.FileName.ToString()
name=os.path.basename(scriptname)
filename,extention=os.path.splitext(name)
#

task_dialog_err = TaskDialog("wxbsystems")
task_dialog_err.FooterText = "wxbsystems@gmail.com"
task_dialog_err.MainIcon = TaskDialogIcon.TaskDialogIconError
task_dialog_err.TitleAutoPrefix = False
task_dialog_err.MainInstruction = "Outdated Script"
task_dialog_err.MainContent = "Please upgrade to the latest version 4.3"
task_dialog_err.CommonButtons = TaskDialogCommonButtons.Ok
task_dialog_err.Show()

try: # outlook send BIM Guru, www.bimguru.com.au
	mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
	mail.Recipients.Add("jeffm.revit.python.scripts@gmail.com")
	#mail.Recipients.Add("wxbsystems@gmail.com")
	mail.Subject = "CBD Opening Script - " + filename
	mail.Body = "Outdated script Report."
	mail.Send();
	wasSent = True	
except:
	wasSent = False

OUT = 0
