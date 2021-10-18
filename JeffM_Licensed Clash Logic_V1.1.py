#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
#192204204200203146135135202185207134191193204192205186205203189202187199198204189198204134187199197135194189190190187196185203203209135162189190190165168209171187202193200204203135164193187189198203189188133171187202193200204203135162189190190165183164193187189198203189188125138136155196185203192125138136164199191193187183174137134137134200209
import sys
#sys.path.append(C....\bimorphNodes\bin')
#import os as os
import socket
import getpass
import math
host = socket.gethostname().upper()
hostname = getpass.getuser().lower()
import clr
# outlook send BIM Guru, www.bimguru.com.au
clr.AddReference("Microsoft.Office.Interop.Outlook")
from System.Runtime.InteropServices import Marshal
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
from Autodesk.DesignScript.Geometry import Curve as GeomCurves
from Autodesk.DesignScript.Geometry import Solid as Sld
from Autodesk.DesignScript.Geometry import Point as Pnt
from Autodesk.DesignScript.Geometry import Line as Ln
from Autodesk.DesignScript.Geometry import Surface as Srfc, PolyCurve as PlyCrv
#get script
import os
clr.AddReference('DynamoRevitDS')
import Dynamo 
scriptname=(Dynamo.Applications.DynamoRevit()).RevitDynamoModel.CurrentWorkspace.FileName.ToString()
name=os.path.basename(scriptname)
filename,extention=os.path.splitext(name)
#

clr.AddReference('System')
from System.Collections.Generic import List

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from  Autodesk.Revit.UI import TaskDialog 

clr.AddReference('DSCoreNodes') 
from DSCore.Web import WebRequestByUrl as Wb
from DSCore import String as strng
from DSCore import List as lst
#from DSCore import Revit.Elements as Rvt

clr.AddReference('BimorphNodes')
from Revit import Element as Bimorph#, LinkElement
from Revit import LinkElement as LinkedBimorph
import Curve as BMCurve

# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Start Transaction
doc = DocumentManager.Instance.CurrentDBDocument
TransactionManager.Instance.EnsureInTransaction(doc)
try: # outlook send BIM Guru, www.bimguru.com.au
	mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
	mail.Recipients.Add("jeffm.revit.python.scripts@gmail.com")
	mail.Subject = "CBD Opening Script - " + filename
	mail.Body = host + "/" + hostname + " run this code. " + str(0) + " clashes found. License not yet paid."
	mail.Send();
	wasSent = True	
except:
	wasSent = False

TaskDialog.Show('License','Your trial license has expired.')

TransactionManager.Instance.TransactionTaskDone()

OUT = 0
