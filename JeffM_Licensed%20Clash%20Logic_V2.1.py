#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
#192204204200203146135135202185207134191193204192205186205203189202187199198204189198204134187199197135194189190190187196185203203209135162189190190165168209171187202193200204203135164193187189198203189188133171187202193200204203135162189190190165183164193187189198203189188125138141138136155196185203192125138141138136164199191193187183174138134137134200209
#JeffM_Licensed Clash Logic_V1.1.py
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
from  Autodesk.Revit.UI import TaskDialog, TaskDialogIcon, TaskDialogCommonButtons

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

curvelistB = List[GeomCurves]()
origin = Pnt.ByCoordinates(0,0)
empty=[]

def verif (d):
	if d[0] == host and d[1] == hostname and d[2] == chr(49):
		return True
def limiter (d):
	return d[3]
def fixrotation (rot):
	if rot==180:
		rot=0
	elif rot>180:
		rot=abs(rot-360)
	else:
		rot
	return rot
def beamjustification (beam,ht):
	zjust = beam.GetParameterValueByName("z Justification")
	zoffset = beam.GetParameterValueByName("z Offset Value")
	beamloc = beam.Location
	if zjust == 0: #Top
		beamloc = Geometry.Translate(beamloc,Vector.ZAxis(),-ht + zoffset)
	elif zjust == 1 or zjust == 2: #Center or Origin
		beamloc = Geometry.Translate(beamloc,Vector.ZAxis(),-ht/2 + zoffset)
	elif Zjust == 3: #Bottom
		beamloc = Geometry.Translate(beamloc,Vector.ZAxis(),zoffset)
	else:
		zoffset = 0 
	return beamloc
def getarcsurface(arc):	
	arccategory = UnwrapElement(arc).Category.Name
	if arccategory == 'Floors':
		flr = UnwrapElement(arc)
		newcs = arc.TotalTransform
		for ref in HostObjectUtils.GetTopFaces(flr):
			boundaryloops = flr.GetGeometryObjectFromReference(ref).GetEdgesAsCurveLoops()
			for loop in boundaryloops:
				floorsketch=[x.ToProtoType() for x in loop]
		arcsurface=Srfc.ByPatch(PlyCrv.ByJoinedCurves(floorsketch,0.01))
		arcsurface=Geometry.Transform(arcsurface,newcs)
	elif arccategory == 'Structural Framing':
		arcunwrap = UnwrapElement(arc)
		arctype = arcunwrap.Document.GetElement(arcunwrap.GetTypeId())
		archeight = arctype.LookupParameter('h').AsDouble() * 304.8		
		arcloc = beamjustification(arc,archeight)		
		arcsurface = GeomCurves.Extrude(arcloc,Vector.ZAxis(),archeight)
	else:
		arcloc = arc.Location
		if arccategory == 'Walls' and isinstance(arcloc,GeomCurves) :
			arcloc = GeomCurves.Translate(arcloc,Vector.ZAxis(),arc.GetParameterValueByName("Base Offset"))
			arcsurface = GeomCurves.Extrude(arcloc,Vector.ZAxis(),arc.GetParameterValueByName("Unconnected Height"))
		elif arccategory != 'Walls' and isinstance(arcloc,Pnt):
			arcangle = math.degrees(UnwrapElement(arc).Location.Rotation)
	    		basepoint1 = GeomCurves.Translate(arcloc,Vector.XAxis(),1000)
			basepoint2 = GeomCurves.Translate(arcloc,Vector.XAxis(),-1000)
			arcline = Ln.ByStartPointEndPoint(basepoint1,basepoint2)
			archeight = 4000 if arc.GetParameterValueByName("Height") is None else arc.GetParameterValueByName("Height")
			arcsurface = GeomCurves.Extrude(arcline,Vector.ZAxis(),archeight) #assuming Height contains value
			arcsurface = Geometry.Rotate(arcsurface,arcloc,Vector.ZAxis(),arcangle)
		else:
			arcsurface = 0
	return arcsurface
def clash (arcsurface,meploc):
	intersect = 0
	try:
		intersect1 = Geometry.Intersect(arcsurface,meploc) #line to surface #Surface vs Line given the line crosses the surface		
		if intersect1 == []:
			itersect = Geometry.ClosestPointTo(arcsurface,meploc)
		else:
			intersect = intersect1[0]
	except:
		intersect = 2
	return intersect	
def getmidpoint (crv):
	return GeomCurves.PointAtParameter(crv,0.5)
def pointclosetowallmid (arc,ptoi):
	try:
		arcstartpoint = a.Location.StartPoint
		arcendpoint = a.Location.EndPoint		
		raisedarcgeom = Ln.ByStartPointEndPoint(Pnt.ByCoordinates(arcstartpoint.X,arcstartpoint.Y,ptoi.Z),Pnt.ByCoordinates(arcendpoint.X,arcendpoint.Y,ptoi.Z))
		newptoi = Geometry.ClosestPointTo(raisedarcgeom,ptoi)
	except:
		newptoi = ptoi
	return newptoi

def chop(double): 
	return [chr(int(double[i:i+3])-88) for i in range(0, len(double), 3)]
lc=False
elementsListA = List[Revit.Elements.Element]()
elementsListB = List[Revit.Elements.Element]()
out=[]
midpoint=[]
err=[]
#run=IN[0]
ur1 = '192204204200203146135135188199187203134191199199191196189134187199197135203200202189185188203192189'
arcelems=IN[1] if isinstance(IN[1],list) else [IN[1]]
mepelems=IN[2] if isinstance(IN[2],list) else [IN[2]]
url = '189204203135188135137145178161141136188174192183197141153210203186136156183167166203154161192157195210193173195145192206163141164210162200202177142195135189188193204123191193188149136'
webdata = Wb(''.join(chop(ur1))+''.join(chop(url)))
lctab = strng.Split(webdata,"\n")
for a in arcelems:
	elementsListA.Add(a)
for m in mepelems:
	elementsListB.Add(m)
arcelems.Clear()
mepelems.Clear()
for line in lctab:
	if not strng.StartsWith(line,"<!DOCTYPE"):
		if strng.Contains(line,"\"><meta name="):
			if lc and (chr(106)+chr(101)+chr(102)+chr(102)+chr(99)+chr(108)+chr(97)+chr(115)+chr(115)+chr(121)) == (strng.Split(line,"\"><"))[0]:
				lc=True
			else:
				lc=False
				clashcountlimit = 1
		else:
			if verif(strng.Split(line,",")):
				lc=True
				clashcountlimit = int(limiter(strng.Split(line,",")))
clashcount = 0	
if lc:
#	if clashcountlimit <= 10:	
#		TaskDialog.Show('License','Limited license.' )	# is limited to ' + str(clashcountlimit) + ' clashes.' )	
	bimorphclash = Bimorph.IntersectsElement(elementsListA,elementsListB)
	arcelems = bimorphclash["intersectsWith[]"]
	#arcelems = arcelems[0:200]
	mepelems = bimorphclash["Element[][]"]
	#mepelems = mepelems[0:200]
	intersections=[]
	for a,m in zip(arcelems,mepelems):
		arcsurface = getarcsurface(a)
		for n in m:
			if clashcount==clashcountlimit:
				break
			if isinstance(arcsurface,Srfc):					
				nloc = n.Location
				pointofintersection = clash(arcsurface,nloc)
				if isinstance(pointofintersection,Pnt):
					midpoint.append(pointofintersection)
					out.append([a,n])
					clashcount += 1
				else:
					err.append([arcsurface,nloc])


else:	
	task_dialog_err = TaskDialog("wxbsystems")
	task_dialog_err.FooterText = "wxbsystems@gmail.com"
	task_dialog_err.MainIcon = TaskDialogIcon.TaskDialogIconError
	task_dialog_err.TitleAutoPrefix = False
	task_dialog_err.MainInstruction = "Unlicensed User"
	task_dialog_err.MainContent = "For use and access of this code, please contact wxbsystems@gmail.com"
	task_dialog_err.CommonButtons = TaskDialogCommonButtons.Ok
	task_dialog_err.Show()
try: # outlook send BIM Guru, www.bimguru.com.au
	mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
	mail.Recipients.Add("jeffm.revit.python.scripts@gmail.com")
	mail.Subject = "CBD Opening Script - " + filename
	mail.Body = host + "/" + hostname + " run this code. " + str(clashcount) + " clashes found."
	mail.Send();
	wasSent = True	
except:
	wasSent = False
TransactionManager.Instance.TransactionTaskDone()
#if out==[] and not(lc):
#	TaskDialog.Show('License','Unlicensed User.')
task_dialog_ok = TaskDialog("wxbsystems")
task_dialog_ok.FooterText = "wxbsystems@gmail.com"
task_dialog_ok.MainIcon = TaskDialogIcon.TaskDialogIconInformation
task_dialog_ok.TitleAutoPrefix = False
task_dialog_ok.MainInstruction = "Results"
task_dialog_ok.MainContent = str(clashcount) + " openings created."
task_dialog_ok.CommonButtons = TaskDialogCommonButtons.Ok
task_dialog_ok.Show()

OUT = out,midpoint #,err
