#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
import sys
#sys.path.append(C....\bimorphNodes\bin')
#import os as os
import socket
import getpass
import math
host = socket.gethostname()
hostname = getpass.getuser()
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
from Autodesk.DesignScript.Geometry import Curve as GeomCurves
from Autodesk.DesignScript.Geometry import Solid as Sld
from Autodesk.DesignScript.Geometry import Point as Pnt
from Autodesk.DesignScript.Geometry import Line as Ln

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

curvelistB = List[GeomCurves]()
origin = Pnt.ByCoordinates(0,0)
empty=[]

def verif (d):
	if d[0] == host and d[1] == hostname and d[2] == chr(49):
		return True
def fixrotation (rot):
	if rot==180:
		rot=0
	elif rot>180:
		rot=abs(rot-360)
	else:
		rot
	return rot
def clash (arc,mep):
	try:
		meploc = mep.Location
		if isinstance(meploc,GeomCurves):
			mepgeom = GeomCurves.TrimByEndParameter(meploc,0.999)
			curvelistB.Add(mepgeom)		
		else:			
			mepgeom = Ln.ByStartPointEndPoint(meploc,Pnt.ByCoordinates(meploc.X,meploc.Y,meploc.Z+1))
			curvelistB.Add(mepgeom)
		intersect = BMCurve.SolidIntersection(Sld.ByUnion(arc.Solids),curvelistB)["Curve[]"]
		curvelistB.Clear()	
	except:
		intersect = 0
	return intersect		
def getmidpoint (crv):
	try:
		ptoi = GeomCurves.PointAtParameter(lst.FirstItem(crv),0.5)
		return ptoi #point of intersection
	except:
		return crv
def pointclosetowallmid (arc,ptoi):
	try:
		arcstartpoint = a.Location.StartPoint
		arcendpoint = a.Location.EndPoint		
		raisedarcgeom = Ln.ByStartPointEndPoint(Pnt.ByCoordinates(arcstartpoint.X,arcstartpoint.Y,ptoi.Z),Pnt.ByCoordinates(arcendpoint.X,arcendpoint.Y,ptoi.Z))
		newptoi = Geometry.ClosestPointTo(raisedarcgeom,ptoi)
	except:
		newptoi = ptoi
	return newptoi

def chop(double): #decrypts every 3char from a string of encrypted url
	return [chr(int(double[i:i+3])-88) for i in range(0, len(double), 3)]
lc=1==2
elementsListA = List[Revit.Elements.Element]()
elementsListB = List[Revit.Elements.Element]()
out=[]
midpoint=[]
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
lc=False
for line in lctab:
	if not strng.StartsWith(line,"<!DOCTYPE"):
		if strng.Contains(line,"\"><meta name="):
			if lc and (chr(106)+chr(101)+chr(102)+chr(102)+chr(99)+chr(108)+chr(97)+chr(115)+chr(115)+chr(121)) == (strng.Split(line,"\"><"))[0]:
				lc=True
			else:
				lc=False
		else:
			if verif(strng.Split(line,",")):
				lc=True			
bimorphclash = Bimorph.IntersectsElement(elementsListA,elementsListB)# if 
arcelems = bimorphclash["intersectsWith[]"]
arcelems = arcelems[0:200]
mepelems = bimorphclash["Element[][]"]
mepelems = mepelems[0:200]
intersections=[]
for a,m in zip(arcelems,mepelems):
	for n in m:
		intersections = clash(a,n) if lc else 0
		if not(intersections == 0 or intersections==[]):
			pointofintersection = getmidpoint(intersections) if lc else 0
		else:
			pointofintersection = intersections
		if isinstance(pointofintersection,Pnt):
			wallcrosspoint = pointclosetowallmid(a,pointofintersection)
			midpoint.append(wallcrosspoint)
			out.append([a,n])
		#intersections=[]
TransactionManager.Instance.TransactionTaskDone()
if out==[] and not(lc):
	TaskDialog.Show('License','Unlicensed User.')
OUT = out,midpoint
