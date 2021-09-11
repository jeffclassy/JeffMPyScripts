#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
#192204204200203146135135202185207134191193204192205186205203189202187199198204189198204134187199197135194189190190187196185203203209135162189190190165168209171187202193200204203135164193187189198203189188133171187202193200204203135162189190190165183164193187189198203189188125138136155196185203192125138136164199191193187183174137134137134200209
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
from Autodesk.DesignScript.Geometry import Surface as Srfc

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
	intersect = 0
	meploc = mep.Location
	arcloc = arc.Location

	if isinstance(meploc,GeomCurves) and isinstance(arcloc,GeomCurves) and UnwrapElement(arc).Category.Name =='Walls':		
		#arcloctop = GeomCurves.Translate(arcloc,Vector.ZAxis(),arc.GetParameterValueByName("Unconnected Height"))
		arcsurface = GeomCurves.Extrude(arcloc,Vector.ZAxis(),arc.GetParameterValueByName("Unconnected Height")) #Srfc.ByLoft([arcloc,arcloctop])		
		intersect = Geometry.Intersect(arcsurface,meploc) #line to surface #Surface vs Line given the line crosses the surface
		if intersect == []:
			intersect = [Geometry.ClosestPointTo(arcsurface,meploc)]
		intersect = intersect[0]
	elif isinstance(meploc,GeomCurves) and isinstance(arcloc,Pnt):
		arcangle = math.degrees(UnwrapElement(a).Location.Rotation)  #Geometry.Intersect(Sld.ByUnion(arc.Solids),meploc)
		basepoint1 = GeomCurves.Translate(arcloc,Vector.XAxis(),1000)
		basepoint2 = GeomCurves.Translate(arcloc,Vector.XAxis(),-1000)
		arcline = Ln.ByStartPointEndPoint(basepoint1,basepoint2)
		arcsurface = GeomCurves.Extrude(arcline,Vector.ZAxis(),arc.GetParameterValueByName("Height")) #assuming Height contains value
		arcsurface = Geometry.Rotate(arcsurface,arcloc,Vector.ZAxis(),arcangle)
		intersect = Geometry.Intersect(arcsurface,meploc)
		if intersect == []:
			intersect = [Geometry.ClosestPointTo(arcsurface,meploc)]
		intersect = intersect[0]
		#if isinstance(intersect[0],GeomCurves):
		#	intersect = GeomCurves.PointAtParameter(intersect[0],0.5)
	else:
		intersect = 11

	return intersect		
def getmidpoint (crv):
	try:
		ptoi = GeomCurves.PointAtParameter(crv,0.5)
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

def chop(double): 
	return [chr(int(double[i:i+3])-88) for i in range(0, len(double), 3)]
lc=False
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
lc=True				
if lc:
	bimorphclash = Bimorph.IntersectsElement(elementsListA,elementsListB)
	arcelems = bimorphclash["intersectsWith[]"]
	#arcelems = arcelems[0:200]
	mepelems = bimorphclash["Element[][]"]
	#mepelems = mepelems[0:200]
	intersections=[]
	clashcount = 0
	clashcountlimit = 200
	for a,m in zip(arcelems,mepelems):
		for n in m:
			intersections = clash(a,n) #if lc else 0
			pointofintersection = intersections
			if isinstance(pointofintersection,Pnt):
				#wallcrosspoint = pointclosetowallmid(a,pointofintersection)
				midpoint.append(pointofintersection)
				out.append([a,n])
				clashcount = clashcount + 1
			if clashcount >= clashcountlimit:
				break
		if clashcount >= clashcountlimit:
			break
else:	
	TaskDialog.Show('License','Unlicensed User.')
TransactionManager.Instance.TransactionTaskDone()
#if out==[] and not(lc):
#	TaskDialog.Show('License','Unlicensed User.')
OUT = out,midpoint#,bimorphclash#.GetRotation
