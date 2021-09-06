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

clr.AddReference('DSCoreNodes') 
from DSCore import Web as wb
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
catA=IN[0][0][1]
catB=IN[0][0][3]

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

def mepparameters (mepelems):
	if catB == "Ducts":
		return ductparams(mepelems)
	elif catB == "Cable Trays":
		return trayparams(mepelems)
	elif catB == "Pipes" or catB == "Conduits":
		return pipeparams(mepelems)
	else:
		return ductparams(mepelems)
ductparameters = ["Width","Height","System Type","Workset"]
trayparameters = ["Width","Height","Service Type","Workset"]
pipeparameters = ["Diameter","System Abbreviation","Workset"]
typeparameters = []

def ductparams (mepelems):
	dp=[]
	for p in ductparameters:
		try:
			dp.append (LinkedBimorph.GetParameterValueByName(mepelems,p))
		except:
			dp.append("")
	return dp
def trayparams (mepelems):
	trayp=[]
	for p in trayparameters:
		try:
			trayp.append (LinkedBimorph.GetParameterValueByName(mepelems,p))
		except:
			trayp.append("")
	return trayp

def pipeparams (mepelems):	
	pipep=[]
	for p in pipeparameters:
		try:
			pipep.append (LinkedBimorph.GetParameterValueByName(mepelems,p))
		except:
			pipep.append("")
	return pipep

elementsListA = List[Revit.Elements.Element]()
elementsListB = List[Revit.Elements.Element]()

out=[]
run=IN[0][1]
rvtlinkinstance1=IN[0][0][0]
rvtlinkinstance2=IN[0][0][2]
ur1 = '19ZI50dVh_m5Azsb0D_ONsBIhEkziUk9hvK5LzJprY6k/edit?usp=sharing/'

category1=Revit.Elements.Category.ByName(catA)
category2=Revit.Elements.Category.ByName(catB)
arcelems=LinkedBimorph.OfCategory(rvtlinkinstance1,category1)
mepelems=LinkedBimorph.OfCategory(rvtlinkinstance2,category2)
url = 'https://docs.google.com/spreadsheets/d/'	
webdata = wb.WebRequestByUrl(url+ur1)
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
mepelems = bimorphclash["Element[][]"]
intersections=[]
for a,m in zip(arcelems,mepelems):
	for n in m:
		intersections = clash(a,n)
		if not(intersections == 0 or intersections==[]):
			pointofintersection = getmidpoint(intersections)
		else:
			pointofintersection = intersections
		if isinstance(pointofintersection,Pnt):
			nunwrap = UnwrapElement(n)
			aunwrap = UnwrapElement(a)
			ntype = nunwrap.Document.GetElement(nunwrap.GetTypeId())		
			atype = aunwrap.Document.GetElement(aunwrap.GetTypeId())
			nparams = mepparameters(n)
			out.append([a,n,aunwrap,ntype,pointofintersection,nparams])
		#intersections=[]
TransactionManager.Instance.TransactionTaskDone()
OUT = out
