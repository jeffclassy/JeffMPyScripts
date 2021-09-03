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
from Autodesk.DesignScript.Geometry import Point as Pnt
from Autodesk.DesignScript.Geometry import Vector as Vctr



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

clr.AddReference('BimorphNodes')
from Revit import Element as Bimorph#, LinkElement
from Revit import LinkElement as LinkedBimorph

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

elementsListA = List[Revit.Elements.Element]()
elementsListB = List[Revit.Elements.Element]()

out=[]
run=IN[0][1]
rvtlinkinstance1=IN[0][0][0]
rvtlinkinstance2=IN[0][0][2]
ur1 = '19ZI50dVh_m5Azsb0D_ONsBIhEkziUk9hvK5LzJprY6k/edit?usp=sharing/'

category1=Revit.Elements.Category.ByName(IN[0][0][1])
category2=Revit.Elements.Category.ByName(IN[0][0][3])
arcelems=LinkedBimorph.OfCategory(rvtlinkinstance1,category1)
mepelems=LinkedBimorph.OfCategory(rvtlinkinstance2,category2)
url = 'https://docs.google.com/spreadsheets/d/'	
webdata = wb.WebRequestByUrl(url+ur1)
lctab = strng.Split(webdata,"\n")
for a in arcelems:
	elementsListA.Add(a)
for m in mepelems:
	elementsListB.Add(m)
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
bimorphclash = Bimorph.IntersectsElement(elementsListA,elementsListB) if lc else []
try:
	for a,m in zip(arcelems,bimorphclash["Element[][]"]):
		for n in m:
			try:
				intersections = Geometry.Intersect(BoundingBox.ToCuboid(a.BoundingBox),n.Location)[0]
				nunwrap = UnwrapElement(n)
				ntype = nunwrap.Document.GetElement(nunwrap.GetTypeId())		
				aunwrap = UnwrapElement(a)
				atype = aunwrap.Document.GetElement(aunwrap.GetTypeId())	
				try:
					pointofintersection = GeomCurves.PointAtParameter(intersections,0.5)
				except:
					pointofintersection = intersections
				try:
					#rotangle = round(math.degrees(Vctr.AngleAboutAxis(a.Location.Direction,Vctr.XAxis(),Vctr.ZAxis()))%360,1)
					rotangle = (Vctr.AngleAboutAxis(a.Location.Direction,Vctr.XAxis(),Vctr.ZAxis()))
				except:
					try:
						rotangle = math.degrees((aunwrap.Location).Rotation)
					except:
						rotangle = 0
				rot=round(fixrotation(rotangle),1)
				out.append([a,n,intersections,atype,ntype,pointofintersection,rot])
			except:
				pass
except:
	out=[]
OUT = out
