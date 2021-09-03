#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com
import sys
#sys.path.append(C....\bimorphNodes\bin')
import clr

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
from Autodesk.DesignScript.Geometry import Curve as GeomCurves

import asd
asd.path.append('C:\Users\jaypm.mercado\Downloads')
asd.AddReference('BimorphNodes')
from Revit import Element as Bimorph#, LinkElement
from Revit import LinkElement as LinkedBimorph

clr.AddReference('System')
from System.Collections.Generic import List

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

elementsListA = List[Revit.Elements.Element]()
elementsListB = List[Revit.Elements.Element]()

out=[]
rvtlinkinstance1=IN[0][0][0]
rvtlinkinstance2=IN[0][0][2]
category1=Revit.Elements.Category.ByName(IN[0][0][1])
category2=Revit.Elements.Category.ByName(IN[0][0][3])
arcelems=LinkedBimorph.OfCategory(rvtlinkinstance1,category1)
mepelems=LinkedBimorph.OfCategory(rvtlinkinstance2,category2) 
run=IN[0][1]
for a in arcelems:
	elementsListA.Add(a)
for m in mepelems:
	elementsListB.Add(m)
bimorphclash = Bimorph.IntersectsElement(elementsListA,elementsListB) if run else []
try:
	for a,m in zip(arcelems,bimorphclash["Element[][]"]):
		for n in m:
			try:
				intersections = Geometry.Intersect(BoundingBox.ToCuboid(a.BoundingBox),n.Location)[0]
				try:
					pointofintersection = GeomCurves.PointAtParameter(intersections,0.5)
				except:
					pointofintersection = intersections
				nunwrap = UnwrapElement(n)
				ntype = nunwrap.Document.GetElement(nunwrap.GetTypeId())		
				aunwrap = UnwrapElement(a)
				atype = aunwrap.Document.GetElement(aunwrap.GetTypeId())	
				out.append([a,n,intersections,atype,ntype,pointofintersection])
			except:
				pass
except:
	out=[]
OUT = out
