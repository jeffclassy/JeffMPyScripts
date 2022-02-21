# Created by JeffMercado jayperson.mercado@gmail.com
# Inspired from lucamanzoni for getting placement points
# from viktor_kuzev for  moving points to new location
# and clockwork for elementbyid


import clr

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

TransactionManager.Instance.EnsureInTransaction(doc)

def ElementById(item, doc):
	try: 
		return doc.GetElement(item).ToDSType(True)
	except:
		try:
			return doc.GetElement(ElementId(item)).ToDSType(True)
		except:
			return None

def moverefpoint(a, b, doc):
	try: 
		outlist=[]
		for i,j in zip(a,b):
			x=j.ToXyz()
			y=UnwrapElement(i)
			trans=(x.Subtract(y.Position))
			ElementTransformUtils.MoveElement(doc, y.Id, trans)
			outlist.append(y)
		return outlist
	except:
		return None

if isinstance(IN[0],list):
	element = UnwrapElement(IN[0])
else:
	element = [UnwrapElement(IN[0])]

pnt = []
for x in element:
	points = AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(x)	 
	refpoint = [ElementById(y, doc) for y in points]
	pnt.append(refpoint)
	
#pnt = IN[0]
pt = IN[1]

if isinstance(pnt, list): 
	out = [moverefpoint(a, b, doc) for a,b in zip(pnt,pt)]
else: 
	out = moverefpoint(pnt, pt, doc)
	
TransactionManager.Instance.TransactionTaskDone()
OUT = out





