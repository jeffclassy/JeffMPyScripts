#Copyright(c) 2021 by Jeff Mercado
#jayperson.mercado@gmail.com
#based on dimitar, rhythm, and bimorph

import clr

clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
out=[]
clr.AddReference("RevitNodes")
import Revit

clr.AddReference("BimorphNodes")
import Revit
from Revit import *

clr.AddReference("DSCoreNodes")
import DSCore
from DSCore import *

import re
reg1 = re.compile(".+Wall")

clr.ImportExtensions(Revit.Elements)
catstr=IN[1]
cat1 = Revit.Elements.Category.ByName(catstr)
cat2 = Revit.Elements.Category.ByName("Columns")
link1 = UnwrapElement(IN[2])
linkDoc = link1.GetLinkDocument()
linkedelement1 = LinkElement.OfCategory(link1,cat1)
linkedelement2 = LinkElement.OfCategory(link1,cat2)
linkedelement = []
for e in linkedelement1:
	linkedelement.append(e)
for e in linkedelement2:
	linkedelement.append(e)
linkedelementid = []
for l in linkedelement:
	linkedelementid.append(l.Id)

TaskDialog.Show('Linked Selection','Pick elements of category ' + (catstr.lower()) + '. Press Finish in upper left to complete.')

sel1 = uidoc.Selection
obt1 = Selection.ObjectType.LinkedElement
class CustomISelectionFilter(Selection.ISelectionFilter):
	def __init__(self, nom_categorie):
		self.nom_categorie = nom_categorie
	def AllowElement(self, e):
		return True
	def AllowReference(self, ref, point):
		e=linkDoc.GetElement(ref.LinkedElementId).ToDSType(False)
		if e.GetCategory.Name == self.nom_categorie or ( catstr == "Walls" and (e.GetCategory.Name == "Columns" and bool(re.match(reg1,e.Name)))):
			return True
		else:
			return False
		
el_ref = sel1.PickObjects(obt1,CustomISelectionFilter(catstr),   'Pick elements of category, ' + (IN[1].lower()) + ' JeffM finish to complete.')
typelist = list()
idlist = list()
for i in el_ref:	
	try:
		#typelist.append(linkDoc.GetElement(i.LinkedElementId).ToDSType(True))
		try:
			for j,k in zip (linkedelementid,linkedelement):
				if i.LinkedElementId.ToString() == j.ToString():
					out.append(k)
		except:
			pass
	except:
		pass		
OUT =  List.UniqueItems(out)#typelist,out


#Copyright(c) 2021 by Jeff Mercado
#jayperson.mercado@gmail.com
#based on dimitar, rhythm, and bimorph
