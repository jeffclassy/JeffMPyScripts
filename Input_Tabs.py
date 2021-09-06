import clr

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('DSCoreNodes') 
import DSCore
from DSCore import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

rvtLinkCollector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
rvtLinks = rvtLinkCollector.ToElements()

#revit links
instancename=[]
for rvt in rvtLinks:
	name = rvt.get_Parameter(BuiltInParameter.RVT_LINK_INSTANCE_NAME).AsString()
	linkname = Element.Name.GetValue(rvt)	
	name = DSCore.String.Split(linkname,".rvt ")
	instancename.append(name[0])
	
#datashapes dropdown forms
class uidropdown():
    def __init__(self,inputname,defaultvalue,_sorted):
        self.inputname = inputname
        self.defaultvalue = defaultvalue
        self._sorted = _sorted
    def __setitem__(self, key, item):
        self.__dict__[key] = item
    def __getitem__(self, key):
        return self.__dict__[key]
    def __iter__(self):
        return iter(self.__dict__)
    def values(self):
        return self.__dict__.values()
    def keys(self):
        return self.__dict__.keys()
    def __repr__(self):
        return 'UI.DropDown input'
        
class uiradio():
    def __init__(self,inputname,defaultvalue):
        self.inputname = inputname
        self.defaultvalue = defaultvalue
    def __setitem__(self, key, item):
        self.__dict__[key] = item
    def __getitem__(self, key):
        return self.__dict__[key]
    def __iter__(self):
        return iter(self.__dict__)
    def values(self):
        return self.__dict__.values()
    def keys(self):
        return self.__dict__.keys()
    def __repr__(self):
        return 'UI.Radio input'
class uibool():
    def __init__(self,inputname,defaultvalue,booltext):
        self.inputname = inputname
        self.defaultvalue = defaultvalue
        self.booltext = booltext
    def __repr__(self):
        return 'UI.Boolean input'	
        
uidroparchi = uidropdown("Select Archi/Struc Model",0,False)
uidropmep = uidropdown("Select MEP Model",1,False)
in1 = instancename #list
in2 = rvtLinks #list
#if isinstance(instancename,list):
#	in1 = instancename
#else:
#	in1 = [instancename]

#links dropdown
for k,v in zip(in1,in2):
	try:
		uidroparchi[str(k)] = v
		uidropmep[str(k)] = v
	except:
		uidroparchi[k.encode('utf-8').decode('utf-8')] = v
		uidropmep[k.encode('utf-8').decode('utf-8')] = v

#category radiobox
uiradioarccategory = uiradio("Select Category",0)
in3 = ["Walls","Columns","Floors","Structural Framings"]
in4 = ["Walls","Columns","Floors","Structural Framings"]
for k,v in zip(in3,in4):
	try:
		uiradioarccategory[str(k)] = v
	except:
		uiradioarccategory[k.encode('utf-8').decode('utf-8')] = v
uiradiomepcategory = uiradio("Select Category",0)
in5 = ["Cable Trays","Conduits","Ducts","Duct Accessories","Pipes"]
in6 = ["Cable Trays","Conduits","Ducts","Duct Accessories","Pipes"]
for k,v in zip(in5,in6):
	try:
		uiradiomepcategory[str(k)] = v
	except:
		uiradiomepcategory[k.encode('utf-8').decode('utf-8')] = v
#boolean input
#uibool= uibool("Fire Dampers",False,"True")
uiboxes = [uidroparchi,uiradioarccategory,uidropmep,uiradiomepcategory] 

OUT = uiboxes
