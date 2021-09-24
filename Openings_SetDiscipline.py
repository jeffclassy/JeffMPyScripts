#Copyright JeffM Python Scripts
#jayperson.mercado@gmail.com

import sys
import clr
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Start Transaction
doc = DocumentManager.Instance.CurrentDBDocument
TransactionManager.Instance.EnsureInTransaction(doc)
# The inputs to this node will be stored as a list in the IN variables.

disciplinelist = ["AC","AS","DR","EL","FS","MEP","PL","TEMP","TG"]
el = IN[0] if isinstance(IN[0], list) else [IN[0]]
elems = UnwrapElement(el)
parameters = IN[1] if isinstance(IN[1], list) else [IN[1]]
disc="AC"
for e,d in zip (elems,parameters):
	for disc in disciplinelist:
		param = e.LookupParameter(disc)
		if disc==d:
			param.Set(1)
		else:
			param.Set(0)
		

TransactionManager.Instance.TransactionTaskDone()
OUT = elems
