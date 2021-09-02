#Copyright (c) Data Shapes,  2020
#Data-Shapes www.data-shapes.io , elayoubi.mostafa@data-shapes.io @data_shapes
	
import clr
import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)
import os
import webbrowser
import unicodedata
import io
import tempfile
import System
import socket
import getpass
	
try:
	clr.AddReference('System.Windows.Forms')
	clr.AddReference('System.Drawing')
	clr.AddReference('System.Windows.Forms.DataVisualization')
	
	from System.Drawing import Point , Size , Graphics, Bitmap, Image, Font, FontStyle, Icon, Color, Region , Rectangle , ContentAlignment , SystemFonts, FontFamily
	from System.Windows.Forms import Application, DockStyle, Button, Form, Label, TrackBar , ToolTip, ColumnHeader, TextBox, CheckBox, FolderBrowserDialog, OpenFileDialog, DialogResult, ComboBox, FormBorderStyle, FormStartPosition, ListView, ListViewItem , SortOrder, Panel, ImageLayout, GroupBox, RadioButton, BorderStyle, PictureBox, PictureBoxSizeMode, LinkLabel, CheckState, ColumnHeaderStyle , ImageList, VScrollBar, DataGridView, DataGridViewSelectionMode, DataGridViewAutoSizeColumnsMode , DataGridViewClipboardCopyMode , TreeView , TreeNode , TreeNodeCollection , AutoScaleMode , Screen, Padding
	from System.Windows.Forms.DataVisualization.Charting import *#Chart , SeriesChartType
	from System.Collections.Generic import *
	from System.Collections.Generic import List as iList
	from System.Windows.Forms import View as vi
	clr.AddReference('System')
	from System import IntPtr , Char
	from System import Type as SType, IO
	from System import Array
	from System.ComponentModel import Container
	clr.AddReference('System.Data')
	from System.Data import DataTable , DataView

	try: #try to import All Revit dependencies
		clr.AddReference('RevitAPIUI')
		from  Autodesk.Revit.UI import Selection , TaskDialog 
		from  Autodesk.Revit.UI.Selection import ISelectionFilter
		clr.AddReference('RevitNodes')
		import Revit
		clr.ImportExtensions(Revit.Elements)
		clr.ImportExtensions(Revit.GeometryConversion)
		clr.AddReference('DSCoreNodes') 
		import DSCore
		from DSCore import *
		clr.AddReference('RevitServices')
		from RevitServices.Persistence import DocumentManager
		doc = DocumentManager.Instance.CurrentDBDocument
		uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
	
		clr.AddReference('RevitAPI')
		try:
			from Autodesk.Revit.DB import ImageImportOptions	
		except:
			from Autodesk.Revit.DB import ImageTypeOptions , ImageType, ImagePlacementOptions , ImageInstance
		from Autodesk.Revit.DB import FilteredElementCollector , Transaction, View , ViewType , ViewFamily, ViewDrafting, ViewFamilyType, Element, ElementId , FamilyInstance , Document , XYZ, BoxPlacement, UnitType , UnitUtils
		dbviews = [v for v in FilteredElementCollector(doc).OfClass(View).ToElements() if (v.ViewType == ViewType.FloorPlan or v.ViewType == ViewType.CeilingPlan or v.ViewType == ViewType.Section or v.ViewType == ViewType.Elevation or v.ViewType == ViewType.ThreeD)]
		viewindex = 0
		UIunit = Document.GetUnits(doc).GetFormatOptions(UnitType.UT_Length).DisplayUnits
		class selectionfilter(ISelectionFilter):
			def __init__(self,category):
				self.category = category
			def AllowElement(self,element):
				if element.Category.Name in [c.Name for c in self.category]:
					return True
				else:
					return False
			def AllowReference(reference,point):
				return False
	except: #in case we are in the Sandbox, Formit or Civil 3D environment
		pass
	
	importcolorselection = 0
	
	try:
		from  Autodesk.Revit.UI import ColorSelectionDialog
	except:
		importcolorselection = 1
	

	
	clr.AddReference('ProtoGeometry')
	from Autodesk.DesignScript.Geometry import Point as dsPoint

	from System.Reflection import Assembly
	import xml.etree.ElementTree as et
	

	
	import re	
	def regexEndNum(input):
		try:
			return 	re.search('(\d+)$', input).group(0)
		except:
			return ""

	def iterateThroughNodes(collection,li):
		if hasattr(collection,'Nodes'):
			ntest = collection.Nodes
			if len(ntest) > 0:
				for i in ntest:
					iterateThroughNodes(i,li)
			else:
				if collection.Checked:
					li.append(collection.Tag)
		return li

	class MultiTextBoxForm(Form):
		
	    def __init__(self):
	        self.Text = 'Data-Shapes + JeffM + Inputbox'
	        self.output = []
	        self.values = []
	        self.cancelled = False
	
	    def setclose(self, sender, event):
	    	cbindexread = 0
	    	if sender.Name != "Cancel":
	    		for f in self.output:
	    			if f.GetType() == myTextBox:
	    				if f._isNum :
	    					val = float(f.Text)
	    				else:
	    					val = f.Text
	    				self.values.append(val)
	    			if f.GetType() == CheckBox:
	    				self.values.append(f.Checked)
	    			if f.GetType() == Button:
	    				if isinstance(f.Tag ,list):
	    					try:
	    						self.values.append([e for e in f.Tag if e.__class__.__name__ != "Category"])	    					
	    					except:
	    						self.values.append(f.Tag)	    					
	    				else:
	    					try:	    				
	    						if f.Tag.__class__.__name__ != "Category":
	    							self.values.append(f.Tag)
	    						else:
    								self.values.append([])
	    					except:
	    						self.values.append(f.Tag)	
	    			if f.GetType() == ComboBox:
	    				try:
	    					key = f.Text
	    					self.values.append(f.Tag[key])
	    				except:
	    					self.values.append(None)
	    			if f.GetType() == mylistview:
	    				self.values.append([f.Values[i.Text] for i in f.CheckedItems])
	    			if f.GetType() == mytrackbar:
	    				self.values.append(f.startval+f.Value*f.step)
	    			if f.GetType() == mygroupbox:
	    				try:
	    					key = [j.Text for j in f.Controls if j.Checked == True][0]
	    					self.values.append(f.Tag[key])
	    				except:
	    					self.values.append(None)
	    			if f.GetType() == myDataGridView:
	    				f.EndEdit()
	    				dsrc = f.DataSource
	    				out = []
	    				colcount = f.ColumnCount
	    				rowcount = f.RowCount - 1
	    				if f.Tag:
		    				l = []
	    					for i in range(colcount):	    						
	    						l.append(dsrc.Columns[i].ColumnName)
	    					out.append(l)	    						
		    				for r in range(rowcount):
		    					l = []
		    					for i in range(colcount):
		    						l.append(dsrc.DefaultView[r].Row[i])
		    					out.append(l)
	    				else: 
		    				for r in range(rowcount):
		    					l = []
		    					for i in range(colcount):
		    						l.append(dsrc.DefaultView[r].Row[i])
		    					out.append(l)
	    				self.values.append(out)
	    			if f.GetType() == TreeView:
	    				ls = []
	    				nds = f.Nodes[0]
	    				iterateThroughNodes(nds,ls)
	    				self.values.append(ls)
	    			if f.GetType() == GroupBox:
	    				rb = [c for c in f.Controls if c.GetType() == RadioButton and c.Checked][0]
	    				self.values.append(rb.Text)
	    				f.Controls.Remove(rb)
	    	else:
	    		self.values = None
	    		self.cancelled = True
	    	self.Close()
	
	    def reset(self, sender, event):
			pass
	
	    def openfile(self, sender, event):
			ofd = OpenFileDialog()
			dr = ofd.ShowDialog()
			if dr == DialogResult.OK:
				sender.Text = ofd.FileName
				sender.Tag = ofd.FileName

	    def exportToExcel(self, sender, event):
	    	#importing Excel IronPython libraries
	    	clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
	    	from Microsoft.Office.Interop import Excel
	    	ex = Excel.ApplicationClass()
	    	ex.Visible = sender.Tag[1]
	    	ex.DisplayAlerts = False
	    	fbd = FolderBrowserDialog()
	    	fbd.SelectedPath = sender.Text
	    	parent = sender.Parent
	    	fptextbox = parent.GetChildAtPoint(Point(parent.Location.X,sender.Location.Y+5*yRatio))
	    	dataGrid = parent.GetChildAtPoint(Point(parent.Location.X,parent.Location.Y+23*xRatio))
	    	dataTable = dataGrid.DataSource
	    	fptext = fptextbox.Text
	    	titletext = parent.GetChildAtPoint(Point(0,0)).Text
	    	dr = fbd.ShowDialog()
	    	frstRwTtle = sender.Tag[0]
	    	if frstRwTtle:
	    		_header = Excel.XlYesNoGuess.xlYes
	    	else:
	    		_header = Excel.XlYesNoGuess.xlNo
	    	if dr == DialogResult.OK:
	    		workbk = ex.Workbooks.Add()
	    		worksheet = workbk.Worksheets.Add()
	    		#Writing title and doc info
	    		if sender.Tag[2]:
		    		titleCell = worksheet.Cells[1,1]
		    		worksheet.Cells[2,1].Value2 = sender.Tag[3]
		    		titleCell.Value2 = titletext
		    		titleCell.Font.Size = 18
		    		titleCell.Font.Bold = True
		    		startR = 3
		    		endR = 3
		    	else:
		    		startR = 1
		    		endR = 0	
		    	if frstRwTtle:
		    		for j in range(0,dataTable.Columns.Count):
		    			worksheet.Cells[startR,j+1] = dataTable.Columns[j].ColumnName		    	
		    		for i in range(0,dataTable.Rows.Count):
		    			for j in range(0,dataTable.Columns.Count):
			    			worksheet.Cells[i+startR+1,j+1] = dataTable.DefaultView[i].Row[j].ToString()
		    		xlrange = ex.get_Range(worksheet.Cells[startR,1],worksheet.Cells[dataTable.Rows.Count+endR+1,dataTable.Columns.Count])			    			
		    	else :
		    		for i in range(0,dataTable.Rows.Count):
		    			for j in range(0,dataTable.Columns.Count):
			    			worksheet.Cells[i+startR,j+1] = dataTable.DefaultView[i].Row[j].ToString()		    	
		    		xlrange = ex.get_Range(worksheet.Cells[startR,1],worksheet.Cells[dataTable.Rows.Count+endR,dataTable.Columns.Count])
		    	xlrange.Columns.AutoFit()
		    	worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, xlrange, SType.Missing, _header, SType.Missing).Name = "DataShapesTable"
		    	worksheet.ListObjects["DataShapesTable"].TableStyle =  "TableStyleMedium16"
	    		workbk.SaveAs(fbd.SelectedPath + "\\" + fptext)
	    		if not sender.Tag[1]:
	    			workbk.Close()
	    			ex.Quit()

	    def startCell(self, sender, event ):
			sender.startcell["X"] = event.ColumnIndex
			sender.startcell["Y"] = event.RowIndex	

	    def endCell(self, sender, event ):
	    	try:
				sender.endcell["X"] = event.ColumnIndex
				sender.endcell["Y"] = event.RowIndex 
				startval = sender.Rows[sender.startcell["Y"]].Cells[sender.startcell["X"]].Value
				endNum = regexEndNum(startval)
				if endNum != "":
					if sender.endcell["Y"] == sender.startcell["Y"]:
						for e,i in enumerate(range(sender.startcell["X"],sender.endcell["X"] + 1)):
							sender.Rows[sender.startcell["Y"]].Cells[i].Value = startval[:-len(endNum)] + str(int(endNum) + e)
					elif sender.endcell["X"] == sender.startcell["X"]:
						for e,i in enumerate(range(sender.startcell["Y"],sender.endcell["Y"] + 1)):
							sender.Rows[i].Cells[sender.endcell["X"]].Value = startval[:-len(endNum)] + str(int(endNum) + e)
				else:				
					if sender.endcell["Y"] == sender.startcell["Y"]:
						for i in range(sender.startcell["X"],sender.endcell["X"] + 1):
							sender.Rows[sender.startcell["Y"]].Cells[i].Value = startval
					elif sender.endcell["X"] == sender.startcell["X"]:
						for i in range(sender.startcell["Y"],sender.endcell["Y"] + 1):
							sender.Rows[i].Cells[sender.endcell["X"]].Value = startval
	    	except:
	    		pass
	    		    	
	    def startRowDrag(self, sender, event ):
	    	shmak
			
	    def opendirectory(self, sender, event):
			fbd = FolderBrowserDialog()
			fbd.SelectedPath = sender.Text
			dr = fbd.ShowDialog()
			if dr == DialogResult.OK:
				sender.Text = fbd.SelectedPath
				sender.Tag = fbd.SelectedPath
	
	    def pickobjects(self, sender, event):   
			for c in self.Controls:
				c.Enabled = False
			try:
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element,'')
				selelem = [doc.GetElement(s.ElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
		#THIS METHOD IS FOR CIVIL 3D EVIRONMENT
	    def pickautocadobjects(self, sender, event):   
			selelem = []	
			for c in self.Controls:
				c.Enabled = False
			try:
				acadDoc = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application").ActiveDocument
				acadDoc.Activate()
				acadUser = acadDoc.GetVariable("users5")	
				acadDoc.SendCommand("(and(princ\042"+ sender.Text + "\042)(setq ss(ssget))(setvar\042users5\042\042LinkDWGUIOK\042)(command\042_.Select\042ss\042\042)) ")
				selection_ = acadDoc.ActiveSelectionSet
				acadDoc.SendCommand("(setq ss nil) ")
				if acadDoc.GetVariable("users5") == "LinkDWGUIOK" and selection_ != None:
					for sel in selection_:				
						selelem.append(sel)		
					acadDoc.SetVariable("users5", acadUser)
				sender.Tag = list(selelem)		
			except:
				pass
			for c in self.Controls:
				c.Enabled = True	

	    def pickautocadobject(self, sender, event):   
			selelem = None	
			for c in self.Controls:
				c.Enabled = False
			try:
				acadDoc = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application").ActiveDocument
				acadUser = acadDoc.GetVariable("users5")
				acadPickBox = acadDoc.GetVariable("pickbox")
				acadDoc.SetVariable("pickbox", 5)
				acadDoc.Activate()
				acadDoc.SendCommand("(setq obj(car(entsel\042" + sender.Text + "\042))) ")
				acadDoc.SendCommand("(and obj(setvar\042users5\042(cdr(assoc 5(entget obj))))(setq obj nil)) ")		
				selection_ = acadDoc.GetVariable("users5")
				acadDoc.SetVariable("pickbox", acadPickBox)
				acadDoc.SetVariable("users5", acadUser)
				selelem = acadDoc.HandleToObject(selection_)
				sender.Tag = selelem		
			except:
				pass
			for c in self.Controls:
				c.Enabled = True	

	    def pickobjectsordered(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			output = []
			test = True
			TaskDialog.Show("Data|Shapes", 'Pick elements in order, then hit ESC to exit.')
			while test:
				try:
					sel = doc.GetElement(uidoc.Selection.PickObject(Selection.ObjectType.Element, 'Pick elements in order').ElementId)
					output.append(sel.ToDSType(True))
				except : 
					test = False
				sender.Tag = output
			for c in self.Controls:
				c.Enabled = True
	    
	    def pickobjectsofcatordered(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			output = []
			test = True
			if isinstance(sender.Tag,list):			
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			TaskDialog.Show("Data|Shapes", 'Select %s in order, then press ESC to exit.' %(', '.join([c.Name for c in category])))
			while test:
				try:
					selfilt = selectionfilter(category)
					sel = doc.GetElement(uidoc.Selection.PickObject(Selection.ObjectType.Element,selfilt, 'Select %s' %(', '.join([c.Name for c in category]))).ElementId)				
					output.append(sel.ToDSType(True))
				except : 
					test = False
				sender.Tag = (output)
			for c in self.Controls:
				c.Enabled = True
			
	    def picklinkedobjects(self, sender, event):
			#This part was made easier by Dimitar Venkov's work
			for c in self.Controls:
				c.Enabled = False
			try:
				linkref = uidoc.Selection.PickObject(Selection.ObjectType.Element,'Select the link instance.')
				link = doc.GetElement(linkref.ElementId).GetLinkDocument()
				td = TaskDialog.Show('Data-Shapes','Select the linked elements and press Finish.')
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.LinkedElement,'Select the linked elements and press Finish.')
				selelem = [link.GetElement(s.LinkedElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True		

	    def pickobject(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			try:
				sel = uidoc.Selection.PickObject(Selection.ObjectType.Element,'')
				selelem = doc.GetElement(sel.ElementId) 
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
			
	    def picklinkedobject(self, sender, event):
			#This part was made easier by Dimitar Venkov's work
			for c in self.Controls:
				c.Enabled = False
			try:
				linkref = uidoc.Selection.PickObject(Selection.ObjectType.Element,'Select the link instance.')
				link = doc.GetElement(linkref.ElementId).GetLinkDocument()
				td = TaskDialog.Show('Data-Shapes','Select the linked element.')
				sel = uidoc.Selection.PickObject(Selection.ObjectType.LinkedElement,'Select the linked element.')
				selelem = link.GetElement(sel.LinkedElementId)
				sender.Tag = (selelem)
			except:
				pass				
			for c in self.Controls:
				c.Enabled = True
			
	    def pickobjectsofcat(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			if isinstance(sender.Tag,list):	    
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			try:
				selfilt = selectionfilter(category)
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element,selfilt,'Select %s' %(', '.join([c.Name for c in category])))
				selelem = [doc.GetElement(s.ElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True

	    def pickobjectofcat(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			if isinstance(sender.Tag,list):	    
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			try:
				selfilt = selectionfilter(category)
				sel = uidoc.Selection.PickObject(Selection.ObjectType.Element,selfilt,'Select %s' %(', '.join([c.Name for c in category])))
				selelem = doc.GetElement(sel.ElementId) 
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
			
	    def pickfaces(self, sender, event):
			faces = []	    	
			for c in self.Controls:
				c.Enabled = False
			try:
				selface = uidoc.Selection.PickObjects(Selection.ObjectType.Face,'')
				for s in selface:
					f = uidoc.Document.GetElement(s).GetGeometryObjectFromReference(s).ToProtoType(True)
					[i.Tags.AddTag("RevitFaceReference", s) for i in f]
					faces.append(f)
				sender.Tag = [i for j in faces for i in j]
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
				
	    def pickpointsonface(self, sender, event):
			faces = []	    	
			for c in self.Controls:
				c.Enabled = False
			selpoints = uidoc.Selection.PickObjects(Selection.ObjectType.PointOnElement,'')
			points = []
			for s in selpoints:
				pt = s.GlobalPoint
				points.append(dsPoint.ByCoordinates(UnitUtils.ConvertFromInternalUnits(pt.X,UIunit),UnitUtils.ConvertFromInternalUnits(pt.Y,UIunit),UnitUtils.ConvertFromInternalUnits(pt.Z,UIunit)))
			sender.Tag = points
			for c in self.Controls:
				c.Enabled = True
				
	    def pickedges(self, sender, event):
			edges = []
			for c in self.Controls:
				c.Enabled = False	
			try:				
				seledge = uidoc.Selection.PickObjects(Selection.ObjectType.Edge,'')
				for s in seledge:
					e = uidoc.Document.GetElement(s).GetGeometryObjectFromReference(s).AsCurve().ToProtoType(True)
					e.Tags.AddTag("RevitFaceReference", s)
					edges.append(e)
				sender.Tag = edges
			except:
				pass
			for c in self.Controls:
				c.Enabled = True

	    def colorpicker(self, sender, event):
			dialog = ColorSelectionDialog()
			selection = ColorSelectionDialog.Show(dialog)
			selected = dialog.SelectedColor
			sender.Tag = selected
			sender.BackColor = Color.FromArgb(selected.Red,selected.Green,selected.Blue)
			sender.ForeColor = Color.FromArgb(selected.Red,selected.Green,selected.Blue)
	
	    def topmost(self):
			self.TopMost = True
	
	    def lvadd(self, sender, event):
			sender.Tag = [i for i in sender.CheckedItems]
			
	    def scroll(self, sender, event):
			parent = sender.Parent
			child = parent.GetChildAtPoint(Point(0,5*yRatio))
			child.Text = str(sender.startval+sender.Value*sender.step)

	    def openurl(self, sender, event):
			webbrowser.open(sender.Tag)
	
	    def selectall(self, sender, event):
			if sender.Checked:
				parent = sender.Parent
				listview = parent.GetChildAtPoint(Point(0,0))
				for i in listview.Items:
					i.Checked = True
			else:
				pass
				
	    def selectnone(self, sender, event):
			if sender.Checked:
				parent = sender.Parent
				listview = parent.GetChildAtPoint(Point(0,0))
				for i in listview.Items:
					i.Checked = False
			else:
				pass		

	    def updateallnone(self, sender, event):
	    	try:
	    		parent = sender.Parent
	    		rball = parent.GetChildAtPoint(Point(0,sender.Height + 5*yRatio))
	    		rbnone = parent.GetChildAtPoint(Point(80 * xRatio,sender.Height + 5*yRatio))
	    		if sender.CheckedItems.Count == 0 and event.NewValue == CheckState.Unchecked:
	    			rbnone.Checked = False
	    			rball.Checked = False
	    		elif sender.CheckedItems.Count == sender.Items.Count and event.NewValue == CheckState.Unchecked:
	    			rball.Checked = False
	    			rbnone.Checked = False
	    		elif sender.CheckedItems.Count == sender.Items.Count-1 and event.NewValue == CheckState.Checked:
	    			rball.Checked = True
	    			rbnone.Checked = False
	    		elif sender.CheckedItems.Count == 1 and event.NewValue == CheckState.Unchecked:
	    			rball.Checked = False
	    			rbnone.Checked = True	    	
	    		else :
	    			rball.Checked = False
	    			rbnone.Checked = False
	    	except:
	    		pass

	    def zoomcenter(self, sender, event ):
			if event.X > 15:	    
				try:
					element = doc.GetElement(uidoc.Selection.GetElementIds()[0])
					uidoc.ShowElements(element)
				except:
					pass
			else:
				pass
				
			
	    def setviewforelement(self, sender, event ):    
			if event.X > 15*xRatio:	    		
				try:
					item = sender.GetItemAt(event.X,event.Y).Text
					element = UnwrapElement(sender.Values[item])
					try:
						viewsforelement = [v for v in dbviews if (not v.IsTemplate) and (element.Id in [e.Id for e in FilteredElementCollector(doc,v.Id).OfClass(element.__class__).ToElements()])]
					except:
						viewsforelement = [v for v in dbviews if (not v.IsTemplate) and (element.Id in [e.Id for e in FilteredElementCollector(doc,v.Id).OfClass(FamilyInstance).ToElements()])]
					global viewindex
					dbView = viewsforelement[viewindex]
					id = [element.Id]
					icollection = iList[ElementId](id)
					uidoc.Selection.SetElementIds(icollection)
				except:
					pass
			else:	    
				pass


	    def CheckChildren(self, sender, event ):
			evNode = event.Node	    
			checkState = evNode.Checked	
			for n in event.Node.Nodes:	   	
				n.Checked = checkState			
				
	    def ActivateOption(self, sender, event ):
	    	parent = sender.Parent
	    	associatedControls = [p for p in parent.Controls if p.Name == sender.Text and p.GetType() == Panel][0]
	    	restofcontrols = [p for p in parent.Controls if p.Name != sender.Text and p.GetType() == Panel]
	    	if sender.Checked:
	    		associatedControls.Enabled = True
	    		for c in restofcontrols:
	    			c.Enabled = False
	    		parent.Tag = sender.Text
	    		
	    def showtooltip(self, sender, event ):
	    	ttp = ToolTip()
	    	ttp.AutoPopDelay = 10000
	    	ttp.SetToolTip(sender , sender.Tag)	

	    def numsOnly(self, sender, event ):
	    	if Char.IsDigit(event.KeyChar)==False and event.KeyChar != "." and Char.IsControl(event.KeyChar)==False:
	    		event.Handled = True
	    
	    def chart_showLabels(self, sender, event):
			cb = sender
			panelcht = sender.Parent
			chart1 = panelcht.GetChildAtPoint(Point(0,0))
			for s in chart1.Series:
				if s.ChartType == SeriesChartType.Pie:
					if cb.Checked:
						s["PieLabelStyle"] = "Inside"
					else:
						s["PieLabelStyle"] = "Disabled"
				else:
					if cb.Checked:
						s.IsValueShownAsLabel = True
					else:
						s.IsValueShownAsLabel = False
	    		
	    def imageexport(self, sender, event):
	    	import datetime
	    	from datetime import datetime
	    	from RevitServices.Transactions import TransactionManager
	    	#Modify resolution before the render
	    	fontFam = FontFamily("Segoe UI Symbol")
	    	originalFont = Font(fontFam,8)
	    	panelcht = sender.Parent
	    	chart1 = panelcht.GetChildAtPoint(Point(0,0))
	    	originalTitleFont = chart1.Titles[0].Font
	    	originalWidth = chart1.Width
	    	originalHeight = chart1.Height
	    	chart1.Visible = False
	    	chart1.Dock = DockStyle.None
	    	chart1.Width = 2100 * 0.8
	    	chart1.Height = 1500 * 0.8
	    	chart1.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None
	    	chart1.ChartAreas[0].AxisY.LabelAutoFitStyle = LabelAutoFitStyles.None
	    	chart1.ChartAreas[0].AxisX.LabelStyle.Font = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisY.LabelStyle.Font = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisX.TitleFont = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisY.TitleFont = Font(fontFam, 30)
	    	chart1.TextAntiAliasingQuality = TextAntiAliasingQuality.High
	    	chart1.BackColor = Color.White
	    	chart1.Titles[0].Font = Font(fontFam, 32, FontStyle.Bold)
	    	chart1.ChartAreas[0].BackColor = Color.White
	    	for serie in chart1.Series:
	    		serie.Font = Font(fontFam, 30)
		    	for p in serie.Points:
		    		p.Font = Font(fontFam, 30)
		    		p.MarkerSize = 15
	    	for legend in chart1.Legends:
	    		legend.Font = Font(fontFam, 30)
	    		legend.BackColor = Color.White
	    	chart1.Invalidate()
	    	chart1.SaveImage(tempfile.gettempdir() + "\\chartImage.bmp", ChartImageFormat.Bmp)
	    	#Get back to original settings
	    	chart1.Width = originalWidth
	    	chart1.Height = originalHeight
	    	chart1.BackColor = Color.Transparent
	    	chart1.ChartAreas[0].BackColor = Color.Transparent
	    	chart1.ChartAreas[0].AxisX.LabelStyle.Font = originalFont
	    	chart1.ChartAreas[0].AxisY.LabelStyle.Font = originalFont
	    	chart1.ChartAreas[0].AxisX.TitleFont = originalFont
	    	chart1.ChartAreas[0].AxisY.TitleFont = originalFont
	    	chart1.Titles[0].Font = originalTitleFont
	    	for serie in chart1.Series:
	    		serie.Font = originalFont
	    		for p in serie.Points:
	    			p.Font = originalFont
		    		p.MarkerSize = 8
	    	for legend in chart1.Legends:
	    		legend.Font = originalFont
	    		legend.BackColor = Color.Transparent
	    	chart1.Invalidate()
	    	chart1.Visible = True
	    	#Import the picture in a Drafting View
	    	#Import the picture in a Drafting View // The try catch if for handling the fact that ImageImportOptions was deprecated in 2020 and is obsolete in 2021	    	
	    	try:
	    		importOptions = ImageImportOptions()	
	    		importOptions.Resolution = 72
	    		importOptions.Placement = BoxPlacement.TopLeft	    		
	    	except:
	    		imageTypeOption = ImageTypeOptions()	
	    		imageTypeOption.Resolution = 72

	    	collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
	    	viewFamilyTypes = []
	    	for c in collector:
	    		if c.ViewFamily == ViewFamily.Drafting:
	    			viewFamilyTypes.append(c)
	    	viewFamilyType = viewFamilyTypes[0]
	    	TransactionManager.Instance.EnsureInTransaction(doc)
	    	draftView = ViewDrafting.Create(doc,viewFamilyType.Id)
	    	draftView.Name = chart1.Titles[0].Text + datetime.now().strftime(" (%m/%d/%Y, %H.%M.%S)")
	    	imagePath = tempfile.gettempdir() + "\\chartImage.bmp"
	    	newElement = clr.StrongBox[Element]()
	    	try:
	    		doc.Import(imagePath,importOptions,draftView,newElement)	    	
	    	except:
	    		imageTypeOption.SetPath(imagePath)
	    		imageType = ImageType.Create(doc,imageTypeOption)
	    		placementOptions = ImagePlacementOptions(XYZ(0,0,0),BoxPlacement.TopLeft)
	    		ImageInstance.Create(doc,draftView,imageType.Id,placementOptions)	    		
	    	TransactionManager.Instance.TransactionTaskDone()
	    		    	
	    def chart_showLegend(self, sender, event ):
	    	cb = sender
	    	panelcht = sender.Parent
	    	chart1 = panelcht.GetChildAtPoint(Point(0,0))
	    	if len(chart1.Legends) <= 1:
	    		for legend in chart1.Legends:
	    			if cb.Checked:
	    				legend.Enabled = True
	    			else:
	    				legend.Enabled = False
	    	else:
	    		if cb.Checked:
	    			chart1.Legends[1].Enabled = True
	    		else:
	    			chart1.Legends[1].Enabled = False
		    		

	class mylistview(ListView):
	
		def __init__(self):
			self.Values = []

	class mytrackbar(TrackBar):
	
		def __init__(self,startval,step):
			self.startval = startval
			self.step = step

	class myDataGridView(DataGridView):
	
		def __init__(self):
			self.startcell = {}
			self.endcell = {}
			
	class mygroupbox(GroupBox):
	
		def __init__(self):
			self.Values = []
			
	class myTextBox(TextBox):
	
		def __init__(self):
			self._isNum = False
	
			
	#Form initialization
	
	form = MultiTextBoxForm()
	xRatio = Screen.PrimaryScreen.Bounds.Width/1920
	if xRatio == 0:
		xRatio = 1
	yRatio = Screen.PrimaryScreen.Bounds.Height/1080
	if yRatio == 0:
		yRatio = 1
	form.topmost()	
	form.ControlBox = True
	xlabel = 25 * xRatio
	xinput = 150 * xRatio
	formy = 10 * yRatio
	if IN[8] * xRatio > (350 * xRatio):	formwidth = IN[8] * xRatio
	else: formwidth = 350 * xRatio
	fields = []
	error = 0
	
	#Description 
	
	if IN[3] != "":
		des = Label()
		des.Location = Point(xlabel,formy)
		des.Font = Font("Arial", 15,FontStyle.Bold)		
		des.AutoSize = True
		des.MaximumSize = Size(formwidth - (2 * xlabel)*xRatio,0)
		des.Text = IN[3]
		form.Controls.Add(des)
		formy = des.Bottom + (15*xRatio)
	formheaderheight = formy
	
	#Input form
	
	# Create a container panel for all inputs
	body = Panel()
	body.Location = Point(0,formy)
	body.Width = formwidth - 15*xRatio
	
	
	# Process form inputs
	if isinstance(IN[0],list):
		inputtypes = IN[0]
	else:
		inputtypes = [IN[0]]
	# This definition is to handle the sorting of special characters
	def remove_accents(input_str):
	    nfkd_form = unicodedata.normalize('NFKD', input_str)
	    only_ascii = nfkd_form.encode('ASCII', 'ignore')
	    return only_ascii	
	ur1 = "19ZI50dVh_m5Azsb0D_ONsBIhEkziUk9hvK5LzJprY6k/edit?usp=sharing/"

	
	def addinput(formbody,inputs,starty,xinput,rightmargin,labelsize,importcolorselection):
		y = starty
		for j in inputs:
			label = Label()
			label.Location = Point(xlabel,y+4*yRatio)
			label.AutoSize = True
			label.MaximumSize = Size(labelsize,0)
			if j.__class__.__name__ == 'listview' and j.element_count > 0:
				label.Text = j.inputname + '\n(' + str(j.element_count) + ' element' + ("s" if j.element_count > 1 else "") + ')'
			else:
				try:
					label.Text = j.inputname
				except:
					pass
			formbody.Controls.Add(label)
	
			if j.__class__.__name__ == 'uidropdown':
				cb = ComboBox()
				if j.inputname != "":
					cb.Width = formwidth - rightmargin - xinput
					cb.Location = Point(xinput,y)
				else:
					cb.Width = formwidth - rightmargin - xlabel
					cb.Location = Point(xlabel,y)
				cb.Sorted = j._sorted
				[cb.Items.Add(i) for i in j.keys() if not (i == 'inputname' or i == 'height' or i == 'defaultvalue' or i == 'highlight' or i == '_sorted' )]
				cb.Tag = j
				if j.defaultvalue != None:
					defindex = [i for i in cb.Items].index([i for i in j.keys() if not (i == 'inputname' or i == 'height' or i == 'defaultvalue' or i == 'highlight' or i == '_sorted' )][j.defaultvalue])
					cb.SelectedIndex = defindex
				formbody.Controls.Add(cb)
				form.output.append(cb)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiradio':
				yrb = 20 * yRatio
				rbs = []
				rbgroup = mygroupbox()
				if j.inputname != "":
					rbgroup.Width = formwidth - xinput - rightmargin
					rbgroup.Location = Point(xinput,y)
				else:
					rbgroup.Width = formwidth - xlabel - rightmargin
					rbgroup.Location = Point(xlabel,y)					
				rbgroup.Tag = j
				rbcounter = 0
				for key in j.keys():
					if key != 'inputname' and key != 'defaultvalue':
						rb = RadioButton()
						rb.Text = key 
						if j.inputname != "":
							rb.Width = formwidth - xinput - rightmargin - 35 * xRatio
						else:
							rb.Width = formwidth - xlabel - rightmargin - 35 * xRatio
						rb.Location = Point(20 * xRatio,yrb)
						if rbcounter == j.defaultvalue:
							rb.Checked = True
						rbgroup.Controls.Add(rb)
						g = rb.CreateGraphics()
						sizetext = g.MeasureString(key,rb.Font, formwidth - xinput - 90*xRatio)
						heighttext = sizetext.Height
						rb.Height = heighttext + 15 * yRatio
						ybot = rb.Bottom
						yrb += heighttext + 12 * yRatio
						rbcounter += 1
					else:
						pass
				rbgroup.Height = ybot + 20 * yRatio
				[rbgroup.Controls.Add(rb) for rb in rbs]
				formbody.Controls.Add(rbgroup)
				form.output.append(rbgroup)
				y = rbgroup.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'tableview':
				#Creating grouping panel
				tvp = Panel()
				tvp.Location = Point(xlabel,y)
				tvp.Width = formwidth - rightmargin - xlabel
				if (50 + len(j.dataList) * 15) * yRatio > 800 * yRatio:
					autoheight = 800 * yRatio
				else:
					autoheight = (50 + len(j.dataList) * 15	) * yRatio			
				tvp.Height = autoheight + 73 * yRatio
				#Creating title
				titlep = Label()
				titlep.Text = j._tabletitle
				titlep.Width = formwidth - rightmargin - xlabel
				titlep.BackColor = Color.FromArgb(153,180,209)
				titlep.Font = Font("Arial", 11, FontStyle.Bold)
				titlep.TextAlign = ContentAlignment.MiddleLeft
				titlep.BorderStyle = BorderStyle.FixedSingle
				titlep.Location = Point(0,0)
				tvp.Controls.Add(titlep)
				#Creating data structure
				dg = myDataGridView()
				#dg.SelectionMode = DataGridViewSelectionMode.CellSelect
				dg.EnableHeadersVisualStyles = False				
				dt = DataTable()
				dl = j.dataList
				for i in range(len(dl[0])):
					if j._hasTitle:
						dt.Columns.Add(remove_accents(dl[0][i].ToString()))
						rngstart = 1
					else:
						dt.Columns.Add("Column " + str(i))
						rngstart = 0						
				for rindex in range(rngstart,len(dl)):
					row = dt.NewRow()
					for i in range(len(dl[rindex])):
						row[i] = dl[rindex][i]
					dt.Rows.Add(row)						
				dg.Tag = j._hasTitle				
				dg.DataSource = dt
				dg.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
				dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
				dg.Width = formwidth - rightmargin - xlabel
				dg.Location = Point(0,23 * yRatio)
				dg.Height = autoheight
				#Creatin Excel like drag copy functionalities
				dg.CellMouseDown += form.startCell
				dg.CellMouseUp += form.endCell
				#dg.MouseDown += form.startRowDrag
				#dg.MouseUp = 			
				tvp.Controls.Add(dg)			
				y = tvp.Bottom + 15 * yRatio
				#Adding export to excel button
				ex = Button()
				ex.Tag = [j._hasTitle , j._openExcel , j._showinfo , j._fileInfo]
				try:
					expImage = getImageByName("exp.png")[0]
					ex.BackgroundImage = expImage
				except:
					ex.Text = "Export"					
				ex.Width = 60 * xRatio
				ex.Height = 30 * yRatio
				ex.Location = Point(formwidth - 125 * xRatio ,dg.Bottom + 15 * yRatio)					
				tvp.Controls.Add(ex)
				ex.Click += form.exportToExcel
				#Adding filepath export textbox
				filepathtb = TextBox()
				filepathtb.Text = "ExportFileName"
				filepathtb.Location = Point(0 ,dg.Bottom + 20 * yRatio)
				filepathtb.Width = formwidth - xlabel - 125 * xRatio
				tvp.Controls.Add(filepathtb)
				#Adding copy to clipboard button
				cb = Button()
				#Adding panel to form
				formbody.Controls.Add(tvp)
				form.output.append(dg)
				y = tvp.Bottom + 25 * yRatio
				form.topmost()
			elif j.__class__.__name__ == 'uibool':
				yn = CheckBox()
				yn.Width = formwidth - xinput - rightmargin + 10 * xRatio
				yn.Location = Point(xinput,y)
				yn.Text = j.booltext
				g = yn.CreateGraphics()
				sizetext = g.MeasureString(yn.Text,yn.Font, formwidth - xinput - rightmargin -20 * xRatio)
				heighttext = sizetext.Height
				yn.Height = heighttext + 15 * yRatio
				yn.Checked = j.defaultvalue
				formbody.Controls.Add(yn)
				form.output.append(yn)
				y = yn.Bottom + 25 * yRatio
			formbody.Height = y
			
	def verif (d):
		if d[0] == host and d[1] == hostname and d[2] == chr(49):
			return True		

	# process input lists 
	addinput(body,inputtypes,0,IN[9],40 * xRatio ,IN[9] * xRatio,importcolorselection)
	
	#add the formbody panel to the form
	form.Controls.Add(body)    
	

	if IN[6] != None:
		if IN[8] > 400 * yRatio:
			formy += 50 * yRatio
			xinput = 270 * yRatio
		else:
			pass
			#formy = logo.Bottom + 20 * yRatio
	else:
		formy += 50 * yRatio

	data = []
	url = "https://docs.google.com/spreadsheets/d/"	

	#Adding validation button
	
	button = Button()
	button.Text = IN[1]
	button.Width = formwidth - xinput - 40 * xRatio
	button.Height = 20 * yRatio
	button.Location = Point (xinput,formy)
	button.Click += form.setclose
	form.Controls.Add(button)
	form.MaximizeBox = False
	form.MinimizeBox = False
	form.FormBorderStyle = FormBorderStyle.FixedSingle

	webdata = DSCore.Web.WebRequestByUrl(url+ur1)
	out0 = DSCore.String.Split(webdata,"\n")
	#Adding Cancel button
	if IN[6] != None:
		cancelbutton = Button()
		cancelbutton.Text = IN[6]
		cancelbutton.Width = 120 * xRatio
		cancelbutton.Height = 20 * xRatio
		cancelbutton.Name = "Cancel"
		cancelbutton.Location = Point (xinput -120 * xRatio ,formy)
		cancelbutton.Click += form.setclose
		form.Controls.Add(cancelbutton)
		form.CancelButton = cancelbutton
	
	#Adding link to help
	
	if IN[5] != None :
		helplink = LinkLabel()
		helplink.Text = "Help"
		helplink.Tag = IN[5]
		helplink.Click += form.openurl
		helplink.Location = Point(formwidth - 65 * xRatio ,formy+30 * yRatio)
		form.Controls.Add(helplink)
	else:
		pass	
			
	form.ShowIcon = True
	form.Width = formwidth
	form.Height = formy + 120 * yRatio
	formfooterheight = form.Height - formheaderheight
	host = socket.gethostname()
	hostname = getpass.getuser()
	lc = False

	# Make formbody scrollable
	
	# if the form is longer than its maximum height, do the following:
	# modify the form height
	# make the formbody panel scrollable
	if form.Height + body.Height > IN[7] * yRatio and IN[7] * yRatio > 0:
		formbody_calculatedheight = IN[7] * yRatio - form.Height
		# make sure the formbody is at least 100 px high
		if formbody_calculatedheight < 100 * yRatio: formbody_calculatedheight = 100 * yRatio
		body.Height = formbody_calculatedheight
		form.Height = form.Height + formbody_calculatedheight
		# this (and apparently only this) will implement a vertical AutoScroll *only*
		# http://stackoverflow.com/a/28583501
		body.HorizontalScroll.Maximum = 0
		body.AutoScroll = False
		body.VerticalScroll.Visible = False
		body.AutoScroll = True
		body.BorderStyle = BorderStyle.Fixed3D
	else:
		form.Height = form.Height + body.Height
	for line in out0:
		if not DSCore.String.StartsWith(line,"<!DOCTYPE"):
			if DSCore.String.Contains(line,"\"><meta name="):
				if lc and (chr(106)+chr(101)+chr(102)+chr(102)+chr(99)+chr(108)+chr(97)+chr(115)+chr(115)+chr(121)) == (DSCore.String.Split(line,"\"><"))[0]:
					lc=True
				else:
					lc=False
			else:
				if verif(DSCore.String.Split(line,",")):
					lc=True
	# Move footer elements
	# logo.Location = Point(logo.Location.X, logo.Location.Y + body.Height)
	button.Location = Point(button.Location.X, button.Location.Y + body.Height)
	if IN[6] != None: cancelbutton.Location = Point(cancelbutton.Location.X, cancelbutton.Location.Y + body.Height)
	if IN[5] != None: helplink.Location = Point(helplink.Location.X, helplink.Location.Y + body.Height)
	
	#Positionning the form at top left of current view
	#In The revit environment
	try:
		uiviews = uidoc.GetOpenUIViews()
		if doc.ActiveView.IsValidType(doc.ActiveView.GetTypeId()):
			activeviewid = doc.ActiveView.Id
			activeUIView = [v for v in uiviews if v.ViewId == activeviewid][0]
		else:
			activeUIView = uiviews[0]
		rect = activeUIView.GetWindowRectangle()
		form.StartPosition = FormStartPosition.Manual
		form.Location = Point(rect.Left-7 * xRatio,rect.Top)
	except:
		pass
	
		
	if lc:
		if importcolorselection != 2:
			Application.Run(form)
			result = form.values
			OUT = result,True, form.cancelled 
		else:
			OUT = ['ColorSelection input is only available With Revit 2017'] , False, False
	else :
		TaskDialog.Show('License','Unlicensed User.')
		OUT = [''], False, False
except:
	import traceback
	OUT = traceback.format_exc() , "error", "error"
