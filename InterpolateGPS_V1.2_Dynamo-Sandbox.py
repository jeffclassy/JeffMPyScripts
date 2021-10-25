import clr
import math
#import sys

clr.AddReference('ProtoGeometry')
clr.AddReference('DSCoreNodes')
clr.AddReference('DSOffice')

import DSCore
from DSCore import *
import DSOffice
from DSOffice import *
from Autodesk.DesignScript.Geometry import *

#Assign the intput 
file1 = IN[0]
sheetName1 = IN[1]
exp = IN[2]
degree = IN[3]
div = IN[4]
interval = IN[5]
one_excel_second = 0.0000115740740740741
splitsecond = interval*one_excel_second/div
timestamp = []

#globalname
tim = 0.0
lati = 0.0
longi = 0.0
eleve = 0.0

#Rounddown
def rounddown (num,pow):
	mod = 1/Math.Pow(10,pow)
	round = Math.Round(num/mod)*mod
	if round<=num:
		return round
	else:
		return round - mod
		

#Import Excel file 
importedData1 = Excel.ReadFromFile(file1, sheetName1, False)
data_in = List.Transpose(List.RestOfItems(importedData1))

#manipulate coordinates
adj1 = rounddown(min(data_in[0]),exp-2)
adj2 = rounddown(min(data_in[1]),exp-2)
adj3 = Math.Pow(10,exp)
latitude_adjusted = []
longitude_adjusted = []
elevation = []
Points_Original = []
for la in data_in[0]:
	latitude_adjusted.append((la-adj1)*adj3)
for lo in data_in[1]:
	longitude_adjusted.append((lo-adj2)*adj3)
for e in data_in[2]:
	elevation.append(e)
for la,lo,e in zip(latitude_adjusted,longitude_adjusted,elevation):
	Points_Original.append(Point.ByCoordinates(la,lo,e))
Points_Original = Point.PruneDuplicates(Points_Original)

#bspline
nurbscrv = NurbsCurve.ByPoints(Points_Original,degree)
nurbscrv_curves = Curve.SplitByPoints(nurbscrv,Points_Original)
bspline = PolyCurve.ByJoinedCurves(List.TakeItems(nurbscrv_curves,nurbscrv_curves.Count-1))

#interpolate points
param_inv = range(div*elevation.Count)
Points_Interpolated=[]
t=0.0
param = 0.0
data_out = []

for p in param_inv:
	param = p*1.0/(div*elevation.Count)
	pnt = Curve.PointAtParameter(bspline,param)
	Points_Interpolated.append(pnt)
	t=t+splitsecond 
	timestamp.append(t)
	x = (pnt.X/adj3)+adj1
	y = (pnt.Y/adj3)+adj2
	data_out.append([x,y,pnt.Z,t])

data_out=List.AddItemToFront(["Latitude","Longitude","Elevation","Timestamp","Time","Date"],data_out)
#control which geometry to be visible
visible = [bspline if IN[6] else 0, Points_Original if IN[7] else 0, Points_Interpolated if IN[8] else 0]
	
OUT = data_out if IN[9] else 0, visible
