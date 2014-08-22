import sys
reload(sys)
sys.setdefaultencoding("latin-1")

import xlrd, arcpy, string, os, zipfile, fileinput, time
from datetime import date
start_time = time.time()

INPUT_PATH = "input"
OUTPUT_PATH = "output"
if arcpy.Exists(OUTPUT_PATH + "\\DrinkingWater.gdb"):
	os.system("rmdir " + OUTPUT_PATH + "\\DrinkingWater.gdb /s /q")
os.system("del " + OUTPUT_PATH + "\\*DrinkingWater*.*")
arcpy.CreateFileGDB_management(OUTPUT_PATH, "DrinkingWater", "9.3")
arcpy.env.workspace = OUTPUT_PATH + "\\DrinkingWater.gdb"

import math

# http://stackoverflow.com/questions/343865/how-to-convert-from-utm-to-latlng-in-python-or-javascript
def utmToLatLng(zone, easting, northing, northernHemisphere=True):
	if not northernHemisphere:
		northing = 10000000 - northing

	a = 6378137
	e = 0.081819191
	e1sq = 0.006739497
	k0 = 0.9996
		
	arc = northing / k0
	mu = arc / (a * (1 - math.pow(e, 2) / 4.0 - 3 * math.pow(e, 4) / 64.0 - 5 * math.pow(e, 6) / 256.0))

	ei = (1 - math.pow((1 - e * e), (1 / 2.0))) / (1 + math.pow((1 - e * e), (1 / 2.0)))

	ca = 3 * ei / 2 - 27 * math.pow(ei, 3) / 32.0

	cb = 21 * math.pow(ei, 2) / 16 - 55 * math.pow(ei, 4) / 32
	cc = 151 * math.pow(ei, 3) / 96
	cd = 1097 * math.pow(ei, 4) / 512
	phi1 = mu + ca * math.sin(2 * mu) + cb * math.sin(4 * mu) + cc * math.sin(6 * mu) + cd * math.sin(8 * mu)

	n0 = a / math.pow((1 - math.pow((e * math.sin(phi1)), 2)), (1 / 2.0))

	r0 = a * (1 - e * e) / math.pow((1 - math.pow((e * math.sin(phi1)), 2)), (3 / 2.0))
	fact1 = n0 * math.tan(phi1) / r0

	_a1 = 500000 - easting
	dd0 = _a1 / (n0 * k0)
	fact2 = dd0 * dd0 / 2

	t0 = math.pow(math.tan(phi1), 2)
	Q0 = e1sq * math.pow(math.cos(phi1), 2)
	fact3 = (5 + 3 * t0 + 10 * Q0 - 4 * Q0 * Q0 - 9 * e1sq) * math.pow(dd0, 4) / 24

	fact4 = (61 + 90 * t0 + 298 * Q0 + 45 * t0 * t0 - 252 * e1sq - 3 * Q0 * Q0) * math.pow(dd0, 6) / 720

	lof1 = _a1 / (n0 * k0)
	lof2 = (1 + 2 * t0 + Q0) * math.pow(dd0, 3) / 6.0
	lof3 = (5 - 2 * Q0 + 28 * t0 - 3 * math.pow(Q0, 2) + 8 * e1sq + 24 * math.pow(t0, 2)) * math.pow(dd0, 5) / 120
	_a2 = (lof1 - lof2 + lof3) / math.cos(phi1)
	_a3 = _a2 * 180 / math.pi

	latitude = 180 * (phi1 - fact1 * (fact2 + fact3 + fact4)) / math.pi

	if not northernHemisphere:
		latitude = -latitude

	longitude = ((zone > 0) and (6 * zone - 183.0) or 3.0) - _a3

	return (latitude, longitude)

def createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields):
	print "Create " + featureName + " feature class"
	featureNameNAD83 = featureName + "_NAD83"
	featureNameNAD83Path = arcpy.env.workspace + "\\"  + featureNameNAD83
	arcpy.CreateFeatureclass_management(arcpy.env.workspace, featureNameNAD83, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
	# Process: Define Projection
	arcpy.DefineProjection_management(featureNameNAD83Path, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	# Process: Add Fields	
	for featrueField in featureFieldList:
		arcpy.AddField_management(featureNameNAD83Path, featrueField[0], featrueField[1], featrueField[2], featrueField[3], featrueField[4], featrueField[5], featrueField[6], featrueField[7], featrueField[8])
	# Process: Append the records
	cntr = 1
	try:
		with arcpy.da.InsertCursor(featureNameNAD83, featureInsertCursorFields) as cur:
			for rowValue in featureData:
				cur.insertRow(rowValue)
				cntr = cntr + 1
	except Exception as e:
		print "\tError: " + featureName + ": " + e.message
	# Change the projection to web mercator
	arcpy.Project_management(featureNameNAD83Path, arcpy.env.workspace + "\\" + featureName, "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	arcpy.Delete_management(featureNameNAD83Path, "FeatureClass")
	print "Finish " + featureName + " feature class."

TreatmentProcessesDict = {}
wb = xlrd.open_workbook('input\\Data\\DW Map File - draft2.1.xlsx')
sh = wb.sheet_by_name(u'Treatment Processes-2')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in TreatmentProcessesDict:
		TreatmentProcessesDict[row[0]].append(row[1])
	else:
		TreatmentProcessesDict[row[0]] = [row[1]]
print len(TreatmentProcessesDict)

SourcesDict = {}
wb = xlrd.open_workbook('input\\Data\\DW Map File - draft2.1.xlsx')
sh = wb.sheet_by_name(u'Sources-3')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in SourcesDict:
		SourcesDict[row[0]].append(row[1])
	else:
		SourcesDict[row[0]] = [row[1]]
print len(SourcesDict)

ReceivingDWSDict = {}
wb = xlrd.open_workbook('input\\Data\\DW Map File - draft2.1.xlsx')
sh = wb.sheet_by_name(u'Receiving DWS-4')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in ReceivingDWSDict:
		ReceivingDWSDict[row[0]].append(row[2])
	else:
		ReceivingDWSDict[row[0]] = [row[2]]
print len(ReceivingDWSDict)

IRRDict = {}
wb = xlrd.open_workbook('input\\Data\\DWO_IRR_Loading_File_-_By_Fiscal_Year2012-2013_prototype_map_aug142014.xlsx')
sh = wb.sheet_by_name(u'Page1_1')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[1], wb.datemode)
	monthStr = str(month)
	if len(monthStr) == 1:
		monthStr = "0" + monthStr
	dayStr = str(day)
	if len(dayStr) == 1:
		dayStr = "0" + dayStr
	row[1] = str(year) + "/" + monthStr + "/" + dayStr
	row[0] = unicode(int(row[0]))
	row[3] = "{0:.2f}%".format(row[3] * 100)
	IRRDict[row[0]] = [row[1], row[2], row[3], row[4], row[5], row[6]]
print len(IRRDict)

DWQDict = {}
wb = xlrd.open_workbook('input\\Data\\DWO_DWQ_2012-13.xlsx')
sh = wb.sheet_by_name(u'Page1-1')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	row[0] = unicode(int(row[0]))
	row[1] = "{0:.2f}%".format(row[1] * 100)
	DWQDict[row[0]] = [row[1], row[2], row[3]]
print len(DWQDict)

DWSPDict = {}
featureName = "DWSP"
wb = xlrd.open_workbook('input\\Data\\DWSP_SelectedParameters_15-Aug-2014.xlsx')
sh = wb.sheet_by_name(u'DWSP Data')
featureData = []
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)

	year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[7], wb.datemode)
	monthStr = str(month)
	if len(monthStr) == 1:
		monthStr = "0" + monthStr
	dayStr = str(day)
	if len(dayStr) == 1:
		dayStr = "0" + dayStr
	row[7] = str(year) + "/" + monthStr + "/" + dayStr
	if (not(row[2] in DWSPDict)):
		DWSPDict[row[2]] = [0, 0, 0, 0]
	if ((row[9] == "2-METHYLISOBORNEOL") or (row[9] == "GEOSMIN")):
		DWSPDict[row[2]][0] = DWSPDict[row[2]][0] + 1
	elif (row[9] == "CHLORIDE"):
		DWSPDict[row[2]][1] = DWSPDict[row[2]][1] + 1
	elif (row[9] == "COLOUR; TRUE"):
		DWSPDict[row[2]][2] = DWSPDict[row[2]][2] + 1
	elif ((row[9] == "ANATOXIN-A") or (row[9] == "MICROCYSTIN-LR") or (row[9] == "MICROCYSTIN-RR") or (row[9] == "MICROCYSTIN-LA") or (row[9] == "MICROCYSTIN-YR")):
		DWSPDict[row[2]][3] = DWSPDict[row[2]][3] + 1
	else:
		print "error"
	featureData.append([(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14]])
print "len(DWSPDict): "
print len(DWSPDict)

featureFieldList = [["SAMPLE_PROGRAM", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["DWS_NAME", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["DWS_NUMBER", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["SAMPLE_TYPE", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["SAMPLE_LOCATION", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["STATION_NUMBER", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["SAMPLE_CONDITION", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["SAMPLE_DATE", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["PARAMETER_GROUP", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["PARAMETER", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["CURRENT_DETECTION_LIMIT", "DOUBLE", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["DETECTION_LIMIT_UNIT", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["RESULT", "DOUBLE", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["RESULT_UNIT", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""], ["QUALIFIER", "TEXT", "", "", "", "", "NON_NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "SAMPLE_PROGRAM", "DWS_NAME", "DWS_NUMBER", "SAMPLE_TYPE", "SAMPLE_LOCATION", "STATION_NUMBER", "SAMPLE_CONDITION", "SAMPLE_DATE", "PARAMETER_GROUP", "PARAMETER", "CURRENT_DETECTION_LIMIT", "DETECTION_LIMIT_UNIT", "RESULT", "RESULT_UNIT", "QUALIFIER")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

featureName = "DWS"
wb = xlrd.open_workbook('input\\Data\\DW Map File - draft2.1.xlsx')
sh = wb.sheet_by_name(u'Base Profile-1')
featureData = []
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if len(str(row[23]).strip()) == 0:
		row[23] = 0.0
	if len(str(row[24]).strip()) == 0:
		row[24] = 0.0
	latitude = 0.0
	longitude = 0.0
	if (len(str(row[22]).strip()) != 0) and (len(str(row[23]).strip()) != 0) and (len(str(row[24]).strip()) != 0):
		latlng = utmToLatLng(int(row[22]), int(row[23]), int(row[24]))
		latitude = latlng[0]
		longitude = latlng[1]
	row[0] = unicode(int(row[0]))
	TreatmentProcesses = ""
	if row[0] in TreatmentProcessesDict:
		TreatmentProcesses = ", ".join(TreatmentProcessesDict[unicode(row[0])])
	Sources = ""
	if row[0] in SourcesDict:
		Sources = ", ".join(SourcesDict[unicode(row[0])])
	ReceivingDWS = ""
	if row[0] in ReceivingDWSDict:
		ReceivingDWS = ", ".join(ReceivingDWSDict[unicode(row[0])])

	IRR_row = ["", "", "", "", "", ""]
	if row[0] in IRRDict:
		IRR_row = IRRDict[row[0]]

	DWQ_row = ["", "", ""]
	if row[0] in DWQDict:
		DWQ_row = DWQDict[row[0]]

	DWSP_row = [0, 0, 0, 0]
	DWSP = [0]
	if row[0] in DWSPDict:
		DWSP_row = DWSPDict[row[0]]
		total = 0
		for elem in DWSP_row:
			total = total + elem
		DWSP = [total]
		
	featureData.append([(longitude, latitude)] + row + [latitude, longitude, TreatmentProcesses, Sources, ReceivingDWS] + IRR_row + DWQ_row + DWSP_row + DWSP)

featureFieldList = [["DWS_NUM", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", ""], ["DWS_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["OWNER_LEGAL_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["OPERATING_AUTHORITY_LEGAL_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DWS_CATEGORY", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["POPULATION_SERVED", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DESIGN_RATED_CAPACITY", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CAPACITYUOM", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["REGION_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DISTRICT_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MUNICIPALITY_NAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MUNICIPALITY_ID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MUNICIPALITY_HOME_URL", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MUNICIPALITY_PHONE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MUNICIPALITY_EMAIL", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LASTARYEAR", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LASTARURL", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ARLIBRARYURL", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MAP_DATUM", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["GEO_REFENCING_METHOD", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ACCURACY_ESTIMATES", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LOCATION_REFERENCE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTM_ZONE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTM_EASTING", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTM_NORTHING", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NUMBER_OF_DWS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LATITUDE", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LONGITUDE", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["TREATMENT_PROCESSES", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""], ["SOURCES", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""], ["RECEIVING_DWS", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""], ["DATE_OF_INSPECTION", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["INSPECTION_ID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SCORE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ENGLISH_DATE_RANGE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["FRENCH_DATE_RANGE", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["KEY", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PERCENTAGE_COMPLIED", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ENGLISH_TIME_PERIOD", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["FRENCH_TIME_PERIOD", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["TASTE_AND_ODOUR", "INTEGER", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CHLORIDE", "INTEGER", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["COLOUR", "INTEGER", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ALGAL_TOXINS", "INTEGER", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DWSP", "INTEGER", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "DWS_NUM", "DWS_NAME", "OWNER_LEGAL_NAME", "OPERATING_AUTHORITY_LEGAL_NAME", "DWS_CATEGORY", "POPULATION_SERVED", "DESIGN_RATED_CAPACITY", "CAPACITYUOM", "REGION_NAME", "DISTRICT_NAME", "MUNICIPALITY_NAME", "MUNICIPALITY_ID", "MUNICIPALITY_HOME_URL", "MUNICIPALITY_PHONE", "MUNICIPALITY_EMAIL", "LASTARYEAR", "LASTARURL", "ARLIBRARYURL", "MAP_DATUM", "GEO_REFENCING_METHOD", "ACCURACY_ESTIMATES", "LOCATION_REFERENCE", "UTM_ZONE", "UTM_EASTING", "UTM_NORTHING", "NUMBER_OF_DWS", "LATITUDE", "LONGITUDE", "TREATMENT_PROCESSES", "SOURCES", "RECEIVING_DWS", "DATE_OF_INSPECTION", "INSPECTION_ID", "SCORE", "ENGLISH_DATE_RANGE", "FRENCH_DATE_RANGE", "KEY", "PERCENTAGE_COMPLIED", "ENGLISH_TIME_PERIOD", "FRENCH_TIME_PERIOD", "TASTE_AND_ODOUR", "CHLORIDE", "COLOUR", "ALGAL_TOXINS", "DWSP")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)



elapsed_time = time.time() - start_time
print elapsed_time