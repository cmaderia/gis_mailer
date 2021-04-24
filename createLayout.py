# Create Layout
#createLayout.py - zoom to adjacents layer, setup and send to layout view / assign case number, save (mxd and shapefiles) to folder with case number


# IMPORT

# Import libraries
import arcpy
import comtypes
import comtypes.client
import os
import sys
from win32com.shell import shell, shellcon
import textwrap
import win32gui
import pythonaddins
import shutil

# Import ArcObjects modules
def GetLibPath():
	##return "C:/Program Files/ArcGIS/com/"
	import _winreg
	keyESRI = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, \
	"SOFTWARE\\ESRI\\ArcGIS"
	)
	return _winreg.QueryValueEx(keyESRI, "InstallDir")[0] + "com\\"
	
def GetModule(sModuleName):
	import comtypes
	from comtypes.client import GetModule
	sLibPath = GetLibPath()
	GetModule(sLibPath + sModuleName)
	GetModule("esriGeometry.olb"
	)

if arcpy.Exists(r'C:\Program Files (x86)')==True:
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriGeometry.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriGeoDatabase.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesRaster.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriSystem.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriMaplex.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriCarto.olb')
    #comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriCartoUI.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDisplay.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesGDB.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesFile.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriOutput.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriFramework.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriArcMapUI.olb')
    comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriArcCatalogUI.olb')
else:
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriGeometry.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriGeoDatabase.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriDataSourcesRaster.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriSystem.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriMaplex.olb')   
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriCarto.olb')
    #comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriCartoUI.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriDisplay.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriDataSourcesGDB.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriDataSourcesFile.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriOutput.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriFramework.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriArcMapUI.olb')
    comtypes.client.GetModule(r'C:\Program Files\ArcGIS\Desktop10.2\com\esriArcCatalogUI.olb')


# METHODS A (ArcObjects)

# Methods to run ArcObjects (used for Clear selection, Set symbology, and Labeling)
# courtesy of http://www.pierssen.com/arcgis10/upload/python/snippets102.py

# Used to create a new object using a defined interface
def NewObj(MyClass, MyInterface):
  # Creates a new comtypes POINTER object where
  # MyClass is the class to be instantiated,
  # MyInterface is the interface to be assigned
  from comtypes.client import CreateObject
  try:
      ptr = CreateObject(MyClass, interface=MyInterface)
      return ptr
  except:
      return None

# Used to cast objects to a different interface (ex., pDoc cast to IMxDocument = pMxDoc)
def CType(obj, interface):
  # Casts obj to interface and returns comtypes POINTER or None
  try:
      newobj = obj.QueryInterface(interface)
      return newobj
  except:
      return None
      
def CLSID(MyClass):
    # Return CLSID of MyClass as string
    return str(MyClass._reg_clsid_)

# Get current ArcMap session     (revised so it gets the TOP ArcMap window session only)
def GetApp(app="ArcMap"):
    # Enumerate windows - code derived from http://www.brunningonline.net/simon/blog/archives/000652.html
    # Get list of all window handles
    def windowEnumerationHandlerHandles(hwnd, whandles):
        #Pass to win32gui.EnumWindows() to generate list of window handle, window text tuples.
        whandles.append(hwnd)    # append the window handle to a list
    #We can pass this, along a list to hold the results, into win32gui.EnumWindows(), as so:
    whandleList = []
    win32gui.EnumWindows(windowEnumerationHandlerHandles, whandleList)
    
    # Get list of all window names    
    def windowEnumerationHandlerNames(hwnd, wnames):
        wnames.append(win32gui.GetWindowText(hwnd))    # append the window name to a list
    wnameList = []
    win32gui.EnumWindows(windowEnumerationHandlerNames, wnameList)
    
    # Get list of all window names containing "ArcMap"    
    wArcMap = []
    for i in wnameList:
        if i.count("ArcMap") > 0:
            wArcMap.append(i)
  
    # In a standalone script, retrieves the first app session found.
    # app must be 'ArcMap' (default) or 'ArcCatalog'
    # Execute GetDesktopModules() first
    if not (app == "ArcMap" or app == "ArcCatalog"):
        print "app must be 'ArcMap' or 'ArcCatalog'"
        return None 
    import comtypes.gen.esriFramework as esriFramework
    import comtypes.gen.esriArcMapUI as esriArcMapUI
    import comtypes.gen.esriCatalogUI as esriCatalogUI
    pAppROT = NewObj(esriFramework.AppROT, esriFramework.IAppROT)
    iCount = pAppROT.Count
    if iCount == 0:
        return None
    for i in range(iCount):
        if pAppROT.Item(i).Caption == wArcMap[0]:    # if pAppROT (app session) is the ArcMap window on top, then use it
            pApp = pAppROT.Item(i)
            if app == "ArcCatalog":
                if CType(pApp, esriCatalogUI.IGxApplication):
                    return pApp
                continue
            if CType(pApp, esriArcMapUI.IMxApplication):
                return pApp



# VARIABLES

# Global variables:
tax_parcels = "Tax Parcels - Mailer"

# Initialize variables
Site_shp = ""                                                                  # Site parcel shapefile
Site_add = ""                                                                  # new layer shapefile with new Site parcels to be added 
Adjacents_shp = ""                                                             # Adjacent parcel shapefile
Adjacents_add = ""                                                             # new Adjacent parcel shapefile with new Adjacent parcels to be added
Adjacents_sort = ""                                                            # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
Adjacents_sort2 = ""                                                           # new Adjacent parcel shapefile with new Adjacent parcels sorted by Key Label (ascending)
Condos_add = ""                                                                # add all condos for each master condo GPIN

# ArcObjects variables to get map document
pApp = GetApp()
pDoc = pApp.Document
pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
pMap = pMxDoc.FocusMap

# Get access to the current mxd 
mxd = arcpy.mapping.MapDocument("CURRENT") 
# Grab the dataframe object you want
df = arcpy.mapping.ListDataFrames(mxd,"*")[0]


# METHODS B (Processes to create map)

# Clear selection and Refresh map
def clearRefresh():
    for i in range(0, pMap.LayerCount-1):
        pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
        pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
        pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
        pLayerSite = pMap.Layer(i)
        pUnkFeatureLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.FeatureLayer))
        pUnkFeatureLayer = pLayerSite
        pFeatureLayer = CType(pUnkFeatureLayer, comtypes.gen.esriCarto.IFeatureLayer) 
    
        pFeatureSelection = CType(pFeatureLayer, comtypes.gen.esriCarto.IFeatureSelection)
        pFeatureSelection.Clear()
    pMxDoc.ActiveView.Refresh()  
  
# Zoom and center map (based on extent of Adjacent parcels)
def zoomAndCenter():
    if arcpy.Exists("Adjacent Parcel") == True or arcpy.Exists("Adjacent Parcels") == True:
        for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
            if lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels":
                buffSize = 0.03*(((df.extent.XMax-df.extent.XMin)+(df.extent.YMax-df.extent.YMin))/2)
                buffSizeString=str(buffSize)+" Feet"
                arcpy.Buffer_analysis(str(lyr.name),"in_memory/layout_buffer",buffSizeString,"FULL","ROUND","NONE","#")
                bufferZoom = arcpy.mapping.Layer("in_memory/layout_buffer")
                ext = bufferZoom.getExtent()
                #ext.XMin = ext.XMin+(0.35*(ext.XMax-ext.XMin))
		#shift it to the left a little
		ext.XMin = ext.XMin + 0.1*(df.extent.XMax-df.extent.XMin)
                df.extent = ext
    
                # create centroid for Adjacents and shift extent slightly left (10% of current width after panning to centroid)
                #arcpy.Dissolve_management(str(lyr.name),"in_memory/adjacents_dissolve","#","#","SINGLE_PART","DISSOLVE_LINES")
                #arcpy.FeatureToPoint_management("in_memory/adjacents_dissolve","in_memory/adjacents_centroid")
                #adjacentsCenter = arcpy.mapping.Layer("in_memory/adjacents_centroid")
                #center_ext = adjacentsCenter.getExtent()
                #center_ext.XMin = center_ext.XMin+(0.5*(center_ext.XMax-center_ext.XMin))	    
                #center_ext.YMax = center_ext.YMax-(0.001*(center_ext.YMax-center_ext.YMin))
                #df.panToExtent(center_ext)
  
# Key Numbering (create fields for labels) - Adjacents (Adjacent Parcels) layer will be used for labels
def createLabelFields():  
    layerSrc = ""  
    if arcpy.Exists(Adjacents_shp):
        layerSrc = Adjacents_shp
    else:
        for lyr in arcpy.mapping.ListLayers(mxd):
            if lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels":
                layerSrc = lyr
        
    # add and calculate Latitude field
    arcpy.AddField_management(layerSrc,"Latitude","DOUBLE","#","#","#","#","NULLABLE","NON_REQUIRED","#")
    arcpy.CalculateField_management(layerSrc,"Latitude","!SHAPE.extent.YMax!","PYTHON_9.3","#")
    
    # sort by Latitude (highest to lowest)
    arcpy.Sort_management(layerSrc, Adjacents_sort, [["Latitude", "DESCENDING"]])

    # add and calculate KeyLabel field (1, 2, 3, etc. in order from N to S) in "Adjacent_sort"
    arcpy.AddField_management(Adjacents_sort,"KeyLabel","SHORT","#","#","#","#","NULLABLE","NON_REQUIRED","#")
    arcpy.CalculateField_management(Adjacents_sort,"KeyLabel","!FID!+1","PYTHON_9.3","#")
    
    # sort by Key Label (lowest to highest)
    arcpy.Sort_management(Adjacents_sort, Adjacents_sort2, [["KeyLabel", "ASCENDING"]])

    # delete Adjacents and Adjacents_sort shapefiles and rename Adjacents_sort2 to Adjacents (before adding)
    arcpy.Delete_management(layerSrc)
    arcpy.Delete_management(Adjacents_sort)    
    arcpy.CopyFeatures_management(Adjacents_sort2, layerSrc)
    arcpy.Delete_management(Adjacents_sort2)
	
    #else:
    #    pythonaddins.MessageBox("Please add the Adjacent Parcels layer.", "Error")
    
    # get list of GPINs to set KeyLabel values and for GetCount (count of GPINs)
    layersToGet = getMasterCondoLayer("adjacent")
    masterGPINlist = getMasterCondo(layersToGet, master=False)
    list1 = masterGPINlist[0]    # ALL GPINs
    list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed
	    
    # make KeyLabel numbering the same for each GPIN (used for labels), so only main parcels are labeled; used in conjunction with "LabelMain" field in removeDuplicates()
    # use list 2 (duplicates removed)
    for lyr in arcpy.mapping.ListLayers(mxd):
        if lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels": 
            for i in range(0, len(list2)):
                arcpy.SelectLayerByAttribute_management(lyr,"NEW_SELECTION", "\"GPIN\" = '"+list2[i]+"'")
                with arcpy.da.UpdateCursor(lyr, ('KeyLabel')) as cursor:
                    for row in cursor:
                        row[0] = i + 1 
                        cursor.updateRow(row)
    

# Enable the Maplex labeling engine	
def enableMaplex():
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)

    pUnkAnnMapMaplex = pFact.Create(CLSID(comtypes.gen.esriMaplex.MaplexAnnotateMap))
    pAnnMapMaplex = CType(pUnkAnnMapMaplex, comtypes.gen.esriMaplex.IMaplexAnnotateMap)

    pMap.AnnotationEngine = pAnnMapMaplex

# Enable the Standard labeling engine
def disableMaplex():    
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)

    pUnkAnnMapStandard = pFact.Create(CLSID(comtypes.gen.esriCarto.AnnotateMap))
    pAnnMapStandard = CType(pUnkAnnMapStandard, comtypes.gen.esriCarto.IAnnotateMap2)

    pMap.AnnotationEngine = pAnnMapStandard

# Set the Maplex placement quality to "Best" (this must be run BEFORE turning on labels but after enabling Maplex)
def setLabelQualityBest():
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
    
    
    pUnkMaplexOverposterProperties = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexOverposterProperties))
    pMaplexOverposterProperties = CType(pUnkMaplexOverposterProperties, comtypes.gen.esriCarto.IMaplexOverposterProperties)
    pMaplexOverposterProperties.PlacementQuality = 3    # set Maplex label quality to highest
    
    # apply it to the Map overposter
    pMapOverposter = CType(pMap, comtypes.gen.esriCarto.IMapOverposter)
    pMapOverposter.OverposterProperties = pMaplexOverposterProperties   
    
    
# CREATE NEW CASE FOLDER AND FILES
# Create new / update existing Case folder and files
# Check to see if new case folder and files has already been created (for the first time); if not, then create them
def saveNewMXD():
    newMXDpath = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "MAILER.mxd"
    mxd.saveACopy(newMXDpath)

    mxdNew = arcpy.mapping.MapDocument(newMXDpath)
    dfNew = arcpy.mapping.ListDataFrames(mxdNew,"*")[0]

    # Set zoom level
    dfNew.extent = df.extent
    dfNew.scale = df.scale

    # Update site path for subject parcels
    for lyr in arcpy.mapping.ListLayers(mxdNew):
        if lyr.name == "Subject Parcels" or lyr.name == "Subject Parcel": 
            lyr.replaceDataSource("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum, "SHAPEFILE_WORKSPACE", "Site")

    # Update adjacents path for adjacent parcels
    for lyr in arcpy.mapping.ListLayers(mxdNew):
        if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel": 
            lyr.replaceDataSource("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum, "SHAPEFILE_WORKSPACE", "Adjacents")
 
    # Save NEW Case map document
    mxdNew.save()
    
def createCase():
    # Set file/folder paths
    newSitepath = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "Site.shp"
    newAdjacentspath = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "Adjacents.shp"
    newMXDpath = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "MAILER.mxd"
    newCasepath = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum)
    #newLabelLegendPathCSV = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "labelLegend.csv"
    #newLabelLegendPathDBF = "W:\\GIS_Mailer\\Mailer_Cases_by_Case\\" + str(caseNum) + "\\" + "labelLegend.dbf"
    # if the Case folder and files already exist, then delete the old ones and replace them with the new ones
    if arcpy.Exists(newCasepath):
        arcpy.Delete_management(newCasepath)
    
    # Save site parcel (for case) and mxd to separate folder (Mailer_Cases_by_Case)
    arcpy.CreateFolder_management("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\", str(caseNum))
    arcpy.CopyFeatures_management(Site_shp, newSitepath)
    arcpy.CopyFeatures_management(Adjacents_shp, newAdjacentspath)
    
    # append new PLAN_ID column
    
    # add appropriate columns from each row to PARCEL and CAMA
  
    planid = arcpy.GetParameterAsText(1) 
    # from newAdjacentspath
    gpinList = []
    gpinCursor = arcpy.da.SearchCursor(newAdjacentspath,["GPIN"])
    for i in gpinCursor:
        gpinList.append(i)
    del gpinCursor
    gpinList = str(gpinList)
    gpinList = gpinList.replace("(u'","")
    gpinList = gpinList.replace("',)","")
    gpinList = gpinList.replace("(u","")
    gpinList = gpinList.replace(",)","")
    
    # append data from shapefiles to corresponding SDE tables
    # get user id (this should work for any user on a Windows 7 computer)
    userid = str(shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0))
    userid = userid.lstrip("C:\\Users\\")
    userid = userid.rstrip("\\Documents")
    
    # get ArcGIS version from dictionary
    arcgisDict = arcpy.GetInstallInfo()
    arcgisVersion = str(arcgisDict['Version'])
    
    # connect to the SDE database (SQL Server); create new connection if not already connected; database=planning (do not save username/password)
    # connectPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde"
    dataList = []
    dataPath = ""
    if str(arcpy.GetParameterAsText(0)) == "COMP":    
        dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.CADJ_PARCEL_CAMA"  
    elif str(arcpy.GetParameterAsText(0)) == "BZA":
        dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.EADJ_PARCEL_CAMA"  
    elif str(arcpy.GetParameterAsText(0)) == "DRD":   
        dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.DADJ_PARCEL_CAMA"  

    fields = ["PLAN_ID", "GPIN"]   
    cursor = arcpy.da.InsertCursor(dataPath, fields)
    for i in range(0, len(gpinList)):
        cursor.insertRow((str(planid), str(gpinList[i]))) # set to NULL if empty, otherwise add value
	
    del cursor
    
    # dataList = []
    # dataPath = ""
    # if str(arcpy.GetParameterAsText(0)) == "COMP":    
        # dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.CADJ_PARCEL_INFO"  
    # elif str(arcpy.GetParameterAsText(0)) == "BZA":
        # dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.EADJ_PARCEL_INFO"  
    # elif str(arcpy.GetParameterAsText(0)) == "DRD":   
        # dataPath = "C:\\Users\\" + userid + "\\AppData\\Roaming\\ESRI\\Desktop" + arcgisVersion + "\\ArcCatalog\\MailerToolbar.sde\\planning.planGIS.DADJ_PARCEL_INFO"  

    # fields = ["PLAN_ID"]   
    # cursor = arcpy.da.InsertCursor(dataPath, fields)
    # #for i in range(0, len(gpinList)):
    # cursor.insertRow((str(planid))) # set to NULL if empty, otherwise add value
	
    # del cursor
		    
   
    
    
    saveNewMXD()
  
  
# Get master condo parcel for each set of selected condos (also the option to get all features in a layer, good for removing duplicates) --- helper function for getMasterCondoLayer(); gets the actual features for the layer specified in getMasterCondoLayer()
def getMasterCondo(lyrList, master=True):    # lyrlist = "subject", "adjacent", "all"; master = master condos only (T/F)
    masterGPINs = list()
    expression = ""
    if master == True:
        expression = "\"CAMA_GPIN\" LIKE '%.000' AND \"USE_CODE\"<>'030' AND \"USE_CODE\"<>'031' AND \"USE_CODE\"<>'322' AND \"USE_CODE\"<>'353' AND \"USE_CODE\"<>'620' AND \"USE_CODE\"<>'660' AND \"USE_CODE\"<>'665' AND \"USE_CODE\"<>'701' AND \"USE_CODE\"<>'703'"
    elif master == False:
        expression = "\"CAMA_GPIN\" LIKE '%'"
    # get list of GPIN's with CAMA_GPIN's containing ".000" in Subject or Adjacents and ALL USE_CODES's (condos and office condos)   XXX AND USE_CODE = 698 (all CONDOS only, with a master (.000))
    for lyr in lyrList:  
        if master == True:
            masterCursor = arcpy.da.SearchCursor(lyr, "GPIN", expression)	        
        elif master == False:
            masterCursor = arcpy.da.SearchCursor(lyr, "GPIN", expression)		
        masterCursorlist = list(masterCursor)
        masterCursorlen = len(masterCursorlist)
        masterGPINlist = list()
        for j in range(0,masterCursorlen):
            masterCursoritem = masterCursorlist[j]
            masterCursoritemString = str(masterCursoritem)
            masterCursoritemStringRep1 = masterCursoritemString.replace("(u'","")
            masterCursoritemStringRep2 = masterCursoritemStringRep1.replace("',)","")
            masterGPINlist.append(masterCursoritemStringRep2)   
	
        masterGPINs.append(masterGPINlist)
    return masterGPINs	
    
# Get layer names to use for extracting master condos
def getMasterCondoLayer(lyrName):
    lyrList = list()
    if lyrName == "subject" and (arcpy.Exists("Site")==True or arcpy.Exists("Subject Parcel")==True or arcpy.Exists("Subject Parcels")==True):
        for lyr in arcpy.mapping.ListLayers(mxd):
            if lyr.name == "Site" or lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels":
                lyrList.append(str(lyr.name))
        return lyrList
    if lyrName == "adjacent" and (arcpy.Exists("Adjacents")==True or arcpy.Exists("Adjacent Parcel")==True or arcpy.Exists("Adjacent Parcels")==True):
        for lyr in arcpy.mapping.ListLayers(mxd):
            if lyr.name == "Adjacents" or lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels":
                lyrList.append(str(lyr.name))
        return lyrList
    if lyrName == "all" and (arcpy.Exists("Site")==True or arcpy.Exists("Subject Parcel")==True or arcpy.Exists("Subject Parcels")==True or arcpy.Exists("Adjacents")==True or arcpy.Exists("Adjacent Parcel")==True or arcpy.Exists("Adjacent Parcels")==True):
        for lyr in arcpy.mapping.ListLayers(mxd):
            if lyr.name == "Site" or lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels" or lyr.name == "Adjacents" or lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels":
                lyrList.append(str(lyr.name))
        return lyrList

# Omit individual condos from Site/Adjacents layers; still show the master parcel for each condo set     
def omitCondos():
    for lyr in arcpy.mapping.ListLayers(mxd):
        if lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels" or lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels": 
            # get list of master condos --- GPIN's with CAMA_GPIN's containing ".000" in Subject or Adjacents and ALL USE_CODES's (condos and office condos) XXX AND USE_CODE = 698 (all CONDOS only, with a master (.000))
            layersToGet = getMasterCondoLayer("all")
            masterGPINlist = getMasterCondo(layersToGet, master=True)
    
            if lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels":
                # if master condos exist
                if len(masterGPINlist[0]) > 0:
		
                    # remove OTHERS from Site/Adjacents layers, so it is *just condos with a master (.000)* + non-condos with a variety of CAMA_GPIN's
                    clearRefresh()
                    arcpy.MakeFeatureLayer_management(lyr, lyr.name+"_lyr")     # make a new feature layer from the site/adjacents layer, to allow for deletion of selection
		    
                    for gpin in masterGPINlist[0]:    # just focus on condos containing a master GPIN (*.000); ignore non-condos (*.001, none, etc.)
                        arcpy.SelectLayerByAttribute_management(lyr.name+"_lyr","ADD_TO_SELECTION", "\"GPIN\" = '"+gpin+"'")
                        arcpy.SelectLayerByAttribute_management(lyr.name+"_lyr","REMOVE_FROM_SELECTION", "\"CAMA_GPIN\" LIKE '%.000'")
                        arcpy.DeleteIdentical_management(lyr.name, ["CAMA_GPIN"])	# delete features with identical CAMA_GPINs    
                    arcpy.DeleteFeatures_management(lyr.name+"_lyr")  

            if lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels": 
                # if master condos exist	    
                if len(masterGPINlist[1]) > 0:
		
                    # remove OTHERS from Site/Adjacents layers, so it is *just condos with a master (.000)* + non-condos with a variety of CAMA_GPIN's
                    clearRefresh()
                    arcpy.MakeFeatureLayer_management(lyr, lyr.name+"_lyr")     # make a new feature layer from the site/adjacents layer, to allow for deletion of selection
		    
                    for gpin in masterGPINlist[1]:    # just focus on condos containing a master GPIN (*.000); ignore non-condos (*.001, none, etc.)
                        arcpy.SelectLayerByAttribute_management(lyr.name+"_lyr","ADD_TO_SELECTION", "\"GPIN\" = '"+gpin+"'")
                        arcpy.SelectLayerByAttribute_management(lyr.name+"_lyr","REMOVE_FROM_SELECTION", "\"CAMA_GPIN\" LIKE '%.000'")
                        arcpy.DeleteIdentical_management(lyr.name, ["CAMA_GPIN"])	# delete features with identical CAMA_GPINs	
                    arcpy.DeleteFeatures_management(lyr.name+"_lyr")
            clearRefresh()
	    
    # delete extra feature layer for site and adjacents
    for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
        if lyr.name == lyr.name+"_lyr":
            arcpy.mapping.RemoveLayer(df, lyr)    
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()

    
# Check for duplicate GPINs (existence of condos) before omitting or restoring condos (TEST FOR COMMS TOWERS!!!!!!!!!!)    
def checkforDuplicates():  
    # get list of all GPINS (including condos) for each layer (subject = 0; adjacents = 1)  
    layersToGet = getMasterCondoLayer("all")
    masterGPINlist = getMasterCondo(layersToGet, master=False)
    
    # remove duplicates to get list of main GPINs only for each layer
    subjectGPINlist = list(set(masterGPINlist[0]))
    adjacentGPINlist = list(set(masterGPINlist[1]))

    # get number of duplicate GPINs for each layer (one duplicate (duplicateCount = 1) = one GPIN has been repeated multiple times)
    duplicateCount = 0 
    for i in range(0, len(subjectGPINlist)):
        if masterGPINlist[0].count(subjectGPINlist[i]) > 1:
            duplicateCount = duplicateCount + 1

    for i in range(0, len(adjacentGPINlist)):
        if masterGPINlist[1].count(adjacentGPINlist[i]) > 1:
            duplicateCount = duplicateCount + 1
	    
    return duplicateCount


# If condos were removed previously for this Case, then restore them (if "Omit Condos" is unchecked); or include all condos    
def restoreCondos():
    # Append new Subject/Adjacents parcels (individual condos) to existing Subject/Adjacents parcels layers	
    for lyr in arcpy.mapping.ListLayers(mxd):
        if lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels" or lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels": 
            # get list of master condos --- GPIN's with CAMA_GPIN's containing ".000" in Subject or Adjacents and ALL USE_CODES's (condos and office condos) XXX AND USE_CODE = 698 (all CONDOS only, with a master (.000))
            layersToGet = getMasterCondoLayer("all")
            masterGPINlist = getMasterCondo(layersToGet, master=True)
	    
            # SUBJECT PARCELS   
            if lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels": 
                # if master condos exist
                if len(masterGPINlist[0]) > 0:
		
                    # get list of all master condos --- GPIN's with CAMA_GPIN's containing ".000" in Subject Parcels
                    clearRefresh()
                    layersToGet = getMasterCondoLayer("subject")
                    masterGPINlist = getMasterCondo(layersToGet, master=True)
	    
                    # get all the GPINs for one master GPIN (not including the master GPIN, because it is already included in Subject Parcels)               
                    #clearRefresh()		    
                    for gpin in masterGPINlist[0]:
                        arcpy.SelectLayerByAttribute_management("Tax Parcels - Mailer","ADD_TO_SELECTION", "\"GPIN\" = '"+gpin+"' AND \"CAMA_GPIN\" NOT LIKE '%.000'")
                       
                    # Add selected parcels from Tax Parcels layer to Condos_add layer  
                    arcpy.CopyFeatures_management(tax_parcels, Condos_add, "", "0", "0", "0")
                    condos_add = arcpy.mapping.Layer(Condos_add)
                    arcpy.mapping.AddLayer(df, condos_add, "TOP")	

                    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.	    
                    # If Condos_add layer contains parcels, then append them to the Subject shapefile
                    desc = arcpy.Describe(condos_add)
                    if int(str(len(desc.fidSet.split(";")))) > 0:
                        if arcpy.Exists("Subject Parcel") == True:
                            arcpy.Append_management(condos_add, "Subject Parcel")
                            arcpy.DeleteIdentical_management("Subject Parcel", ["CAMA_GPIN"])	# delete features with identical CAMA_GPINs	    
                        elif arcpy.Exists("Subject Parcels") == True:
                            arcpy.Append_management(condos_add, "Subject Parcels")
                            arcpy.DeleteIdentical_management("Subject Parcels", ["CAMA_GPIN"])	
		    
            # ADJACENT PARCELS            
            elif lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels": 
                # if master condos exist
                if len(masterGPINlist[1]) > 0:
	    
                    # get list of master condos --- GPIN's with CAMA_GPIN's containing ".000" in Adjacent Parcels
                    clearRefresh()
                    layersToGet = getMasterCondoLayer("adjacent")
                    masterGPINlist = getMasterCondo(layersToGet, master=True)

                    # get all the GPINs for one master GPIN (not including the master GPIN, because it is already included in Adjacent Parcels)	    
                    clearRefresh()		    
                    for gpin in masterGPINlist[0]:
                        arcpy.SelectLayerByAttribute_management("Tax Parcels - Mailer","ADD_TO_SELECTION", "\"GPIN\" = '"+gpin+"' AND \"CAMA_GPIN\" NOT LIKE '%.000'")	
	        
                    # Add selected parcels from Tax Parcels layer to Condos_add layer  
                    arcpy.CopyFeatures_management(tax_parcels, Condos_add, "", "0", "0", "0")	    
	        
                    # Generate label fields so that label fields for Condos_add and Adjacents_shp are matching
                    arcpy.AddField_management(Condos_add,"Latitude","DOUBLE","#","#","#","#","NULLABLE","NON_REQUIRED","#")
                    arcpy.AddField_management(Condos_add,"KeyLabel","SHORT","#","#","#","#","NULLABLE","NON_REQUIRED","#")
                    arcpy.AddField_management(Condos_add,"LabelMain","SHORT","#","#","#","#","NULLABLE","NON_REQUIRED","#")                
                    # Add selected parcels from Tax Parcels layer to Condos_add layer  
                    condos_add = arcpy.mapping.Layer(Condos_add)
                    arcpy.mapping.AddLayer(df, condos_add, "TOP")

                    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.	    
                    # If Condos_add layer contains parcels, then append them to the Adjacents shapefile
                    desc = arcpy.Describe(condos_add)
                    if int(str(len(desc.fidSet.split(";")))) > 0:
                        if arcpy.Exists("Adjacent Parcel") == True:
                            arcpy.Append_management(condos_add, "Adjacent Parcel")
                            arcpy.DeleteIdentical_management("Adjacent Parcel", ["CAMA_GPIN"])
                        elif arcpy.Exists("Adjacent Parcels") == True:
                            arcpy.Append_management(condos_add, "Adjacent Parcels")
                            arcpy.DeleteIdentical_management("Adjacent Parcels", ["CAMA_GPIN"])
	
            # Remove the condos_add layer (temporary layer to hold parcels to add)
            for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
                if lyr.name == "Condo_add":
                    arcpy.mapping.RemoveLayer(df, lyr)
            arcpy.Delete_management(Condos_add)
	    clearRefresh()
            arcpy.RefreshActiveView()   


# Remove duplicate parcel labels and set up other labeling properties	
def removeDuplicates():
    # calculate LabelMain field using omitCondos(), or only main parcels, used to label main parcels only (omitting duplicates and maintaining proper positioning of label in center of parcel)
    arcpy.AddField_management(Adjacents_shp,"LabelMain","SHORT","#","#","#","#","NULLABLE","NON_REQUIRED","#")

    # expression to fill LabelMain field with '0' values for all CAMA_GPINS not containing *.000 (masters), and '1' values for all CAMA_GPINS containing *.000 (masters) or not containing *.***
    expression = "getLabelMain(str(!CAMA_GPIN!))"
    codeblock = """def getLabelMain(cama_gpin):
        if cama_gpin.count('.') > 0:
            endsWith = int(cama_gpin.rsplit('.')[1])
            if endsWith > 0:
                return 0
            else:
                return 1
        else:
            return 1"""
    
    arcpy.CalculateField_management(Adjacents_shp,"LabelMain",expression,"PYTHON_9.3",codeblock)
	
    # Get Adjacent Parcels layer	
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)        
    pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
    pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
    pLayerAdj = pMap.Layer(1)
    pUnkFeatureLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.FeatureLayer))
    pUnkFeatureLayer = pLayerAdj
    pFeatureLayer = CType(pUnkFeatureLayer, comtypes.gen.esriCarto.IFeatureLayer)    
    pGeoFeatureLayer = CType(pFeatureLayer, comtypes.gen.esriCarto.IGeoFeatureLayer)

    
    # Set up Maplex labeling properties - Remove Duplicates, etc.
    # 1
    #-----------------------------------------------------
    # Class 1 - for large parcels
    pUnkMaplexOverposterLayerProps = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexOverposterLayerProperties))
    pMaplexOverposterLayerProps = CType(pUnkMaplexOverposterLayerProps, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties)

    pMaplexOverposterLayerProps.CanStackLabel = False
    pMaplexOverposterLayerProps.MaximumLabelOverrun = 3   # 3 points (default = 36 points)
    pMaplexOverposterLayerProps.CanOverrunFeature = True    # can Overrun Feature by a certain amount ...(see above)
    pMaplexOverposterLayerProps.CanReduceFontSize = False
    pMaplexOverposterLayerProps.CanAbbreviateLabel = False
    pMaplexOverposterLayerProps.ThinningDistance = 0
    pMaplexOverposterLayerProps.RepeatLabel = False    # removes label from one of two parcels with same owner/GPIN (can remove - USED FOR LINE FEATURES ONLY)
    pMaplexOverposterLayerProps.PreferHorizontalPlacement = False
    pMaplexOverposterLayerProps.CanPlaceLabelOutsidePolygon = False
    pMaplexOverposterLayerProps.PolygonPlacementMethod = 0    # Horizontal placement within polygon
 
    #------------------------------------------------------
    # Class 2 - for small parcels
    pUnkMaplexOverposterLayerPropsSM = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexOverposterLayerProperties))
    pMaplexOverposterLayerPropsSM = CType(pUnkMaplexOverposterLayerPropsSM, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties)

    pMaplexOverposterLayerPropsSM.CanStackLabel = False
    pMaplexOverposterLayerPropsSM.MaximumLabelOverrun = 3   # 3 points (default = 36 points)
    pMaplexOverposterLayerPropsSM.CanOverrunFeature = True    # can Overrun Feature by a certain amount ...(see above)
    pMaplexOverposterLayerPropsSM.CanReduceFontSize = False
    pMaplexOverposterLayerPropsSM.CanAbbreviateLabel = False
    pMaplexOverposterLayerPropsSM.ThinningDistance = 0
    pMaplexOverposterLayerPropsSM.RepeatLabel = False    # removes label from one of two parcels with same owner/GPIN (can remove - USED FOR LINE FEATURES ONLY)
    pMaplexOverposterLayerPropsSM.PreferHorizontalPlacement = False
    pMaplexOverposterLayerPropsSM.CanPlaceLabelOutsidePolygon = False
    pMaplexOverposterLayerPropsSM.PolygonPlacementMethod = 0    # Horizontal placement within polygon          
    
    # get parcel count (number of labels), and if it's 100+, then set label placement to Straight within polygon (so that all labels are shown)
    layersToGet = getMasterCondoLayer("adjacent")
    masterGPINlist = getMasterCondo(layersToGet, master=False)
    list1 = masterGPINlist[0]    # ALL GPINs
    list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed
    if len(list2) > 99:
        pMaplexOverposterLayerPropsSM.PolygonPlacementMethod = 1    # Straight placement within polygon
    #------------------------------------------------------   
   
    # for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
        # if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel":
            # lyrName = lyr.name
	
            # # get list of GPINs to set KeyLabel values and for GetCount (count of GPINs)
            # layersToGet = getMasterCondoLayer("adjacent")
            # masterGPINlist = getMasterCondo(layersToGet, master=False)
            # list1 = masterGPINlist[0]
            # list2 = sorted(set(list1), key=list1.index)	   
	
            # # get ratio of smallest area polygon (parcel) to largest area polygon (parcel); used to tell ArcMap when to implement key numbering
            # parcelAreaField = [f.name for f in arcpy.ListFields(lyrName, "*STAr*")]
            # area = [row[0] for row in arcpy.da.SearchCursor(lyrName, parcelAreaField)]
            # minArea = min(area)
            # maxArea = max(area)
            # ratioArea = minArea/maxArea

            # # If number of parcels > 15 OR if ratio of smallest to largest parcel is less than 0.15 (KEY NUMBERING), then use Horizontal labeling; otherwise, use Straight labeling
            # if len(list2) > 15 and df.scale > 1900 or ratioArea < 0.15:
                # pMaplexOverposterLayerProps.PolygonPlacementMethod = 0
            # else:
                # pMaplexOverposterLayerProps.PolygonPlacementMethod = 0
	
	
    # 2		
    for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
        if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel":

            # get list of GPINs to set KeyLabel values and for GetCount (count of GPINs)
            layersToGet = getMasterCondoLayer("adjacent")
            masterGPINlist = getMasterCondo(layersToGet, master=False)
            list1 = masterGPINlist[0]
            list2 = sorted(set(list1), key=list1.index)	   

            # If there are parcels with the same GPINs (condos), then center the label; if not, then use regular settings (fixed positioning off)
            if len(list2) != len(list1):    
                # 2 - set label zone priority (center only; 8 = 1)
                pMaplexOverposterLayerProps2 = CType(pMaplexOverposterLayerProps, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties2)   
                pMaplexOverposterLayerProps2SM = CType(pMaplexOverposterLayerPropsSM, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties2) 
                # pMaplexOverposterLayerProps2.PolygonInternalZones[0] = 0
                # pMaplexOverposterLayerProps2.PolygonInternalZones[1] = 0 #3
                # pMaplexOverposterLayerProps2.PolygonInternalZones[2] = 0
                # pMaplexOverposterLayerProps2.PolygonInternalZones[3] = 0 #2
                # pMaplexOverposterLayerProps2.PolygonInternalZones[4] = 2 #5
                # pMaplexOverposterLayerProps2.PolygonInternalZones[5] = 0
                # pMaplexOverposterLayerProps2.PolygonInternalZones[6] = 0
                # pMaplexOverposterLayerProps2.PolygonInternalZones[7] = 0 #4
                # pMaplexOverposterLayerProps2.PolygonInternalZones[8] = 1
                pMaplexOverposterLayerProps2.EnablePolygonFixedPosition = False
                pMaplexOverposterLayerProps2SM.EnablePolygonFixedPosition = False
                pMaplexOverposterLayerProps2.PolygonFeatureType = 1    # Land Parcel placement	  
                pMaplexOverposterLayerProps2SM.PolygonFeatureType = 1    # Land Parcel placement	
            else:
                pMaplexOverposterLayerProps2 = CType(pMaplexOverposterLayerProps, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties2)  	
                pMaplexOverposterLayerProps2SM = CType(pMaplexOverposterLayerPropsSM, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties2) 		
                pMaplexOverposterLayerProps2.EnablePolygonFixedPosition = False
                pMaplexOverposterLayerProps2SM.EnablePolygonFixedPosition = False
                pMaplexOverposterLayerProps2.PolygonFeatureType = 1    # Land Parcel placement	
                pMaplexOverposterLayerProps2SM.PolygonFeatureType = 1    # Land Parcel placement	

		
    # 3
    pMaplexOverposterLayerProps3 = CType(pMaplexOverposterLayerProps, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties4)
    pMaplexOverposterLayerProps3SM = CType(pMaplexOverposterLayerPropsSM, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties4)
    
    # Check for and label all multi-part features (all parts)
    for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
        if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel":

            # count variable for number of total parts in layer
            numParts = 0
            for row in arcpy.da.SearchCursor(lyr, ["OID@", "SHAPE@"]):
                # step through each part of the feature
                for part in row[1]:
                    # accumulate the parts
                    numParts = numParts + 1
            lyrCount = arcpy.GetCount_management(lyr)
	    # if a feature in the layer has multi-part polygons (parcels), then keep duplicate labels; otherwise, remove duplicate labels
            if numParts / float(str(lyrCount)) > 1:
                pMaplexOverposterLayerProps.ThinDuplicateLabels = False
                pMaplexOverposterLayerPropsSM.ThinDuplicateLabels = False               
                pMaplexOverposterLayerProps3.LabelLargestPolygon = False
                pMaplexOverposterLayerProps3SM.LabelLargestPolygon = False
            else:
                pMaplexOverposterLayerProps.ThinDuplicateLabels = True
                pMaplexOverposterLayerPropsSM.ThinDuplicateLabels = True
                pMaplexOverposterLayerProps3.LabelLargestPolygon = True    
                pMaplexOverposterLayerProps3SM.LabelLargestPolygon = True  
		
		
    # Get annotation (labeling) properties for Adjacent Parcels --- AnnotateLayerPropertiesCollection
    pUnkAnnoLayerPropsColl = pFact.Create(CLSID(comtypes.gen.esriCarto.AnnotateLayerPropertiesCollection))
    pAnnoLayerPropsColl = CType(pUnkAnnoLayerPropsColl, comtypes.gen.esriCarto.IAnnotateLayerPropertiesCollection)
    pAnnoLayerPropsColl = pGeoFeatureLayer.AnnotationProperties

    # Set Class 1 labeling properties --- AnnotateLayerProperties
    pUnkAnnoLayerProps = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexLabelEngineLayerProperties))
    pAnnoLayerProps = CType(pUnkAnnoLayerProps, comtypes.gen.esriCarto.IAnnotateLayerProperties)
    pAnnoLayerProps.WhereClause = "\"LabelMain\" = 1 AND \"SHAPE_STAr\" > 20000"    # for large parcels          
    #pAnnoLayerProps.WhereClause = "\"GPIN\" = \"CAMA_GPIN\" OR \"CAMA_GPIN\" LIKE '%.000'"

    # Set Class 2 labeling properties --- AnnotateLayerProperties
    pUnkAnnoLayerPropsSM = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexLabelEngineLayerProperties))
    pAnnoLayerPropsSM = CType(pUnkAnnoLayerPropsSM, comtypes.gen.esriCarto.IAnnotateLayerProperties)
    pAnnoLayerPropsSM.WhereClause = "\"LabelMain\" = 1 AND \"SHAPE_STAr\" <= 20000"    # for small parcels
    #pAnnoLayerProps.WhereClause = "\"GPIN\" = \"CAMA_GPIN\" OR \"CAMA_GPIN\" LIKE '%.000'"    

    # LabelEngineLayerProperties - Class 1 (large parcels)
    pLabelEngineLayerProps2 = CType(pAnnoLayerProps, comtypes.gen.esriCarto.ILabelEngineLayerProperties2)
    pLabelEngineLayerProps2.OverposterLayerProperties = pMaplexOverposterLayerProps2
    pLabelEngineLayerProps2.Expression = "[GPIN]"
    
    # LabelEngineLayerProperties - Class 2 (small parcels)
    pLabelEngineLayerProps2SM = CType(pAnnoLayerPropsSM, comtypes.gen.esriCarto.ILabelEngineLayerProperties2)
    pLabelEngineLayerProps2SM.OverposterLayerProperties = pMaplexOverposterLayerProps2SM
    pLabelEngineLayerProps2SM.Expression = "[GPIN]"

    # Apply them to Adjacent Parcels layer
    pAnnoLayerPropsColl.Clear()
    pGeoFeatureLayer.DisplayAnnotation = True
    pAnnoLayerPropsColl.Add(pAnnoLayerProps)
    pAnnoLayerPropsColl.Add(pAnnoLayerPropsSM)
    pMxDoc.ActiveView.Refresh()
    
    
    
    
    
    
# def labelWithCondos():    # omit condos unchecked    # CREATE A SEPARATE	 def labelWithoutCondos()     

# Label main Adjacent Parcels only (condos included)    # create separate method for Adjacent Parcels (no condos)     --- GeoFeatureLayer








# i = 0
# for i in range(pAnnoLayerPropsColl.Count-1):
    # pAnnoLayerProps = CType(pAnnoLayerPropsColl.QueryItem(0), comtypes.gen.esriCarto.IAnnotateLayerProperties)
    # #pAnnoLayerProps = pAnnoLayerPropsColl.QueryItem(i)

    # pUnkMaplexLabelEngineLayerProps = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexLabelEngineLayerProperties))
    # pLabelEngineLayerProps = CType(pUnkMaplexLabelEngineLayerProps, comtypes.gen.esriCarto.ILabelEngineLayerProperties)
    # pLabelEngineLayerProps = pAnnoLayerProps
    # pLabelEngineLayerProps.Expression = "[GPIN]"
    
    # pMxDoc.ActiveView.Refresh()


# #pUnkPlacedElements = pFact.Create(CLSID(comtypes.gen.esriCarto.ElementCollection))
# #pPlacedElements = CType(pUnkPlacedElements, comtypes.gen.esriCarto.IElementCollection)
# #pUnkUnplacedElements = pFact.Create(CLSID(comtypes.gen.esriCarto.ElementCollection))
# #pUnplacedElements = CType(pUnkUnplacedElements, comtypes.gen.esriCarto.IElementCollection)

# # Set up annotation properties    
# #pUnkAnnoPropsCol = pFact.Create(CLSID(comtypes.gen.esriCarto.AnnotateLayerPropertiesCollection))
# #pAnnoPropsCol = CType(pUnkAnnoPropsCol, comtypes.gen.esriCarto.IAnnotateLayerPropertiesCollection)

# # Set up Maplex label engine and IAnnotateLayer properties
# #pUnkMaplexLabelEngineLayerProperties = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexLabelEngineLayerProperties))
# #pAnnoProps = CType(pUnkMaplexLabelEngineLayerProperties, comtypes.gen.esriCarto.IAnnotateLayerProperties)

# # Set up Label Engine Properties and Maplex Overposter Properties
# pUnkLabelEngineLayerProperties = pFact.Create(CLSID(comtypes.gen.esriCarto.LabelEngineLayerProperties))
# pAnnoProps = CType(pUnkLabelEngineLayerProperties, comtypes.gen.esriCarto.IAnnotateLayerProperties)

# pUnkMaplexOverposterLayerProperties = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexOverposterLayerProperties))
# pMaplexOverposterLayerProperties = CType(pUnkMaplexOverposterLayerProperties, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties)

# # if labeling for Adjacents is turned on
 
# if pGeoFeatureLayer.DisplayAnnotation == True:
    # pAnnoPropsCol = pGeoFeatureLayer.AnnotationProperties
    # pAnnoPropsCol2 = CType(pAnnoPropsCol, comtypes.gen.esriCarto.IAnnotateLayerPropertiesCollection2)
    
    # i = 0
    # if i <= pAnnoPropsCol2.Count - 1:
        # #pAnnoProps = CType(pAnnoPropsCol.QueryItem(i), comtypes.gen.esriCarto.IAnnotateLayerProperties)
        # #pAnnoProps = pAnnoPropsCol.QueryItem(i)
        # pAnnoLayerProp = pAnnoPropsCol2.QueryItem(i)[0]
	
	# pAnnoLayerProp.WhereClause = "\"CAMA_GPIN\" LIKE '%.000'"
	
	# pLabelEngineProp = CType(pAnnoLayerProp, comtypes.gen.esriCarto.ILabelEngineLayerProperties2)
	# pLabelEngineProp.OverposterLayerProperties = pL
   
        # pLabelEngineLayerProperties2 = CType(pAnnoProps, comtypes.gen.esriCarto.ILabelEngineLayerProperties2)
   
        # pMaplexOverposterLayerProperties = pLabelEngineLayerProperties2.OverposterLayerProperties
        # pMaplexOverposterLayerProperties.ThinDuplicateLabels = True
    
        # #pUnkMaplexOverposterLayerProperties = pFact.Create(CLSID(comtypes.gen.esriCarto.MaplexOverposterLayerProperties))
        # #pMaplexOverposterLayerProperties = CType(pUnkMaplexOverposterLayerProperties, comtypes.gen.esriCarto.IMaplexOverposterLayerProperties)
        # #pMaplexOverposterLayerProperties.ThinDuplicateLabels = True
    
        # # apply it to the Map overposter
        # pMapOverposter = CType(pMap, comtypes.gen.esriCarto.IMapOverposter)
        # pMapOverposter.OverposterProperties = pMaplexOverposterLayerProperties  
 
def createGraphics():    # creates all graphics, except for Label Legend (*must be run before createCase())
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
    
    pAV = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IActiveView)    # map layout extent (used to create Envelope for graphic element position)
    pSD = pAV.ScreenDisplay
    pEnv = pAV.Extent 

    # SET INITIAL GRAPHICS ELEMENT PROPERTIES

    # Colors
    pUnkBlack = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))    # black
    pBlack = CType(pUnkBlack, comtypes.gen.esriDisplay.IRgbColor)
    pBlack.Red = 0
    pBlack.Green = 0
    pBlack.Blue = 0	    

    pUnkWhite = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))    # white
    pWhite = CType(pUnkWhite, comtypes.gen.esriDisplay.IRgbColor)
    pWhite.Red = 255    #333
    pWhite.Green = 255    #111
    pWhite.Blue = 255    #222	    
    
    # Set fill and line color (see COLORS above) (for Label Legend and Legend Title Text boxes   # USE FOR LABEL LEGEND AND RECTANGLE BOX ONLY; THE REST YOU DON'T NEED IT)   
    # Create line symbol with outline color    
    pUnkSimpleLineSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleLineSymbol))
    pSimpleLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ISimpleLineSymbol)
    pLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ILineSymbol)
    
    pLineSymbol.Color = pWhite
    pLineSymbol.Width = 0	 
    
    pUnkSymbolBorder = pFact.Create(CLSID(comtypes.gen.esriCarto.SymbolBorder))
    pSymbolBorder = CType(pUnkSymbolBorder, comtypes.gen.esriCarto.ISymbolBorder)
    pSymbolBorder.LineSymbol = pLineSymbol	
    
    # Create fill symbol with fill color    
    pUnkSimpleFillSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleFillSymbol))
    pSimpleFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.ISimpleFillSymbol)
    pFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.IFillSymbol)
    
    pSimpleFillSymbol.Color = pWhite
    pSimpleFillSymbol.Outline = pLineSymbol
    
    pUnkSymbolBackground = pFact.Create(CLSID(comtypes.gen.esriCarto.SymbolBackground))		
    pSymbolBackground = CType(pUnkSymbolBackground, comtypes.gen.esriCarto.ISymbolBackground)
    pSymbolBackground.FillSymbol = pSimpleFillSymbol


    # Loop through elements and get Data Frame element (used to snap Rectangle box to Data Frame element)    
    #pGCSelDF = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    pGCSelDF = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    pGCSelDF.SelectAllElements()
    elementCountDF = pGCSelDF.ElementSelectionCount    # get count of all elements in Page Layout (including data frame)
    
    pGraphicsContainerDF = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    pGraphicsContainerDF.Reset()
    pElemDF = pGraphicsContainerDF.Next()  # each element
    #pElemReportDF = []    # report of names and types of all elements
    pElemListDF = []    # list of all elements (minus the data frame element, which we want to keep)
    for i in range(0, elementCountDF):
        if pElemDF is not None:
            pElemPropDF = CType(pElemDF, comtypes.gen.esriCarto.IElementProperties)
           # pElemReportDF.append("Name: " + pElemPropDF.Name + ", Type: " + pElemPropDF.Type)
            pElemListDF.append(pElemDF)   # append each element to the list of elements
            pElemDF = pGraphicsContainerDF.Next() # go to the next element  
        i = i + 1	
    #pythonaddins.MessageBox(pElemReportDF, 0)    # report the Name and Type properties of each element   
    pGCSelDF.UnselectAllElements()
    pMxDoc.ActiveView.Refresh()	

    
    # Replace the following each time:  # NEED TO RUN CREATE CASE() AFTER THESE
    
    # CODE FOR RECTANGULAR BOX (CONTAINER FOR THE ABOVE) - ***** do this one next! # http://edndoc.esri.com/arcobjects/9.2/ComponentHelp/esriCarto/IElement.htm
    # Set rectangular box location 
    pDataFrame = CType(pElemListDF[elementCountDF-1], comtypes.gen.esriCarto.IElement)
    dXmaxRect = pDataFrame.Geometry.Envelope.XMax    # relative to Data Frame element
    dYmaxRect = pEnv.YMax - 7.40 #- 0.25    # relative to dYmaxLegendTitle
    dXminRect = pEnv.XMax - 5.25          # relative to dXminLegendTitle
    dYminRect = pDataFrame.Geometry.Envelope.YMin    # relative to Data Frame element
    
    # Re-size Rectangular Box according to existence/length of Name/Description text
    # if Name and Description text are BOTH empty
    if str(arcpy.GetParameterAsText(3)) == "" and str(arcpy.GetParameterAsText(4)) == "":     # also means font size = 13
        dYmaxRect = dYmaxRect - 0.25    # reduce box YMax by 2 lines
    # if only Description text is empty (assuming Description text will not exist without Name text)
    elif str(arcpy.GetParameterAsText(3)) != "" and str(arcpy.GetParameterAsText(4)) == "":
        dYmaxRect = dYmaxRect - 0.15    # reduce box YMax by 1 line
    
    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminRect, dYminRect, dXmaxRect, dYmaxRect)
 
    # Create rectangle element
    pUnkRectangle = pFact.Create(CLSID(comtypes.gen.esriCarto.RectangleElement))
    pRectangle = CType(pUnkRectangle, comtypes.gen.esriCarto.IRectangleElement)
    #pRectangle.Symbol = pLegendTitleTextSymbol
    
    # Set frame properties	 
    pFillProps = CType(pRectangle, comtypes.gen.esriCarto.IFillShapeElement)
    pLineSymbol.Color = pBlack
    pLineSymbol.Width = 1.0    
    pSimpleFillSymbol.Outline = pLineSymbol
    pSimpleFillSymbol.Color = pWhite    
    pFillProps.Symbol = pSimpleFillSymbol

    pRectElement = CType(pRectangle, comtypes.gen.esriCarto.IElement)
    pRectElement.Geometry = pEnvelope

    # Create the graphics container
    pRectContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)

    # Add the text element to the graphics container
    pRectContainer.AddElement(pRectElement, 0)

    # Do a partial refresh
    #pGCSel = CType(pMap, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    #pGCSel.SelectElement(pRectElement)
    #iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    #pAV.PartialRefresh(iOpt, None, None)   
    clearRefresh() 
    
    # CODE FOR LEGEND
    # Create container for legend
    pUnkMSFrame = pFact.Create(CLSID(comtypes.gen.esriCarto.MapSurroundFrame))
    pMSFrame = CType(pUnkMSFrame, comtypes.gen.esriCarto.IMapSurroundFrame)
    pGContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    #pGContainer.AddElement(pMSFrame, 0)

    # Set legend location
    pAV = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IActiveView)    # map layout extent
    pSD = pAV.ScreenDisplay
    pEnv = pAV.Extent
    
    # If Name and Description are greater than 28 characters, then adjust Legend position
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        dXmaxLegend = dXmaxRect - 1.00    # all relative to the position of rectangular box
        dYmaxLegend = dYminRect + 0.05
        dYminLegend = dYminRect + 0.01 
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        dXmaxLegend = dXmaxRect - 1.00
        dYmaxLegend = dYminRect + 0.05
        dYminLegend = dYminRect + 0.01 
    else:        
        dXmaxLegend = dXmaxRect - 1.00    # all relative to the position of rectangular box    #pEnv.XMax - 0.75
        dYmaxLegend = dYminRect + 0.20    #pEnv.YMax - 5.25  
    dXminLegend = dXminRect + 0.10    #pEnv.XMax - 5.13
    dYminLegend = dYminRect + 0.10    #pEnv.YMax - 8.50

    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminLegend, dYminLegend, dXmaxLegend, dYmaxLegend)
 
    # Create Legend
    pUnkLegend = pFact.Create(CLSID(comtypes.gen.esriCarto.Legend))
    pLegend = CType(pUnkLegend, comtypes.gen.esriCarto.ILegend2)

    pLegend.Map = pMxDoc.FocusMap
    pLegend.AutoAdd = 0            # 0 = false, 1 = true
    #pLegend.ScaleSymbols = 1
    pLegend.Title = ""
    pLegend.ClearItems()
   
    # Set text symbol for labels
    pUnkLegendLabelTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pLegendLabelTextSymbol = CType(pUnkLegendLabelTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pLegendLabelTextSymbol.Color = pBlack
    pLegendLabelTextSymbol.Size = 8  

    pUnkLegItemClassFormatSite = pFact.Create(CLSID(comtypes.gen.esriCarto.LegendClassFormat))        
    pLegItemClassFormatSite = CType(pUnkLegItemClassFormatSite, comtypes.gen.esriCarto.ILegendClassFormat)

    # If Name and Description are greater than 28 characters, then adjust Legend text    
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pLegendLabelTextSymbol.VerticalAlignment = 3
        pLegItemClassFormatSite.LabelSymbol = pLegendLabelTextSymbol
        #pLegItemClassFormatSite.PatchHeight = 10
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pLegendLabelTextSymbol.VerticalAlignment = 3
        pLegItemClassFormatSite.LabelSymbol = pLegendLabelTextSymbol
        #pLegItemClassFormatSite.PatchHeight = 10

    # Add Subject and Adjacents Layers to Legend
    pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
    pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
    pLayerSite = pMap.Layer(0)

    pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
    pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
    pLayerAdj = pMap.Layer(1)

    pUnkLegItemSite = pFact.Create(CLSID(comtypes.gen.esriCarto.HorizontalLegendItem))
    pLegItemSite = CType(pUnkLegItemSite, comtypes.gen.esriCarto.ILegendItem)
    pLegItemSite.Layer = pLayerSite
    pLegItemSite.LegendClassFormat = pLegItemClassFormatSite

    pUnkLegItemAdj = pFact.Create(CLSID(comtypes.gen.esriCarto.HorizontalLegendItem))
    pLegItemAdj = CType(pUnkLegItemAdj, comtypes.gen.esriCarto.ILegendItem)
    pLegItemAdj.Layer = pLayerAdj
    pLegItemAdj.LegendClassFormat = pLegItemClassFormatSite
    
    pLegend.AddItem(pLegItemSite)
    pLegend.AddItem(pLegItemAdj)

    pUnkLegendFormat = pFact.Create(CLSID(comtypes.gen.esriCarto.LegendFormat))
    pLegendFormat = CType(pUnkLegendFormat, comtypes.gen.esriCarto.ILegendFormat)

        # If Name and Description are greater than 28 characters, then adjust Legend spacing
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pLegendFormat.VerticalItemGap = 2.5
        pLegendFormat.VerticalPatchGap = 2.5
        pLegendFormat.LayerNameGap = 1
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pLegendFormat.VerticalItemGap = 2.5
        pLegendFormat.VerticalPatchGap = 2.5
        pLegendFormat.LayerNameGap = 1	
    else:
        pLegendFormat.VerticalItemGap = 5
        pLegendFormat.VerticalPatchGap = 5
    pLegend.Format = pLegendFormat

    # Add legend to Map Surround object
    pMSFrame.MapSurround = pLegend

    # Create Map Surround element and assign geometry
    pMSElement = CType(pMSFrame, comtypes.gen.esriCarto.IElement)
    pMSElement.Geometry = pEnvelope

    # Add Map Surround element (containing legend) to the Graphics Container (so it displays on the layout) 
    pGContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    pGContainer.AddElement(pMSElement, 0)

    # Refresh it
    iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    pAV.PartialRefresh(iOpt, None, None)

    
    # CODE FOR LEGEND TITLE TEXT
    # Set text symbol
    pUnkLegendTitleTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pLegendTitleTextSymbol = CType(pUnkLegendTitleTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pLegendTitleTextSymbol.Color = pBlack

    # If Name and Description are greater than 28 characters, then adjust Legend title position and size  
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pLegendTitleTextSymbol.Size = 10   
        dXminLegendTitle = dXminLegend - 0.25    # go ahead and adjust legend title position
        dYminLegendTitle = dYmaxLegend + 0.30   
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pLegendTitleTextSymbol.Size = 10 
        dXminLegendTitle = dXminLegend - 0.25
        dYminLegendTitle = dYmaxLegend + 0.30   
    else:
        pLegendTitleTextSymbol.Size = 13   
        dXminLegendTitle = dXminLegend
        dYminLegendTitle = dYmaxLegend + 0.25 
    pLegendTitleFTextSymbol = CType(pLegendTitleTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)
    pLegendTitleFTextSymbol.TypeSetting = True	 
    
    # Set legend title text location 
    dXmaxLegendTitle = dXmaxLegend - 2.75     # all relative to the position of the Legend element
    dYmaxLegendTitle = dYmaxLegend + 0.55 
    #dXminLegendTitle = dXminLegend #+ 0.05    #- 1.65
    #dYminLegendTitle = dYmaxLegend + 0.25     #dYminLegend - 0.05

    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminLegendTitle, dYminLegendTitle, dXmaxLegendTitle, dYmaxLegendTitle)
 
    # Create text element
    pUnkLegendTitleText = pFact.Create(CLSID(comtypes.gen.esriCarto.TextElement))
    pLegendTitleText = CType(pUnkLegendTitleText, comtypes.gen.esriCarto.ITextElement)
    pLegendTitleText.Symbol = pLegendTitleTextSymbol
    # insert text from label fields (KeyNum. + GPIN)
    pLegendTitleText.Text = '<BOL><FNT name="Arial">' + "Legend" + '</FNT></BOL>' 

    pLegendTitle = CType(pLegendTitleText, comtypes.gen.esriCarto.IElement)
    pLegendTitle.Geometry = pEnvelope

    # Create the graphics container
    pLegendTitleContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)

    # Add the text element to the graphics container
    pLegendTitleContainer.AddElement(pLegendTitle, 0)

    # Do a partial refresh
    #pGCSel = CType(pMap, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    #pGCSel.SelectElement(pLegendTitle)
    #iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    #pAV.PartialRefresh(iOpt, None, None)    
    clearRefresh() 
    
    # CODE FOR CASE NUMBER TEXT
    # Set text symbol
    pUnkCaseNumTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pCaseNumTextSymbol = CType(pUnkCaseNumTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pCaseNumTextSymbol.Color = pBlack
    
    # Set case number text location - 1
    dXmaxCaseNum = dXmaxRect - 2.00       # all relative to the position of the rectangular box    #dXmaxLegendTitle + 2.0   #dXmaxLegend - 1.00
    dYmaxCaseNum = dYmaxRect - 0.15       #dYmaxLegendTitle - 0.15    #+ 0.15   #dYmaxLegend - 0.25 
    dXminCaseNum = dXminRect + 2.00       #dXminLegendTitle + 2.0   #dXminLegend - 1.65
    dYminCaseNum = dYminRect + 0.90       #dYminLegendTitle    #dYminLegend - 0.05    
    
    # If Name and Description are greater than 28 characters, then adjust Case Number text size and location
    if len(str(arcpy.GetParameterAsText(3))) > 28:    # Name
        pCaseNumTextSymbol.Size = 10   
        dXminCaseNum = dXminCaseNum + 0.25
    elif len(str(arcpy.GetParameterAsText(4))) > 28:    # Description
        pCaseNumTextSymbol.Size = 10 
        dXminCaseNum = dXminCaseNum + 0.25
    else:
        pCaseNumTextSymbol.Size = 13
        dXminCaseNum = dXminRect + 2.00	
    pCaseNumFTextSymbol = CType(pCaseNumTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)
    pCaseNumFTextSymbol.TypeSetting = True	# allows for tagging of fonts, etc. 
    
    # Re-position Case Number text according to existence/length of Name/Description text
    # if Name and Description text are BOTH empty
    if str(arcpy.GetParameterAsText(3)) == "" and str(arcpy.GetParameterAsText(4)) == "":     # also means font size = 13
        dYmaxCaseNum = dYmaxCaseNum - 0.25    # reduce box YMax by 2 lines
    # if only Description text is empty (assuming Description text will not exist without Name text)
    elif str(arcpy.GetParameterAsText(3)) != "" and str(arcpy.GetParameterAsText(4)) == "":
        dYmaxCaseNum = dYmaxCaseNum - 0.15    # reduce box YMax by 1 line

    # Set case number text location - 2
    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminCaseNum, dYminCaseNum, dXmaxCaseNum, dYmaxCaseNum)
 
    # Create text element
    pUnkCaseNumText = pFact.Create(CLSID(comtypes.gen.esriCarto.TextElement))
    pCaseNumText = CType(pUnkCaseNumText, comtypes.gen.esriCarto.ITextElement)
    pCaseNumText.Symbol = pCaseNumTextSymbol
    # insert text from first user input box (case number)
    pCaseNumText.Text = '<BOL><FNT name="Arial">' + arcpy.GetParameterAsText(1) + '</FNT></BOL>' 

    pCaseNum = CType(pCaseNumText, comtypes.gen.esriCarto.IElement)
    pCaseNum.Geometry = pEnvelope

    # Create the graphics container
    pCaseNumContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)

    # Add the text element to the graphics container
    pCaseNumContainer.AddElement(pCaseNum, 0)

    # Do a partial refresh
    #pGCSel = CType(pMap, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    #pGCSel.SelectElement(pCaseNum)
    #iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    #pAV.PartialRefresh(iOpt, None, None)   
    clearRefresh()
    
    # CODE FOR NAME TEXT (OPTIONAL)
    # Set text symbol
    pUnkNameTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pNameTextSymbol = CType(pUnkNameTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pNameTextSymbol.Color = pBlack  
    
    # If Name and/or Description is greater than 28 characters, then adjust Name text size 
    if len(str(arcpy.GetParameterAsText(3))) > 28:    # Name
        pNameTextSymbol.Size = 10 
        #dXminName = dXminCaseNum  + 0.25
        #pythonaddins.MessageBox(len(arcpy.GetParameterAsText(4)), 0)
        #pythonaddins.MessageBox(len(arcpy.GetParameterAsText(4)), 0)
    elif len(str(arcpy.GetParameterAsText(4))) > 28:    # Description
        pNameTextSymbol.Size = 10     
        #dXminName = dXminCaseNum  + 0.25
        #pythonaddins.MessageBox(len(arcpy.GetParameterAsText(4)), 0)
        #pythonaddins.MessageBox(len(arcpy.GetParameterAsText(4)), 0)
    else:
        pNameTextSymbol.Size = 13
        dXmaxName = dXmaxCaseNum  
    pNameFTextSymbol = CType(pNameTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)
    pNameFTextSymbol.TypeSetting = True	  # allows for tagging of fonts, etc. 

    # Set name text location
    dXmaxName = dXmaxCaseNum               # all relative to the position of the Case Number Text element    #dXmaxLegend - 1.00    
    dYmaxName = dYmaxCaseNum - 0.20        #dYmaxLegend - 0.25 
    dXminName = dXminCaseNum               #dXminLegend - 1.65
    dYminName = dYminCaseNum - 0.20        #dYminLegend - 0.05
    
    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminName, dYminName, dXmaxName, dYmaxName)
 
    # Create text element
    pUnkNameText = pFact.Create(CLSID(comtypes.gen.esriCarto.TextElement))
    pNameText = CType(pUnkNameText, comtypes.gen.esriCarto.ITextElement)
    pNameText.Symbol = pNameTextSymbol
    # insert text from first user input box (case number)
    pNameText.Text = '<BOL><FNT name="Arial">' + arcpy.GetParameterAsText(3) + '</FNT></BOL>' 

    pName = CType(pNameText, comtypes.gen.esriCarto.IElement)
    pName.Geometry = pEnvelope

    # Create the graphics container
    pNameContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)

    # Add the text element to the graphics container
    pNameContainer.AddElement(pName, 0)

    # Do a partial refresh
    #pGCSel = CType(pMap, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    #pGCSel.SelectElement(pName)
    #iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    #pAV.PartialRefresh(iOpt, None, None)      
    clearRefresh() 
    # if len(str(arcpy.GetParameterAsText(2 and 3)) == 0:
    # MOVE DOWN YMAX OF NAME, NORTH ARROW, AND RECTANGLE ONLY
    
    #     dXmaxName = dXmaxName             # all relative to the position of the Case Number Text element     #dXmaxLegend - 1.00
    #dYmaxDescription = dYmaxName - 0.20      #dYmaxLegend - 0.25 
    #dXminDescription = dXminName             #dXminLegend - 1.65
    #dYminDescription = dYminName - 0.20      #dYminLegend - 0.05  
    
    
    # CODE FOR DESCRIPTION TEXT (OPTIONAL)
    # Set text symbol
    pUnkDescriptionTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pDescriptionTextSymbol = CType(pUnkDescriptionTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pDescriptionTextSymbol.Color = pBlack  

    # If Name and/or Description is greater than 28 characters, then adjust Description text size
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pDescriptionTextSymbol.Size = 10
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pDescriptionTextSymbol.Size = 10
    else:
        pDescriptionTextSymbol.Size = 13 
    pDescriptionFTextSymbol = CType(pDescriptionTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)
    pDescriptionFTextSymbol.TypeSetting = True	  # allows for tagging of fonts, etc. 
    
    # Set description text location 
    dXmaxDescription = dXmaxName             # all relative to the position of the Case Number Text element     #dXmaxLegend - 1.00
    dYmaxDescription = dYmaxName - 0.20      #dYmaxLegend - 0.25 
    dXminDescription = dXminName             #dXminLegend - 1.65
    dYminDescription = dYminName - 0.20      #dYminLegend - 0.05

    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminDescription, dYminDescription, dXmaxDescription, dYmaxDescription)
 
    # Create text element
    pUnkDescriptionText = pFact.Create(CLSID(comtypes.gen.esriCarto.TextElement))
    pDescriptionText = CType(pUnkDescriptionText, comtypes.gen.esriCarto.ITextElement)
    pDescriptionText.Symbol = pDescriptionTextSymbol
    # insert text from first user input box (case number)
    pDescriptionText.Text = '<BOL><FNT name="Arial">' + arcpy.GetParameterAsText(4) + '</FNT></BOL>' 

    pDescription = CType(pDescriptionText, comtypes.gen.esriCarto.IElement)
    pDescription.Geometry = pEnvelope

    # Create the graphics container
    pDescriptionContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)

    # Add the text element to the graphics container
    pDescriptionContainer.AddElement(pDescription, 0)

    # Do a partial refresh
    #pGCSel = CType(pMap, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    #pGCSel.SelectElement(pDescription)
    #iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    #pAV.PartialRefresh(iOpt, None, None)   
    clearRefresh() 
    
    # CODE FOR SCALE BAR
    # Get scale bar style from the Style Gallery
    pSBStyleGallery = pMxDoc.StyleGallery
    pEnumSBStyleGallery = pSBStyleGallery.Items("Scale Bars", "ESRI.style", "DoubleAlternatingScaleBar")
    pEnumSBStyleGallery.Reset()
    pScaleBarStyle = pEnumSBStyleGallery.Next()
    scalebarList = []    # list to hold the scale bars
    for i in range(0, 3):    # there are four Double Alternating Scale Bars in the list
        if pScaleBarStyle is not None:
            if pScaleBarStyle.Name == "Double Alternating Scale Bar 1":    # get the scale bar style with this name
                scalebarList.append(pScaleBarStyle)
            pScaleBarStyle = pEnumSBStyleGallery.Next()
    #pythonaddins.MessageBox(scalebarList, 0)

    # Set text symbol for scale bar Labels
    pUnkScaleBarTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    pScaleBarTextSymbol = CType(pUnkScaleBarTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    pScaleBarTextSymbol.Color = pBlack

    # If Name and Description are greater than 28 characters, then adjust Scale Bar text size    
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pScaleBarTextSymbol.Size = 7   
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pScaleBarTextSymbol.Size = 7 
    else:
        pScaleBarTextSymbol.Size = 8     

    pScaleBarFTextSymbol = CType(pScaleBarTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)   # allows for font/etc tags
    pScaleBarFTextSymbol.TypeSetting = True
    
    # Create scale bar using selected style
    pScaleBarStyle = scalebarList[0]
    pScaleBar = CType(pScaleBarStyle.Item, comtypes.gen.esriCarto.IScaleBar)
    pScaleBar.UseMapSettings() # use map to determine scale of scale bar
    pScaleBar.Units = 3    # Feet
    pScaleBar.BarHeight = 5    # 1 point = 1/72 inch
    pScaleBar.LabelSymbol = pScaleBarTextSymbol	   # symbology for labels
    pScaleBar.UnitLabelSymbol = pScaleBarTextSymbol    # symbology for unit label (Feet)
    
    # Create container for scale bar
    pUnkScaleBarFrame = pFact.Create(CLSID(comtypes.gen.esriCarto.MapSurroundFrame))
    pScaleBarFrame = CType(pUnkScaleBarFrame, comtypes.gen.esriCarto.IMapSurroundFrame)
    pScaleBarContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    #pGContainer.AddElement(pMSFrame, 0)

    # Set scale bar location    
    dXmaxScaleBar = dXmaxRect - 0.08    # all relative to the position of rectangular box    #pEnv.XMax - 0.75
    dYmaxScaleBar= dYminRect + 0.07    #pEnv.YMax - 5.25  
    dXminScaleBar = dXminRect + 2.00    #pEnv.XMax - 5.13
    dYminScaleBar = dYminRect + 0.05    #pEnv.YMax - 8.50

    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminScaleBar, dYminScaleBar, dXmaxScaleBar, dYmaxScaleBar)

    # Add scale bar to Map Surround object
    pScaleBarFrame.MapSurround = pScaleBar

    # Create Map Surround element and assign geometry
    pScaleBarElement = CType(pScaleBarFrame, comtypes.gen.esriCarto.IElement)
    pScaleBarElement.Geometry = pEnvelope

    # Add Map Surround element (containing scale bar) to the Graphics Container (so it displays on the layout) 
    pScaleBarContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    pScaleBarContainer.AddElement(pScaleBarElement, 0)

    # Refresh it
    iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    pAV.PartialRefresh(iOpt, None, None)
   
    
    
    # CODE FOR NORTH ARROW
    # Get north arrow style from the Style Gallery
    pNAStyleGallery = pMxDoc.StyleGallery
    pEnumNAStyleGallery = pNAStyleGallery.Items("North Arrows", "ESRI.style", "Default")
    pEnumNAStyleGallery.Reset()
    pNorthArrowStyle = pEnumNAStyleGallery.Next()
    northarrowList = []    # list to hold the scale bars
    for i in range(0, 96):    # there are 97 Default North Arrows in the list
        if pNorthArrowStyle is not None:
            if pNorthArrowStyle.Name == "ESRI North 8":    # get the north arrow style with this name
                northarrowList.append(pNorthArrowStyle)
            pNorthArrowStyle = pEnumNAStyleGallery.Next()
    #pythonaddins.MessageBox(scalebarList, 0)

    # # Set text symbol for north arrow Labels
    # pUnkScaleBarTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
    # pScaleBarTextSymbol = CType(pUnkScaleBarTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
    # pScaleBarTextSymbol.Color = pBlack
    # pScaleBarTextSymbol.Size = 8    
    # pScaleBarFTextSymbol = CType(pScaleBarTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)   # allows for font/etc tags
    # pScaleBarFTextSymbol.TypeSetting = True
    
    # Create north arrow using selected style
    pNorthArrowStyle = northarrowList[0]
    pNorthArrow = CType(pNorthArrowStyle.Item, comtypes.gen.esriCarto.INorthArrow)

    # If Name and Description are greater than 28 characters (or if one does not exist), then adjust North Arrow size and position
    if len(str(arcpy.GetParameterAsText(3))) > 28:
        pNorthArrow.Size = 45    # size in points   
        dXminNorthArrow = dXmaxRect - 0.55
        dYminNorthArrow = dYminRect + 0.65 	
    elif len(str(arcpy.GetParameterAsText(4))) > 28:
        pNorthArrow.Size = 45    # size in points 
        dXminNorthArrow = dXmaxRect - 0.55
        dYminNorthArrow = dYminRect + 0.65 
    elif str(arcpy.GetParameterAsText(3)) == "" or str(arcpy.GetParameterAsText(4)) == "":
        pNorthArrow.Size = 45    # size in points 
        dXminNorthArrow = dXmaxRect - 0.55
        dYminNorthArrow = dYminRect + 0.65         
    else:
        pNorthArrow.Size = 60    # size in points
        dXminNorthArrow = dXmaxRect - 0.75
        dYminNorthArrow = dYminRect + 0.50 
    
    # Create container for north arrow
    pUnkNorthArrowFrame = pFact.Create(CLSID(comtypes.gen.esriCarto.MapSurroundFrame))
    pNorthArrowFrame = CType(pUnkNorthArrowFrame, comtypes.gen.esriCarto.IMapSurroundFrame)
    pNorthArrowContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    #pGContainer.AddElement(pMSFrame, 0)
	
    # Set north arrow location    
    dXmaxNorthArrow = dXmaxRect - 0.08    # all relative to the position of rectangular box    #pEnv.XMax - 0.75
    dYmaxNorthArrow= dYminRect + 1.00
    
    # Move North Arrow down according to existence/length of Name/Description text
    # if Name and Description text are BOTH empty
    if str(arcpy.GetParameterAsText(3)) == "" and str(arcpy.GetParameterAsText(4)) == "":     # also means font size = 13
        dYmaxNorthArrow = dYmaxNorthArrow - 0.57    # move north arrow down 2 lines
    # if only Description text is empty (assuming Description text will not exist without Name text)
    elif str(arcpy.GetParameterAsText(3)) != "" and str(arcpy.GetParameterAsText(4)) == "":
        dYmaxNorthArrow = dYmaxNorthArrow - 0.285    # move north arrow down 1 line

    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
    pEnvelope.PutCoords(dXminNorthArrow, dYminNorthArrow, dXmaxNorthArrow, dYmaxNorthArrow)

    # Add north arrow to Map Surround object
    pNorthArrowFrame.MapSurround = pNorthArrow

    # Create Map Surround element and assign geometry
    pNorthArrowElement = CType(pNorthArrowFrame, comtypes.gen.esriCarto.IElement)
    pNorthArrowElement.Geometry = pEnvelope

    # Add Map Surround element (containing north arrow) to the Graphics Container (so it displays on the layout) 
    pNorthArrowContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    pNorthArrowContainer.AddElement(pNorthArrowElement, 0)

    # Refresh it
    iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
    pAV.PartialRefresh(iOpt, None, None)
       
    
    #pGCSel.UnselectAllElements()
    clearRefresh()   
    arcpy.RefreshActiveView()

 
# LABELING
def labelParcels():

    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
    
    # CODE TO REFRESH GRAPHICS EACH TIME(WITHOUT DELETING THE DATA FRAME)    
    # Loop through and delete all graphics elements (everything except the data frame)    
    pGCSel = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
    pGCSel.SelectAllElements()
    elementCount = pGCSel.ElementSelectionCount-1    # get count of all elements in Page Layout    # (excluding data frame, -1)
    
    pGraphicsContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
    pGraphicsContainer.Reset()
    pElem = pGraphicsContainer.Next()  # each element
    #pElemReport = []    # report of names and types of all elements
    pElemList = []    # list of all elements (minus the data frame element, which we want to keep)
    for i in range(0, elementCount):
        if pElem is not None:
            pElemProp = CType(pElem, comtypes.gen.esriCarto.IElementProperties)
            if pElemProp.Type != "Data Frame":    # delete all elements except the data frame
                #pElemReport.append("Name: " + pElemProp.Name + ", Type: " + pElemProp.Type)
                pElemList.append(pElem)   # append each element to the list of elements
                pElem = pGraphicsContainer.Next() # go to the next element 
        i = i + 1	
    #pythonaddins.MessageBox(pElemReport, 0)    # report the Name and Type properties of each element    
    for i in range(0, elementCount):
        pGraphicsContainer.DeleteElement(pElemList[i])    # delete all elements, except the data frame
    pGCSel.UnselectAllElements()
    pMxDoc.ActiveView.Refresh()	
    
    
    
    # SET INITIAL GRAPHICS ELEMENT PROPERTIES
    # Colors
    pUnkBlack = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))    # black
    pBlack = CType(pUnkBlack, comtypes.gen.esriDisplay.IRgbColor)
    pBlack.Red = 0
    pBlack.Green = 0
    pBlack.Blue = 0	    

    pUnkWhite = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))    # white
    pWhite = CType(pUnkWhite, comtypes.gen.esriDisplay.IRgbColor)
    pWhite.Red = 255    #333
    pWhite.Green = 255    #111
    pWhite.Blue = 255    #222	    
    
    # Set fill and line color (see COLORS above) (for Label Legend and Legend Title Text boxes   # USE FOR LABEL LEGEND AND RECTANGLE BOX ONLY; THE REST YOU DON'T NEED IT)   
    # Create line symbol with outline color    
    pUnkSimpleLineSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleLineSymbol))
    pSimpleLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ISimpleLineSymbol)
    pLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ILineSymbol)
    
    pLineSymbol.Color = pWhite
    pLineSymbol.Width = 0	 
    
    pUnkSymbolBorder = pFact.Create(CLSID(comtypes.gen.esriCarto.SymbolBorder))
    pSymbolBorder = CType(pUnkSymbolBorder, comtypes.gen.esriCarto.ISymbolBorder)
    pSymbolBorder.LineSymbol = pLineSymbol	
    
    # Create fill symbol with fill color    
    pUnkSimpleFillSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleFillSymbol))
    pSimpleFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.ISimpleFillSymbol)
    pFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.IFillSymbol)
    
    pSimpleFillSymbol.Color = pWhite
    pSimpleFillSymbol.Outline = pLineSymbol
    
    pUnkSymbolBackground = pFact.Create(CLSID(comtypes.gen.esriCarto.SymbolBackground))		
    pSymbolBackground = CType(pUnkSymbolBackground, comtypes.gen.esriCarto.ISymbolBackground)
    pSymbolBackground.FillSymbol = pSimpleFillSymbol
    
    
    for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
        if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel":
            lyrName = lyr.name
	    
	    # zoom to full extent of map (page) layout
            pPageLayout = pMxDoc.PageLayout
	    pPageLayout.ZoomToWhole()
	    pMxDoc.ActiveView.Refresh()
	    

            # get list of GPINs to set KeyLabel values and for GetCount (count of GPINs)
            layersToGet = getMasterCondoLayer("adjacent")
            masterGPINlist = getMasterCondo(layersToGet, master=False)
            list1 = masterGPINlist[0]    # ALL GPINs
            list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed
	
            # get ratio of smallest area polygon (parcel) to largest area polygon (parcel); used to tell ArcMap when to implement key numbering
            parcelAreaField = [f.name for f in arcpy.ListFields(lyrName, "*STAr*")]
            area = [row[0] for row in arcpy.da.SearchCursor(lyrName, parcelAreaField)]
            minArea = min(area)
            maxArea = max(area)
            ratioArea = minArea/maxArea
	

            # If number of parcels > 15 OR if ratio of smallest to largest parcel is less than 0.15, then do KEY NUMBERING
            if len(list2) > 15 and df.scale > 1900 or ratioArea < 0.05:
	    
                # if contains no master gpins or condos, but contains other cama/master gpins with .000, 0.001, etc	    
                # if contains no other cama/master gpins with .000, 0.001, etc
                # Create new fields for labeling (key numbering) - see function above	    
                createLabelFields()
   
                # set Labeling to "Adjacent Parcels", KeyLabel field
                clearRefresh()	    
                arcpy.RefreshActiveView()
                enableMaplex()
                setLabelQualityBest()
                removeDuplicates()    # remove duplicate GPIN labels and set other Maplex labeling properties   
			    
                arcpy.RefreshActiveView() 
                arcpy.RefreshTOC()
		

                #if len(masterGPINlist[0]) == 0:	     # if condos are not present		

                    # # set Labeling to "Adjacent Parcels", KeyLabel field
                    # clearRefresh()	    
                    # enableMaplex()
                    # setLabelQualityBest()
                    # removeDuplicates()    # remove duplicate GPIN labels and set other Maplex labeling properties  
		    # clearRefresh()
		
		
                # get list of master condos --- REMOVE ---
                # layersToGet = getMasterCondoLayer("adjacent")
                # masterGPINlist = getMasterCondo(layersToGet, master=True)		
                # if len(masterGPINlist[0]) > 0:                         # if condos exist
                    # removeDuplicates()
                # else:      # --- KEEP ---
                    # # Re-sort by Latitude (or KeyLabel), both should give an equivalent result
                    # # get Adjacent Parcels layer
                    # layerSrc = ""  
                    # if arcpy.Exists(Adjacents_shp):
                        # layerSrc = Adjacents_shp
                    # else:
                        # layerSrc = lyr		
                    # # sort by Latitude (highest to lowest)
                    # arcpy.Sort_management(layerSrc, Adjacents_sort, [["Latitude", "DESCENDING"]])
                    # arcpy.Delete_management(layerSrc)
                    # arcpy.CopyFeatures_management(Adjacents_sort, layerSrc)
                    # arcpy.Delete_management(Adjacents_sort)		
		
		
                # Label using KeyLabel field
                layer = arcpy.mapping.ListLayers(mxd, "*")[1]
                layer.labelClasses[0].expression="[KeyLabel]"    # large parcels
		layer.labelClasses[1].expression="[KeyLabel]"    # small parcels
                layer.showLabels = True
                clearRefresh()
                arcpy.RefreshActiveView()
	    
	    
	    
                # Set up Label Legend (key numbering)
                # use SearchCursor to get each GPIN (one per parcel; duplicates removed) and append to it a number (ex., "1. 555-555-5555"), and export each number+GPIN as a string
                numbering = [row[0] for row in arcpy.da.SearchCursor(lyrName, "KeyLabel", "\"LabelMain\" = 1")]
                gpin = [row[0] for row in arcpy.da.SearchCursor(lyrName, "GPIN", "\"LabelMain\" = 1")]
                labelLegend = list()
                labelLegend = []
                for i in range(0,len(gpin)):
                    labelLegend.append(str(numbering[i]) + ". " + str(gpin[i]))
            
                # create a TextElement (clone a current one) and then add each number+GPIN string to it (one per line with columns)
                # adjust formatting (spacing, rows, columns; halo 1.0, position, background, etc.)... ARCOBJECTS
                #legendname = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[0]   #1
                #northArrow = arcpy.mapping.ListLayoutElements(mxd, "MAPSURROUND_ELEMENT")[0]
                #caseNum = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[4] 

                #keyNum = legendname.clone()
                #keyNum.elementPositionY = northArrow.elementPositionY + 7.10
                #keyNum.elementPositionX = northArrow.elementPositionX - 1.10

                #for elm in arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT"):
                #    if elm.name == arcpy.GetParameterAsText(4):
                #        keyNum.elementPositionX = elm.elementPositionX + 2.25
                #arcpy.RefreshActiveView()	

                #keyNum.text = str(labelLegend)
                labelLegendstr = str(labelLegend)
                labelLegendstr = labelLegendstr.replace(",","")
                labelLegendstr = labelLegendstr.replace("[","")
                labelLegendstr = labelLegendstr.replace("]","")
                labelLegendstr = labelLegendstr.replace("'","")

	    
                #keyNum.text = '\r\n'.join(textwrap.wrap(labelLegendstr, 16))

                #font, etc. formatting
                #also, check to make sure caseNum has correct element number [3] or [2]?
		

                # Set text symbol
                pUnkLabelLegendTextSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.TextSymbol))
                pLabelLegendTextSymbol = CType(pUnkLabelLegendTextSymbol, comtypes.gen.esriDisplay.ITextSymbol)
                pLabelLegendTextSymbol.Color = pBlack
                pLabelLegendTextSymbol.Size = 7
		
                # # Create the mask (DOESN'T SEEM TO WORK)
                # pMask = CType(pLabelLegendTextSymbol, comtypes.gen.esriDisplay.IMask)
                # pMask.MaskSize = 2.0
		
                pLabelLegendFTextSymbol = CType(pLabelLegendTextSymbol, comtypes.gen.esriDisplay.IFormattedTextSymbol)
                pLabelLegendFTextSymbol.TypeSetting = True		
		
                # Create paragraph text element and properties
                pUnkLabelLegendParText = pFact.Create(CLSID(comtypes.gen.esriCarto.ParagraphTextElement))
                pLabelLegendParText = CType(pUnkLabelLegendParText, comtypes.gen.esriCarto.IParagraphTextElement)

                # Count the number of active labels		 
                with arcpy.da.SearchCursor("Adjacent Parcels","KeyLabel","\"KeyLabel\" > 0") as cursor:
                    rows = {row[0] for row in cursor}
                 
                count = 0
                for row in rows:
                    count += 1	
                labelsCount = count    

                # set up margins and position based on number of parcels in list
                # Set the position
                # get midpoint of focus map
                pAV = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IActiveView)    # map layout extent (used to create Envelope for graphic element position)
                pSD = pAV.ScreenDisplay
                pEnv = pAV.Extent 
                #dXmin = pEnv.XMax - 3.75 
                dXmax = pEnv.XMax - 0.85
                #dYmin = pEnv.YMax - ((0.115*labelsCount)+0.85)     #4.25  
                dYmax = pEnv.YMax - 0.85  		

                # Set columns	
                pColumnProps = CType(pLabelLegendParText, comtypes.gen.esriCarto.IColumnProperties)	
                pColumnProps.Gap = 1	
        
                # For multiple columns:	
                if labelsCount > 40:
		
                    # Create list of length 2878 (increments of 40 up to 115,120 (slightly over current total # of parcels in the county), used to set column lengths of 40
                    interval = 40
                    thresholds = []
                    for i in range(2878):
                        thresholds.insert(i, interval)
                        interval += 40

                    # Iterate through the list of thresholds until you reach the one just below the current number of labels (this is the one used to determine column count),
                    # ex., for 45 labels, it would be 40;
                    # then get the index of that threshold in the list of thresholds     
                    thresholdList=[]    
                    for j in thresholds:
                        if labelsCount > j:
                            thresholdList.append(j)
			
                    # Generate a list of the number of columns corresponding to each threshold; 
                    # list indices for threshold and number of columns will match (ex., > 40 labels would be 2 columns)
                    minColumns = 2
                    numColumns = []
                    for i in range(len(thresholdList)):
                        numColumns.insert(i, minColumns)
                        minColumns += 1
		    # Set number of columns
                    pColumnProps.Count = numColumns[len(thresholdList)-1]

                    # Generate a list of the x-Min boundaries corresponding to each threshold; 
                    # list indices for threshold and x-Min boundaries will match (ex., > 40 labels would be an x-Min boundary of 2.825)   
                    minXmin = 2.825
                    xMin = []
                    for i in range(len(thresholdList)):
                        xMin.insert(i, minXmin)
                        minXmin += 1
		    # Set x-Min and y-Min
                    dXmin = pEnv.XMax - xMin[len(thresholdList)-1]
                    dYmin = pEnv.YMax - ((0.115*((labelsCount/pColumnProps.Count + 1))) + 0.85)    # adjust y-Min (bottom boundary) according to number of labels (reduces white space)  #((0.115*(40))+0.85) 		    
		
			
                # For single column (if number of labels <= 40):
                else:
                    pColumnProps.Count = 1		    
                    dXmin = pEnv.XMax - 1.875		    
                    dYmin = pEnv.YMax - ((0.115*labelsCount)+0.85)  		
          			    
		    
                # Set margins			    
                pMarginProps = CType(pLabelLegendParText, comtypes.gen.esriDisplay.IMarginProperties)
                pMarginProps.Margin = 1	
		
                # Set frame properties	 
                pFrameProps = CType(pLabelLegendParText, comtypes.gen.esriCarto.IFrameProperties)
                pFrameProps.Background = pSymbolBackground
                pFrameProps.Border = pSymbolBorder		
		
                # Create text element
                pUnkLabelLegendText = pFact.Create(CLSID(comtypes.gen.esriCarto.TextElement))
                pLabelLegendText = CType(pUnkLabelLegendText, comtypes.gen.esriCarto.ITextElement)
                pLabelLegendText = CType(pLabelLegendParText, comtypes.gen.esriCarto.ITextElement)
                pLabelLegendText.Symbol = pLabelLegendTextSymbol
                # insert text from label fields (KeyNum. + GPIN)
                pLabelLegendText.Text = '<BOL><FNT name="Arial">' + '\r\n'.join(textwrap.wrap(labelLegendstr, 17)) + '</FNT></BOL>'

                # remove previous label legends	(csv and dbf)
	        mailerPath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0) + "\\" + "Mailer_files"		
                if os.path.exists("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv"):
                    os.remove("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv")
                if os.path.exists("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.dbf"):
                    os.remove("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.dbf")
                if os.path.exists(str(mailerPath)+"\\labelLegend.csv"):
                    os.remove(str(mailerPath)+"\\labelLegend.csv")
                if os.path.exists(str(mailerPath)+"\\labelLegend.dbf"):
                    os.remove(str(mailerPath)+"\\labelLegend.dbf")
                # If Label Legend contains more than 2 columns
                # Create CSV and DBF files of Label Legend instead of displaying them on map	
                if pColumnProps.Count > 2:
                    # create graphics and then case folder and files (then add the CSV and DBF tables to it)
                    createGraphics()    # add the other graphics elements (must be done before createCase(), in order to preserve user-input for Create Layout parameters)
                    createCase()
		    
                    labelLegendstr = labelLegendstr.replace(". ",",")						
                    labelLegendstr = labelLegendstr.replace(" ","\n")		
                    labelLegendTXT = open("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv", "w")
                    labelLegendTXT.write(labelLegendstr)
                    labelLegendTXT.close()
		
                    # Prepend header row to csv file (code courtesy of... http://stackoverflow.com/questions/5287762/how-to-insert-a-new-line-before-the-first-line-in-a-file-using-python)
                    # save to appropriate case folder
                    def prepend(filename, data, bufsize=1<<15):
                        # backup the file
                        backupname = filename + os.extsep+'bak'
                        try: os.unlink(backupname) # remove previous backup if it exists
                        except OSError: pass
                        os.rename(filename, backupname)

                        # open input/output files,  note: outputfile's permissions lost
                        with open(backupname) as inputfile, open(filename, 'w') as outputfile:
                            # prepend
                            outputfile.write(data)
                            # copy the rest
                            buf = inputfile.read(bufsize)
                            while buf:
                                outputfile.write(buf)
                                buf = inputfile.read(bufsize)

                        # remove backup on success
                        try: os.unlink(backupname)
                        except OSError: pass
                    
                    prepend("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv", 'key_number, gpin\n')
                    # Convert csv to dbf format
                    arcpy.TableToDBASE_conversion("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv","W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum)

                # If Label Legend contains 2 columns or less, then display it on the map		
                else:
                    pUnkEnvelope = pFact.Create(CLSID(comtypes.gen.esriGeometry.Envelope))
                    pEnvelope = CType(pUnkEnvelope, comtypes.gen.esriGeometry.IEnvelope)
                    pEnvelope.PutCoords(dXmin, dYmin, dXmax, dYmax)
                    pLabelLegend = CType(pLabelLegendText, comtypes.gen.esriCarto.IElement)
                    pLabelLegend.Geometry = pEnvelope
		
                    #pLabelLegendProp = CType(pLabelLegend, comtypes.gen.esriCarto.IElementProperties)
                    #pLabelLegendProp.Name = "Label Legend"

                    # Create the graphics container (FOCUS MAP) ---  Use Focus Map for all Layout elements
                    pLabelContainer = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainer)
		    
                    # Add the text element to the graphics container
                    pLabelContainer.AddElement(pLabelLegend, 0)
		
                    clearRefresh()
                    createGraphics()    # add the other graphics elements (must be done before createCase(), in order to preserve user-input for Create Layout parameters)
		    
                    # unselect all graphics elements
                    pGCSelClr = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
                    pGCSelClr.UnselectAllElements()
		    
                    createCase()    # create case folder and files		    

                clearRefresh()
                arcpy.RefreshActiveView()
		
            else:
                # set Labeling to "Adjacent Parcels", GPIN field
                clearRefresh()
                arcpy.RefreshActiveView()
                enableMaplex()
                setLabelQualityBest()
                removeDuplicates()    # remove duplicate GPIN labels and set other Maplex labeling properties  	    
                layer = arcpy.mapping.ListLayers(mxd, "*")[1]
                layer.labelClasses[0].expression="[GPIN]"
                layer.showLabels = True
    
                clearRefresh()
                arcpy.RefreshTOC()	 	    
                arcpy.RefreshActiveView()
                createGraphics()    # add the other graphics elements (must be done before createCase(), in order to preserve user-input for Create Layout parameters)

                # unselect all graphics elements
                pGCSelClr = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
                pGCSelClr.UnselectAllElements()
		        
                createCase()    # create case folder and files


    
    
    
# # to loop through selection
# pGCSel = CType(pMxDoc.PageLayout, comtypes.gen.esriCarto.IGraphicsContainerSelect)
# pGCSel.SelectAllElements()
# clearRefresh()
# pGCSel.ElementSelectionCount

# pGraphicsContainer = CType(pMxDoc.ActiveView, comtypes.gen.esriCarto.IGraphicsContainer)
# pGraphicsContainer.Reset()
# pElem = pGraphicsContainer.Next()
# pElemProp = CType(pElem, comtypes.gen.esriCarto.IElementProperties)  
# pElemReport = []
# for i in range(0, 100):
     # pElemReport.append("Name: " + pElemProp.Name + ", Type: " + pElemProp.Type)
     # pGraphicsContainer.Reset()
     # pElem = pGraphicsContainer.Next()
# pythonaddins.MessageBox(pElemReport, 0)
# pMxDoc.ActiveView.Refresh()	



# pMxDoc.ActiveView.Selection.SelectAll() #???
# #pEnumElem = CType(pGCSel.SelectedElements.Next(), comtypes.gen.esriCarto.IEnumElement)
# #pEnumElem.Reset()
# pSelectedElem = pGCSel.SelectedElements.Next()
# pElemProps = CType(pSelectedElem, comtypes.gen.esriCarto.IElementProperties)
# elemList = []
# for i in range(0, pGCSel.ElementSelectionCount):    # for i in range(0,pGCSel.ElementSelectionCount) and if element TYPE not data frame    #while pSelectedElem != "":
    # elemList.append(pElemProps)
    # pElemProps = CType(pSelectedElem, comtypes.gen.esriCarto.IElementProperties)
    # pSelectedElem = pGCSel.SelectedElements.Next()

#----------------------------------------------------------------------
# Go to Page Layout view (make it the Active View)
#pMxDoc.ActiveView = pMxDoc.PageLayout
arcpy.mapping.MapDocument("current").activeView = "PAGE_LAYOUT"

# Set overwrite to True (allows revised Subject and Adjacents layers to be updated without having to create new layer files)
arcpy.env.overwriteOutput = 1
  
# Set the workspace
folderPath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0)
mailerPath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0) + "\\" + "Mailer_files"

if arcpy.Exists(str(mailerPath)):
    arcpy.env.workspace = str(mailerPath)
else:
    arcpy.CreateFolder_management(str(folderPath), "Mailer_files")
    arcpy.env.workspace = str(mailerPath)

# Go to Data View
#mxd.activeView = df

# Check to see if new case folder and files has already been created (for the first time); if not, then create them
for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
    if lyr.name == "Subject Parcels" or lyr.name == "Subject Parcel" or lyr.name == "Site":
        isCopy = lyr.dataSource                                                            # is this a copy of the original Site shapefile for this site?
       
        if "Case" not in isCopy and "case" not in isCopy:                                  # check to see if the directory containing the layers contains any variation of "case"; if not, then save it in memory
            Site_shp = str(mailerPath) + "\\Site.shp"                                           # Site parcel shapefile
            Site_add = str(mailerPath) + "\\Site_add.shp"                                       # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = str(mailerPath) + "\\Adjacents.shp"                                 # Adjacent parcel shapefile
            Adjacents_add = str(mailerPath) + "\\Adjacent_add.shp"                              # new Adjacent parcel shapefile with new Adjacent parcels to be added
            Adjacents_sort = str(mailerPath) + "\\Adjacent_sort.shp"                            # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
            Adjacents_sort2 = str(mailerPath) + "\\Adjacent_sort2.shp"                          # new Adjacent parcel shapefile with new Adjacent parcels sorted by Key Label (ascending)    
            Condos_add = str(mailerPath) + "\\Condo_add.shp"                                    # add all condos for each master GPIN
	    
        else:
            dirNew = lyr.workspacePath                               # if "case" is in the directory path, then save it to that directory
            Site_shp = str(dirNew) + "\\Site.shp"                         # Site parcel shapefile
            Site_add = str(dirNew) + "\\Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = str(dirNew) + "\\Adjacents.shp"               # Adjacent parcel shapefile
            Adjacents_add = str(dirNew) + "\\Adjacent_add.shp"            # new Adjacent parcel shapefile with new Adjacent parcels to be added
            Adjacents_sort = str(dirNew) + "\\Adjacent_sort.shp"          # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
            Adjacents_sort2 = str(dirNew) + "\\Adjacent_sort2.shp"        # new Adjacent parcel shapefile with new Adjacent parcels sorted by Key Label (ascending)
            Condos_add = str(dirNew) + "\\Condo_add.shp"                  # add all condos for each master GPIN

# Clear old files in memory
arcpy.Delete_management("in_memory")
arcpy.Delete_management(Adjacents_sort)
arcpy.Delete_management(Adjacents_sort2)


# Rename Site and Adjacents layers to 'Subject Parcel(s if < 1)' and 'Adjacent Parcels' (FOR LEGEND)
for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
    if lyr.name == "Site" or lyr.name == "Subject Parcel" or lyr.name == "Subject Parcels":
        # get GPINS of master condos
        # get all other GPINS
        # append them into one list
        # count the len of list...

        # get count of ALL main GPINs (no duplicates)
        layersToGet = getMasterCondoLayer("subject")
        masterGPINlist = getMasterCondo(layersToGet, master=False)
        list1 = masterGPINlist[0]    # ALL GPINs
        list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed     
        if len(list2) > 1:
            lyr.name = "Subject Parcels"
        else:
            lyr.name = "Subject Parcel"
	    
    if lyr.name == "Adjacents" or lyr.name == "Adjacent Parcel" or lyr.name == "Adjacent Parcels" :
        # get count of ALL main GPINs (no duplicates)
        layersToGet = getMasterCondoLayer("adjacent")
        masterGPINlist = getMasterCondo(layersToGet, master=False)
        list1 = masterGPINlist[0]    # ALL GPINs
        list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed  
        if len(list2) > 1:
            lyr.name = "Adjacent Parcels"
        else:
            lyr.name = "Adjacent Parcel"
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView() 




# Create new case folder and files
#createCase()
   
# get list of master condos --- GPIN's with CAMA_GPIN's containing ".000" in Subject or Adjacents and ALL USE_CODES's (condos and office condos) XXX AND USE_CODE = 698 (all CONDOS only, with a master (.000))
layersToGet = getMasterCondoLayer("all")
masterGPINlist = getMasterCondo(layersToGet, master=True)   

# ORDER OF FUNCTION CALLS HERE IS CRUCIAL TO PROPER LABELING (must be run in this order)
# If OMIT CONDOS box is checked then Omit Condos only
duplicatesExist = checkforDuplicates() 
condoCheck = arcpy.GetParameterAsText(2)
if str(condoCheck) == 'true':
    caseNum = arcpy.GetParameterAsText(1)     # get case number to use in createCase() and saveNewMXD()
    omitCondos()    # omit condos
    labelParcels()    # label main parcels and create case folder and files
    clearRefresh()
    arcpy.RefreshActiveView()	
	
# If not, then omit nothing, OR, if condos were omitted previously, then restore the omitted condos
elif str(condoCheck) == 'false':
    caseNum = arcpy.GetParameterAsText(1)     # get case number to use in createCase() and saveNewMXD()
    omitCondos()    # omit condos first, to label main parcels only	
    restoreCondos()    # restore condos to attribute tables
    labelParcels()    # label main parcels and create case folder and files
    clearRefresh()
    arcpy.RefreshActiveView()
    
    # # get list of GPINs to set KeyLabel values and for GetCount (count of GPINs)
    # layersToGet = getMasterCondoLayer("adjacent")
    # masterGPINlist = getMasterCondo(layersToGet, master=False)
    # list1 = masterGPINlist[0]    # ALL GPINs
    # list2 = sorted(set(list1), key=list1.index)    # GPIN list, with duplicates removed
    # # make KeyLabel numbering the same for each GPIN (used for labels), so only main parcels are labeled; used in conjunction with "LabelMain" field in removeDuplicates()		    
    # # use list1 (all GPINs), in case multiple records with the same GPIN exist (that are NOT condos)
    # for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
        # if lyr.name == "Adjacent Parcels" or lyr.name == "Adjacent Parcel":
            # for i in range(0, len(list2)):
                # arcpy.SelectLayerByAttribute_management(lyr,"NEW_SELECTION", "\"GPIN\" = '"+list2[i]+"'")

                # with arcpy.da.UpdateCursor(lyr, ('KeyLabel')) as cursor:
                    # for row in cursor:
                        # row[0] = i + 1 
                        # cursor.updateRow(row)	
        # arcpy.RefreshActiveView() 
        # arcpy.RefreshTOC()    
	    
# Clear all layout graphics (except data frame) - GENERATES ERROR, CLEARS EVERYTHING!
# pGraphics = CType(pMxDoc.FocusMap, comtypes.gen.esriCarto.IGraphicsContainer)
# pGraphics.Reset()
# pElement = pGraphics.Next()
# pElementProp = CType(pElement, comtypes.gen.esriCarto.IElementProperties)
# while (pElementProp is not None):
    # if (pElementProp.Type != "Data Frame"):
        # pGraphics.DeleteElement(pElement)
        
        # pAV = CType(pMap, comtypes.gen.esriCarto.IActiveView)
        # iOpt = comtypes.gen.esriCarto.esriViewGraphics + comtypes.gen.esriCarto.esriViewGraphicSelection
        # pAV.PartialRefresh(iOpt, None, None)
	
        # pGraphics = CType(pMxDoc.FocusMap, comtypes.gen.esriCarto.IGraphicsContainer)
        # pGraphics.Reset()
        # pElement = pGraphics.Next()
        # pElementProp = CType(pElement, comtypes.gen.esriCarto.IElementProperties)





  
# CREATE LAYOUT VIEW --------------------------------------------------------------------------

# Go to Layout View
arcpy.mapping.MapDocument("CURRENT").activeView = "PAGE_LAYOUT"

#ZOOM to extent of Adjacents layer + buffer (10% of the average of the current extent width and height)
#error if Adjacents layer does not exist (Please add Adjacents first.) -- see Tool Validator code
zoomAndCenter()
arcpy.RefreshActiveView()





    # #create legend
        # #title (Legend)
# #insert case number (if case number present, then insert...)
# caseNum = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[3]  #0 is Legend
# if arcpy.GetParameterAsText(1) != "":
    # caseNum.text = arcpy.GetParameterAsText(1)
# else:
    # caseNum.text = " "
    
# #insert name (if present, then insert...)
# name = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[2]
# if arcpy.GetParameterAsText(4) != "":
    # name.text = arcpy.GetParameterAsText(4)
    
# #otherwise, leave blank
# else:
    # name.text = " "

# #insert description (if present, then insert...)
# description = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[1]
# if arcpy.GetParameterAsText(4) != "":
    # description.text = arcpy.GetParameterAsText(4)
# #otherwise, leave blank
# else:
    # description.text = " "
    

# #description.text = arcpy.GetParameterAsText(4)
	# #insert north arrow, scale bar



# legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT", "Legend")[0]
# legendname = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")[0]
# box = arcpy.mapping.ListLayoutElements(mxd,"GRAPHIC_ELEMENT")[0]
# northArrow = arcpy.mapping.ListLayoutElements(mxd, "MAPSURROUND_ELEMENT")[0]

# #set legend title
# legendname.text = "Legend"

# #set graphics and text box heights based on number of lines of text
# if str(name.text) == " " and str(description.text) == " ":   # if both name and description fields are blank
    # box.elementHeight = 1.25
    # northArrow.elementHeight = 0.65
    # northArrow.elementPositionX = 10

    # northArrow.elementPositionY = 0.7
    # legendname.elementPositionY = 1.1   
    # caseNum.elementPositionY = 1.1
    # arcpy.RefreshActiveView()
    
# elif str(name.text) == " " and str(description.text) != " ":    # if just name is blank
    # box.elementHeight = 1.25
    # northArrow.elementHeight = 0.7
    
    # northArrow.elementPositionX = 10

    # northArrow.elementPositionY = 0.65  
    # legendname.elementPositionY = 1.1     
    # caseNum.elementPositionY = 1.1  
    # description.elementPositionY = caseNum.elementPositionY - 0.28    # caseNum - 0.28
    # arcpy.RefreshActiveView()
    
# elif str(description.text) == " " and str(name.text) != " ":       # if just description is blank
    # box.elementHeight = 1.25
    # northArrow.elementHeight = 0.7
    
    # northArrow.elementPositionX = 10   
    
    # northArrow.elementPositionY = 0.65
    # legendname.elementPositionY = 1.1   
    # caseNum.elementPositionY = 1.1
    # name.elementPositionY = caseNum.elementPositionY - 0.28    # caseNum - 0.28
    # arcpy.RefreshActiveView()
    
# else:                          # if none are blank, set values back to default
    # box.elementHeight = 1.7120999999988271
    # northArrow.elementHeight = 0.8819999999996071
    
    # northArrow.elementPositionX = 9.801400000000285
    
    # northArrow.elementPositionY = 0.8677000000006956
    # legendname.elementPositionY = 1.5627000000004045
    # caseNum.elementPositionY = 1.5342000000000553
    # name.elementPositionY = 1.2533000000003085
    # description.elementPositionY = 0.9722999999994499
    # arcpy.RefreshActiveView()
   
# copy label legend files to Documents folder for backup	
mailerPath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0) + "\\" + "Mailer_files"	 
if os.path.exists("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv"):		    
    shutil.copy("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.csv", str(mailerPath)+"\\labelLegend.csv")
if os.path.exists("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.dbf"):
    shutil.copy("W:\\GIS_Mailer\\Mailer_Cases_by_Case\\"+caseNum+"\\labelLegend.dbf", str(mailerPath)+"\\labelLegend.dbf")   
    
# Make all layers selectable, if they are not already
pSite = CType(pMap.Layer(0), comtypes.gen.esriCarto.IFeatureLayer)
pAdjacents = CType(pMap.Layer(1), comtypes.gen.esriCarto.IFeatureLayer)
pParcels = CType(pMap.Layer(2), comtypes.gen.esriCarto.IFeatureLayer)
pSite.Selectable = True
pAdjacents.Selectable = True
pParcels.Selectable = True
arcpy.RefreshTOC()
clearRefresh()
arcpy.RefreshActiveView()
    
# -----------------------------------------------------------------------------------------
    
# Save CURRENT map document	    
mxd.save()