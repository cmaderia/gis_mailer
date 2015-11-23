# Select Adjacent Parcels
#adjacentSelect.py - select and zoom to immediate adjacents, colorize them, and allow for selection of additional adjacent parcels


# IMPORT

# Import libraries
import arcpy
import comtypes
import comtypes.client
import os
import sys
from win32com.shell import shell, shellcon

# Import ArcObjects modules
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriGeometry.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriGeoDatabase.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesRaster.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriSystem.olb')

comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriCarto.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDisplay.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesGDB.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriDataSourcesFile.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriOutput.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriFramework.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriArcMapUI.olb')
comtypes.client.GetModule(r'C:\Program Files (x86)\ArcGIS\Desktop10.2\com\esriArcCatalogUI.olb')


# METHODS

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

# Get current ArcMap session
def GetApp(app="ArcMap"):
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
      pApp = pAppROT.Item(i)
      if app == "ArcCatalog":
          if CType(pApp, esriCatalogUI.IGxApplication):
              return pApp
          continue
      if CType(pApp, esriArcMapUI.IMxApplication):
          return pApp
  return None
  
  
# Set Adjacent parcel symbology
def setAdjacentSymbology():
    # Cast objects
    pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
    
    # Get Adjacents layer	
    pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
    pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
    pLayerSite = pMap.Layer(1)
    
    pUnkFeatureLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.FeatureLayer))
    pUnkFeatureLayer = pLayerSite
    pFeatureLayer = CType(pUnkFeatureLayer, comtypes.gen.esriCarto.IFeatureLayer)
    
    pGeoFeatureLayer = CType(pFeatureLayer, comtypes.gen.esriCarto.IGeoFeatureLayer)

    # Set parcel fill color    
    pUnkGreen = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))
    pGreen = CType(pUnkGreen, comtypes.gen.esriDisplay.IRgbColor)
    pGreen.Red = 211
    pGreen.Green = 255
    pGreen.Blue = 190

    # Set parcel outline color       
    pUnkGray = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))
    pGray = CType(pUnkGray, comtypes.gen.esriDisplay.IRgbColor)
    pGray.Red = 110
    pGray.Green = 110
    pGray.Blue = 110

    # Create line symbol with outline color    
    pUnkSimpleLineSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleLineSymbol))
    pSimpleLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ISimpleLineSymbol)
    pLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ILineSymbol)
    
    pLineSymbol.Color = pGray
    pLineSymbol.Width = 1.00

    # Create fill symbol with fill color    
    pUnkSimpleFillSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleFillSymbol))
    pSimpleFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.ISimpleFillSymbol)
    pFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.IFillSymbol)
    
    pSimpleFillSymbol.Color = pGreen
    pSimpleFillSymbol.Outline = pLineSymbol
    
    # Send it to the renderer    
    pUnkSimpleRenderer = pFact.Create(CLSID(comtypes.gen.esriCarto.SimpleRenderer))
    pSimpleRenderer = CType(pUnkSimpleRenderer, comtypes.gen.esriCarto.ISimpleRenderer)
    
    pSimpleRenderer.Symbol = pSimpleFillSymbol

    # Apply the renderer to the layer    
    pGeoFeatureLayer.Renderer = pSimpleRenderer

# Site = first layer; Adjacents = second layer	
def moveLayer():
    for lyr in arcpy.mapping.ListLayers(mxd,"",df):
      if lyr.name == "Adjacents":
        moveLayer = lyr
      if lyr.name == "Site":
        refLayer = lyr
    arcpy.mapping.MoveLayer(df,refLayer,moveLayer,"AFTER")	
    
# Clear selection and Refresh map
def clearRefresh():
    pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMap.ClearSelection()
    pMxDoc.ActiveView.Refresh()    
    
    
# VARIABLES

# Global variables:
tax_parcels = "Tax Parcels - Mailer"

# Initialize variables
Site_shp = ""                                                                  # Site parcel shapefile
Site_add = ""                                                                  # new layer shapefile with new Site parcels to be added 
Adjacents_shp = ""                                                             # Adjacent parcel shapefile
Adjacents_add = ""                                                             # new Adjacent parcel shapefile with new Adjacent parcels to be added
Adjacents_sort = ""                                                            # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
zoom_to_buffer = ""                                                            # buffer around Adjacent parcel(s) used to set zoom level

#----------------------------------------------------------------------


    
# Set the workspace
mailerPath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0) + "\\" + "Mailer_files"

if arcpy.Exists(mailerPath):
    arcpy.env.workspace = mailerPath
else:
    arcpy.CreateFolder_management(shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0), "Mailer_files")
    arcpy.env.workspace = mailerPath

# Get access to the current mxd 
mxd=arcpy.mapping.MapDocument("CURRENT") 
 
# Grab the dataframe object you want
df = arcpy.mapping.ListDataFrames(mxd,"*")[0]

#Go to Data View
mxd.activeView = df

# check to see if new case folder and files has already been created (for the first time); if not, then create them
for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
    if lyr.name == "Subject Parcels" or lyr.name == "Subject Parcel" or lyr.name == "Site":
        isCopy = lyr.dataSource                                                            # is this a copy of the original Site shapefile for this site?
       
        if "Case" not in isCopy and "case" not in isCopy:                                  # check to see if the directory containing the layers contains any variation of "case"; if not, then save it in memory
	    Site_shp = mailerPath + "\\Site.shp"                                           # Site parcel shapefile
            Site_add = mailerPath + "\\Site_add.shp"                                       # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = mailerPath + "\\Adjacents.shp"                                 # Adjacent parcel shapefile
	    Adjacents_add = mailerPath + "\\Adjacent_add.shp"                              # new Adjacent parcel shapefile with new Adjacent parcels to be added
	    Adjacents_sort = mailerPath + "\\Adjacent_sort.shp"                            # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
	    zoom_to_buffer = mailerPath + "\\zoom_buffer.shp"                              # buffer around Adjacent parcel(s) used to set zoom level
	    
	else:
	    dirNew = lyr.workspacePath                               # if "case" is in the directory path, then save it to that directory
            Site_shp = dirNew + "\\Site.shp"                         # Site parcel shapefile
            Site_add = dirNew + "\\Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = dirNew + "\\Adjacents.shp"               # Adjacent parcel shapefile
	    Adjacents_add = dirNew + "\\Adjacent_add.shp"            # new Adjacent parcel shapefile with new Adjacent parcels to be added
	    Adjacents_sort = dirNew + "\\Adjacent_sort.shp"          # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
	    zoom_to_buffer = dirNew + "\\zoom_buffer.shp"            # buffer around Adjacent parcel(s) used to set zoom level


# Get user input for parameters:  Add and Remove adjacent parcel check boxes (boolean, T/F)
removeCheck = arcpy.GetParameterAsText(0)
addCheck = arcpy.GetParameterAsText(1)



# NEW ADJACENTS
# If Add and Remove boxes are not checked,  proceed with automatic selection and addition of immediate adjacents; otherwise, Add or Remove selected parcels
if not (str(removeCheck) == 'true') and not (str(addCheck) == 'true'):

    # Clear any selections
    clearRefresh()  
    
    # Delete previous Adjacents shapefile and buffer
    arcpy.Delete_management(Adjacents_shp)
    arcpy.Delete_management(Adjacents_sort)    
    arcpy.Delete_management(zoom_to_buffer)

    # Remove previously selected Adjacents parcel shapefile from data frame
    for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
        if lyr.name == "Adjacents":
            arcpy.mapping.RemoveLayer(df, lyr)
    arcpy.RefreshActiveView()
    
    # Select immediate adjacents in Tax Parcel shapefile
    arcpy.SelectLayerByLocation_management("Tax Parcels - Mailer", "BOUNDARY_TOUCHES", "Site")
     
    # Write the selected features to a new feature class (shapefile) and clear selection from Tax Parcels layer
    arcpy.CopyFeatures_management(tax_parcels, Adjacents_shp, "", "0", "0", "0")
    clearRefresh() 

    # Add Adjacents shapefile to current map document, make it the second layer from the top, and select it
    adjacentParcel = arcpy.mapping.Layer(Adjacents_shp)
    arcpy.mapping.AddLayer(df, adjacentParcel, "TOP")
    moveLayer()
    
    # Set parcel color to Green -- see ArcObjects function above
    setAdjacentSymbology()
    
    # Set zoom level - 300 foot buffer around Adjacents
    arcpy.Buffer_analysis("Adjacents",zoom_to_buffer,"300 Feet","FULL","ROUND","NONE","#")
    bufferZoom = arcpy.mapping.Layer(zoom_to_buffer)
    ext = bufferZoom.getExtent()
    df.extent = ext
    
    # Clear selection (using ArcObjects) and refresh map
    pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMap.ClearSelection()
    pMxDoc.UpdateContents()
    pMxDoc.ActiveView.Refresh()    

    # Save MXD
    mxd.save()

    
# REMOVE
# If Remove box is checked, remove selected parcels
elif str(removeCheck) == 'true':
    # Execute Describe.fidset to get selected parcels, then execute Delete Features to remove the selected parcels.
    desc = arcpy.Describe("Adjacents")
    if int(str(len(desc.fidSet.split(";")))) > 0:
        arcpy.DeleteFeatures_management("Adjacents")

        # Clear selection and refresh map
        clearRefresh()
    # Save MXD	
    mxd.save()
    

# ADD
# If Add box is checked, add selected parcels    
elif str(addCheck) == 'true':

    # Delete previous layer used to hold previous "Add" parcels    
    arcpy.Delete_management(Adjacents_add)
    arcpy.Delete_management(Adjacents_sort)
    arcpy.Delete_management(zoom_to_buffer)
    
    # Some code to indicate if an Adjacents GPIN has already been selected and added to the layer; if it has, then remove the selection so it doesn't get added again
    
    # Run search cursor to get list of selected Tax Parcel GPINs
    taxCursor = arcpy.da.SearchCursor("Tax Parcels - Mailer", ["GPIN"])
    taxCursorlist = list(taxCursor)
    taxCursorlen = len(taxCursorlist)
    taxGPINlist = list()
    
    for i in range(0,taxCursorlen):
      taxCursoritem = taxCursorlist[i]
      taxCursoritemString = str(taxCursoritem)
      taxCursoritemStringRep1 = taxCursoritemString.replace("(u'","")
      taxCursoritemStringRep2 = taxCursoritemStringRep1.replace("',)","")
      taxGPINlist.append(taxCursoritemStringRep2)      
    
    # Run search cursor to get list of selected Adjacents GPINs
    adjacentsCursor = arcpy.da.SearchCursor("Adjacents", ["GPIN"])
    adjacentsCursorlist = list(adjacentsCursor)
    adjacentsCursorlen = len(adjacentsCursorlist)
    adjacentsGPINlist = list()
    
    for j in range(0,adjacentsCursorlen):
      adjacentsCursoritem = adjacentsCursorlist[j]
      adjacentsCursoritemString = str(adjacentsCursoritem)
      adjacentsCursoritemStringRep1 = adjacentsCursoritemString.replace("(u'","")
      adjacentsCursoritemStringRep2 = adjacentsCursoritemStringRep1.replace("',)","")
      adjacentsGPINlist.append(adjacentsCursoritemStringRep2)  
    
    # Generate set of Tax Parcel GPINs that already exist in Adjacents layer and deselect from Adjacents and Tax Parcels - Mailer layers	
    GPINset = set(adjacentsGPINlist).intersection(taxGPINlist)
    
    if len(GPINset) > 0:
      for k in range(0,len(list(GPINset))):
        expression = "GPIN = " +"'" + list(GPINset)[k] + "'"
        arcpy.SelectLayerByAttribute_management("Adjacents","REMOVE_FROM_SELECTION",expression)
        arcpy.SelectLayerByAttribute_management("Tax Parcels - Mailer","REMOVE_FROM_SELECTION",expression)
          
    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.
    # Add selected parcels from Tax Parcels layer to Adjacents_add layer  
    arcpy.CopyFeatures_management(tax_parcels, Adjacents_add, "", "0", "0", "0")
    
    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.
    # Add selected parcels from Tax Parcels layer to Adjacents_add layer  
    adjacents_add = arcpy.mapping.Layer(Adjacents_add)
    arcpy.mapping.AddLayer(df, adjacents_add, "TOP")
    
    # If Adjacents_add layer contains parcels, then append them to the Adjacents shapefile
    desc = arcpy.Describe(adjacents_add)
    if  int(str(len(desc.fidSet.split(";")))) > 0:
        arcpy.Append_management(adjacents_add,"Adjacents")
	
	# Remove the Adjacents_add layer (temporary layer to hold parcels to add)
        for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
            if lyr.name == "Adjacent_add":
                arcpy.mapping.RemoveLayer(df, lyr)
        arcpy.RefreshActiveView()
        arcpy.Delete_management(Adjacents_add)
	
	# Set zoom level - 300 foot buffer around Adjacents
        arcpy.Buffer_analysis("Adjacents",zoom_to_buffer,"300 Feet","FULL","ROUND","ALL","#")
        bufferZoom = arcpy.mapping.Layer(zoom_to_buffer)
        ext = bufferZoom.getExtent()
        df.extent = ext
    
	# Refresh Table of Contents and Clear selection
	arcpy.RefreshTOC()
        clearRefresh()
	
    # Save MXD	
    mxd.save()	
	
