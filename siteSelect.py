# Mailer Site Selection
# siteSelect.py - select site parcel, colorize, and zoom (700-foot buffer)

# IMPORT

# Import libraries
import arcpy
import pythonaddins
import win32ui, win32con, win32gui, ctypes
import comtypes
import comtypes.client
import os
import sys

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

# Methods to run ArcObjects (used for Clear Selection)
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
    #Return CLSID of MyClass as string
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
  
# Set Site parcel symbology
def setSiteSymbology():
    # Cast objects
    pFact = CType(pApp, comtypes.gen.esriFramework.IObjectFactory)
   
    # Get Site layer
    pUnkLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.LayerFile))
    pLayer = CType(pUnkLayer, comtypes.gen.esriCarto.ILayer)
    pLayerSite = pMap.Layer(0)
    
    pUnkFeatureLayer = pFact.Create(CLSID(comtypes.gen.esriCarto.FeatureLayer))
    pUnkFeatureLayer = pLayerSite
    pFeatureLayer = CType(pUnkFeatureLayer, comtypes.gen.esriCarto.IFeatureLayer)
    
    pGeoFeatureLayer = CType(pFeatureLayer, comtypes.gen.esriCarto.IGeoFeatureLayer)

    # Set parcel fill color
    pUnkRose = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))
    pRose = CType(pUnkRose, comtypes.gen.esriDisplay.IRgbColor)
    pRose.Red = 255
    pRose.Green = 190
    pRose.Blue = 190
    
    # Set parcel outline color    
    pUnkBlack = pFact.Create(CLSID(comtypes.gen.esriDisplay.RgbColor))
    pBlack = CType(pUnkBlack, comtypes.gen.esriDisplay.IRgbColor)
    pBlack.Red = 0
    pBlack.Green = 0
    pBlack.Blue = 0
    
    # Create line symbol with outline color
    pUnkSimpleLineSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleLineSymbol))
    pSimpleLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ISimpleLineSymbol)
    pLineSymbol = CType(pUnkSimpleLineSymbol, comtypes.gen.esriDisplay.ILineSymbol)
    
    pLineSymbol.Color = pBlack
    pLineSymbol.Width = 0.40
    
    # Create fill symbol with fill color
    pUnkSimpleFillSymbol = pFact.Create(CLSID(comtypes.gen.esriDisplay.SimpleFillSymbol))
    pSimpleFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.ISimpleFillSymbol)
    pFillSymbol = CType(pUnkSimpleFillSymbol, comtypes.gen.esriDisplay.IFillSymbol)
    
    pSimpleFillSymbol.Color = pRose
    pSimpleFillSymbol.Outline = pLineSymbol
    
    # Send it to the renderer
    pUnkSimpleRenderer = pFact.Create(CLSID(comtypes.gen.esriCarto.SimpleRenderer))
    pSimpleRenderer = CType(pUnkSimpleRenderer, comtypes.gen.esriCarto.ISimpleRenderer)
    
    pSimpleRenderer.Symbol = pSimpleFillSymbol
    
    # Apply the renderer to the layer
    pGeoFeatureLayer.Renderer = pSimpleRenderer
    
  
# VARIABLES

# Global variables:
tax_parcels = "Tax Parcels - Mailer"

site_rose = "W:/GIS_Mailer/new_mailer_toolbar/Site_color_rose.lyr"             # contains symbology (color) for Site parcel layer
adjacent_green = "W:/GIS_Mailer/new_mailer_toolbar/Adjacent_color_green.lyr"   # contains symbology (color) for Adjacent parcel layer
zoom_to_buffer = "W:/GIS_Mailer/new_mailer_toolbar/Site_buffer_700ft.shp"      # buffer around Site parcel(s) used to 

# Initialize variables
Site_shp = ""                                                                  # Site parcel shapefile
Site_add = ""                                                                  # new layer shapefile with new Site parcels to be added 
Adjacents_shp = ""                                                             # Adjacent parcel shapefile

#-------------------------------------------------------------------------------



# Set the workspace
arcpy.env.workspace = "W:\GIS_Mailer\new_mailer_toolbar"

# Get access to the current mxd 
mxd=arcpy.mapping.MapDocument("CURRENT") 
 
# Grab the dataframe object you want
df = arcpy.mapping.ListDataFrames(mxd,"*")[0]

# Go to Data View
mxd.activeView = df

# check to see if new case folder and files has already been created (for the first time); if not, then create them
for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
    if lyr.name == "Subject Parcels" or lyr.name == "Subject Parcel" or lyr.name == "Site":
        isCopy = lyr.dataSource         # is this a copy of the original Site shapefile for this site?
       
        if "Case" not in isCopy and "case" not in isCopy:                                  # check to see if the directory containing the layers contains any variation of "case"; if not, then save it in memory
	    Site_shp = "W:/GIS_Mailer/new_mailer_toolbar/Site.shp"                         # Site parcel shapefile
            Site_add = "W:/GIS_Mailer/new_mailer_toolbar/Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = "W:/GIS_Mailer/new_mailer_toolbar/Adjacents.shp"               # Adjacent parcel shapefile
	    
	else:
	    dirNew = lyr.workspacePath                               # if "case" is in the directory path, then save it to that directory
            Site_shp = dirNew + "\\Site.shp"                         # Site parcel shapefile
            Site_add = dirNew + "\\Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = dirNew + "\\Adjacents.shp"               # Adjacent parcel shapefile
	
	
# Get user input for parameters:  GPIN (parcel ID), Add and Remove site parcel check boxes (boolean, T/F)
parcel_id = arcpy.GetParameterAsText(0)
removeCheck = arcpy.GetParameterAsText(1)
addCheck = arcpy.GetParameterAsText(2)



# NEW SITE
# If Add and Remove boxes are not checked, proceed with site parcel selection process; otherwise, Add or Remove selected parcels
if not (str(removeCheck) == 'true') and not (str(addCheck) == 'true'):

    # Clear selection (using ArcObjects) and refresh map
    pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMap.ClearSelection()
    pMxDoc.ActiveView.Refresh()   
    
    # GPIN input processing (Tool Validator)
    # If GPIN contains dashes, remove them so that it can be classified as int (see Tool Validator code)
    # Then, if GPIN is not a number or is not 10 digits long, throw an error (see Tool Validator code)
    
    # Add in dashes so that GPIN can be found
    parcel_id = list(parcel_id)[0]+list(parcel_id)[1]+list(parcel_id)[2]+"-"+list(parcel_id)[3]+list(parcel_id)[4]+list(parcel_id)[5]+"-"+list(parcel_id)[6]+list(parcel_id)[7]+list(parcel_id)[8]+list(parcel_id)[9]
    
    # Search for GPIN in Tax Parcels layer; if not found, throw an error (Tool Validator)
    
    # Create expression for GPIN (that user has input)
    expression = "GPIN = " + "'" + parcel_id + "'"

    # Delete previous layers and buffers
    arcpy.Delete_management(Site_shp)
    arcpy.Delete_management(zoom_to_buffer)

    # Remove previously selected Site parcel shapefile from data frame
    for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
        if lyr.name == "Site":
            arcpy.mapping.RemoveLayer(df, lyr)
	if lyr.name == "Adjacents":
	    arcpy.mapping.RemoveLayer(df, lyr)
    arcpy.RefreshActiveView()
	
    # Select a tax parcel in Tax Parcel shapefile using GPIN expression (user input)
    arcpy.SelectLayerByAttribute_management(tax_parcels, "NEW_SELECTION", expression)

    # Write the selected features to a new feature class (shapefile) and clear selection from Tax Parcels layer
    arcpy.CopyFeatures_management(tax_parcels, Site_shp, "", "0", "0", "0")   # in, out, optional
    for lyr in arcpy.mapping.ListLayers(mxd,"",df):   
	    arcpy.SelectLayerByAttribute_management(lyr,"CLEAR_SELECTION")
    arcpy.RefreshActiveView()

    # Add Site shapefile to current map document and select it
    siteParcel = arcpy.mapping.Layer(Site_shp)
    arcpy.mapping.AddLayer(df, siteParcel, "TOP")

    for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
        if lyr.name == "Site":
            arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", expression)
    arcpy.RefreshActiveView()

    # Set parcel color to Rose -- see ArcObjects function above
    setSiteSymbology()
    
    # Set zoom level for selection of Adjacent parcels (button 2) - 700 foot buffer
    arcpy.Buffer_analysis("Site",zoom_to_buffer,"700 Feet","FULL","ROUND","NONE","#")
    bufferZoom = arcpy.mapping.Layer(zoom_to_buffer)
    ext = bufferZoom.getExtent()
    df.extent = ext

    # Clear selection (using ArcObjects) and refresh map
    pMap.ClearSelection()
    pMxDoc.UpdateContents()
    pMxDoc.ActiveView.Refresh()    
    
    # Save MXD
    mxd.save()
    
    
# REMOVE
# If Remove box is checked, remove selected parcels    
elif str(removeCheck) == 'true':
    # Execute Describe.fidset to get selected parcels, then execute Delete Features to remove the selected parcels.
    desc = arcpy.Describe("Site")
    if int(str(len(desc.fidSet.split(";")))) > 0:
	arcpy.DeleteFeatures_management("Site")

        # Clear selection and refresh map
        pApp = GetApp()
        pDoc = pApp.Document
        pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
        pMap = pMxDoc.FocusMap
        pMap.ClearSelection()
        pMxDoc.ActiveView.Refresh()
    # Save MXD	
    mxd.save()	

    
# ADD
# If Add box is checked, add selected parcels
elif str(addCheck) == 'true':
    
    # Delete previous layer used to hold previous "Add" parcels
    arcpy.Delete_management(Site_add)
    arcpy.Delete_management(zoom_to_buffer)    
    
    # Some code to indicate if a Site GPIN has already been selected and added to the layer; if it has, then remove the selection so it doesn't get added again
    
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
    
    # Run search cursor to get list of selected Site GPINs
    siteCursor = arcpy.da.SearchCursor("Site", ["GPIN"])
    siteCursorlist = list(siteCursor)
    siteCursorlen = len(siteCursorlist)
    siteGPINlist = list()
    
    for j in range(0,siteCursorlen):
      siteCursoritem = siteCursorlist[j]
      siteCursoritemString = str(siteCursoritem)
      siteCursoritemStringRep1 = siteCursoritemString.replace("(u'","")
      siteCursoritemStringRep2 = siteCursoritemStringRep1.replace("',)","")
      siteGPINlist.append(siteCursoritemStringRep2)  
    
    # Generate set of Tax Parcel GPINs that already exist in Site layer and deselect from Site and Tax Parcels - Mailer layers	
    GPINset = set(siteGPINlist).intersection(taxGPINlist)
    
    if len(GPINset) > 0: 
      for k in range(0,len(list(GPINset))):
        expression = "GPIN = " +"'" + list(GPINset)[k] + "'"
        arcpy.SelectLayerByAttribute_management("Site","REMOVE_FROM_SELECTION",expression)
        arcpy.SelectLayerByAttribute_management("Tax Parcels - Mailer","REMOVE_FROM_SELECTION",expression)
          
    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.
    # Add selected parcels from Tax Parcels layer to Site_add layer  
    arcpy.CopyFeatures_management(tax_parcels, Site_add, "", "0", "0", "0")
    site_add = arcpy.mapping.Layer(Site_add)
    arcpy.mapping.AddLayer(df, site_add, "TOP")
    
    # If Site_add layer contains parcels, then append them to the Site shapefile
    desc = arcpy.Describe(site_add)
    if int(str(len(desc.fidSet.split(";")))) > 0:
        arcpy.Append_management(site_add,"Site")
	
	arcpy.Delete_management("Site_add")
	
        # Set zoom level - 700 foot buffer around Site
        arcpy.Buffer_analysis("Site",zoom_to_buffer,"700 Feet","FULL","ROUND","NONE","#")
        bufferZoom = arcpy.mapping.Layer(zoom_to_buffer)
        ext = bufferZoom.getExtent()
        df.extent = ext
	
	# Clear selection
        pApp = GetApp()
        pDoc = pApp.Document
        pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
        pMap = pMxDoc.FocusMap
        pMap.ClearSelection()
        pMxDoc.ActiveView.Refresh()
	
    # Save MXD	
    mxd.save()