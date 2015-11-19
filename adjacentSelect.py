# Select Adjacent Parcels
#adjacentSelect.py - select and zoom to immediate adjacents, colorize them, and allow for selection of additional adjacent parcels


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
  #"""Creates a new comtypes POINTER object where\n\
  #MyClass is the class to be instantiated,\n\
  #MyInterface is the interface to be assigned"""
  from comtypes.client import CreateObject
  try:
      ptr = CreateObject(MyClass, interface=MyInterface)
      return ptr
  except:
      return None

# Used to cast objects to a different interface (ex., pDoc cast to IMxDocument = pMxDoc)
def CType(obj, interface):
  #"""Casts obj to interface and returns comtypes POINTER or None"""
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
  #"""In a standalone script, retrieves the first app session found.\n\
  #app must be 'ArcMap' (default) or 'ArcCatalog'\n\
  #Execute GetDesktopModules() first"""
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
    
# Key Numbering (create fields for labels) - Adjacents layer will be used for labels
def keyNumber():    
    # add new field with Latitude (y) coordinate for each GPIN
    arcpy.AddField_management(Adjacents_shp,"Latitude","DOUBLE","#","#","#","#","NULLABLE","NON_REQUIRED","#")
    arcpy.CalculateField_management(Adjacents_shp,"Latitude","!SHAPE.extent.YMax!","PYTHON_9.3","#")
    
    # sort by Latitude (highest to lowest) and add KeyLabel field (1, 2, 3, etc. in order from N to S) in "Adjacent_sort"
    arcpy.Sort_management(Adjacents_shp, Adjacents_sort, [["Latitude", "DESCENDING"]])

    # add and calculate KeyLabel field
    arcpy.AddField_management(Adjacents_sort,"KeyLabel","SHORT","#","#","#","#","NULLABLE","NON_REQUIRED","#")
    arcpy.CalculateField_management(Adjacents_sort,"KeyLabel","!FID!+1","PYTHON_9.3","#")
    
    # delete Adjacents shapefile and rename Adjacents_sort to Adjacents (before adding)
    arcpy.Delete_management(Adjacents_shp)
    arcpy.CopyFeatures_management(Adjacents_sort, Adjacents_shp)
    arcpy.Delete_management(Adjacents_sort)
    
    
# VARIABLES

# Global variables:
tax_parcels = "Tax Parcels - Mailer"

site_rose = "W:/GIS_Mailer/new_mailer_toolbar/Site_color_rose.lyr"             # contains symbology (color) for Site parcel layer
adjacent_green = "W:/GIS_Mailer/new_mailer_toolbar/Adjacent_color_green.lyr"   # contains symbology (color) for Adjacent parcel layer
zoom_to_buffer = "W:/GIS_Mailer/new_mailer_toolbar/Site_buffer_700ft.shp"      # buffer around Site parcel(s) used to 


#Adjacents_shp = "in_memory\\Adjacents"
#Adjacents_add = "in_memory\\Adjacent_add"
Adjacentsadd_sort = "W:/GIS_Mailer/new_mailer_toolbar/Adjacentadd_sort.shp"

# Initialize variables
Site_shp = ""                                                                  # Site parcel shapefile
Site_add = ""                                                                  # new layer shapefile with new Site parcels to be added 
Adjacents_shp = ""                                                             # Adjacent parcel shapefile
Adjacents_add = ""                                                             # new Adjacent parcel shapefile with new Adjacent parcels to be added
Adjacents_sort = ""

#----------------------------------------------------------------------


    
# Set the workspace
arcpy.env.workspace = "W:\GIS_Mailer\new_mailer_toolbar"

# Get access to the current mxd 
mxd=arcpy.mapping.MapDocument("CURRENT") 
 
# Grab the dataframe object you want
df = arcpy.mapping.ListDataFrames(mxd,"*")[0]

#Go to Data View
mxd.activeView = df

# check to see if new case folder and files has already been created (for the first time); if not, then create them
for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
    if lyr.name == "Subject Parcels" or lyr.name == "Subject Parcel" or lyr.name == "Site":
        isCopy = lyr.dataSource         # is this a copy of the original Site shapefile for this site?
       
        if "Case" not in isCopy and "case" not in isCopy:
	    Site_shp = "W:/GIS_Mailer/new_mailer_toolbar/Site.shp"                         # Site parcel shapefile
            Site_add = "W:/GIS_Mailer/new_mailer_toolbar/Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = "W:/GIS_Mailer/new_mailer_toolbar/Adjacents.shp"               # Adjacent parcel shapefile
	    Adjacents_add = "W:/GIS_Mailer/new_mailer_toolbar/Adjacent_add.shp"            # new Adjacent parcel shapefile with new Adjacent parcels to be added
	    Adjacents_sort = "W:/GIS_Mailer/new_mailer_toolbar/Adjacent_sort.shp"          # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
	    
	    
	else:
	    dirNew = lyr.workspacePath                               # if "case" is in the directory path, then save it to that directory
            Site_shp = dirNew + "\\Site.shp"                         # Site parcel shapefile
            Site_add = dirNew + "\\Site_add.shp"                     # new layer shapefile with new Site parcels to be added 
            Adjacents_shp = dirNew + "\\Adjacents.shp"               # Adjacent parcel shapefile
	    Adjacents_add = dirNew + "\\Adjacent_add.shp"            # new Adjacent parcel shapefile with new Adjacent parcels to be added
	    Adjacents_sort = dirNew + "\\Adjacent_sort.shp"          # new Adjacent parcel shapefile with new Adjacent parcels sorted by Latitude (N to S)
	    


# Get user input for parameters:  Add and Remove adjacent parcel check boxes (boolean, T/F)
removeCheck = arcpy.GetParameterAsText(0)
addCheck = arcpy.GetParameterAsText(1)



# NEW ADJACENTS
# If Add and Remove boxes are not checked,  proceed with automatic selection and addition of immediate adjacents; otherwise, Add or Remove selected parcels
if not (str(removeCheck) == 'true') and not (str(addCheck) == 'true'):

    # Clear any selections
    pApp = GetApp()
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, comtypes.gen.esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pMap.ClearSelection()
    pMxDoc.ActiveView.Refresh()  
    
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
    
    pMap.ClearSelection()
    pMxDoc.ActiveView.Refresh()  
    
    # Create new fields for labeling (key numbering) - see function above
    keyNumber()

    # Add Adjacents shapefile to current map document, make it the second layer from the top, and select it
    adjacentParcel = arcpy.mapping.Layer(Adjacents_shp)
    arcpy.mapping.AddLayer(df, adjacentParcel, "TOP")
    for lyr in arcpy.mapping.ListLayers(mxd,"",df):
      if lyr.name == "Adjacents":
        moveLayer = lyr
      if lyr.name == "Site":
        refLayer = lyr
    arcpy.mapping.MoveLayer(df,refLayer,moveLayer,"AFTER")
    
    # Set parcel color to Green -- see ArcObjects function above
    setAdjacentSymbology()
    
    # Set zoom level - 300 foot buffer around Adjacents
    arcpy.Buffer_analysis("Adjacents",zoom_to_buffer,"300 Feet","FULL","ROUND","NONE","#")
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
    desc = arcpy.Describe("Adjacents")
    if int(str(len(desc.fidSet.split(";")))) > 0:
        arcpy.DeleteFeatures_management("Adjacents")

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
    
    # Create new fields for labeling (key numbering) - see function above
    keyNumber()
   
    # Execute Describe and if some parcels have been selected, then execute CopyFeatures and Append_management to add the selected parcels.
    # Add selected parcels from Tax Parcels layer to Adjacents_add layer  
    adjacents_add = arcpy.mapping.Layer(Adjacents_add)
    arcpy.mapping.AddLayer(df, adjacents_add, "TOP")
    
    # If Adjacents_add layer contains parcels, then append them to the Adjacents shapefile
    desc = arcpy.Describe(adjacents_add)
    if  int(str(len(desc.fidSet.split(";")))) > 0:
        arcpy.Append_management(adjacents_add,"Adjacents")
        arcpy.ApplySymbologyFromLayer_management("Adjacents", adjacent_green)
	
	# Remove the Adjacents_add layer (temporary layer to hold parcels to add)
        for lyr in arcpy.mapping.ListLayers(mxd, "*",df):
            if lyr.name == "Adjacent_add":
                arcpy.mapping.RemoveLayer(df, lyr)
        arcpy.RefreshActiveView()
        arcpy.Delete_management(Adjacents_add)
	
	# Set zoom level - 300 foot buffer around Adjacents
        arcpy.Buffer_analysis("Adjacents",zoom_to_buffer,"300 Feet","FULL","ROUND","NONE","#")
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
	
	
