import sqlite3
from struct import *
import os.path

# See examples for usage
    
def ExtractMSTTrace(blob, norm=False):        
    i=0
    x=[]
    y=[]

    if blob==None:
        #print("Error: this is a non-MST trace containing entry")
        return (["No data"],["No data"])
    elif len(blob)==0:
        return (["No data"],["No data"])
        
    while i<len(blob):
        a = blob[i:i+8]
        c = unpack('ff',a)

        x.append(c[0])
        y.append(c[1])

        i = i + 8

    if norm==True:# Normalise data before returning it       
        total = [y[i] for i,x in enumerate(x) if x<0]

        if len(total)==0: # some data is already normalised to 1000
            average = 1000.0
        else:
            average = sum(total)*1.0/len(total)

        y = [1000.0*j/average for j in y]
    
    return (x,y)

def ExtractCapTrace(blob, xoffset=0, norm=False):    
        i=0
        x=[]
        y=[]

        if blob==None:
            #print("Error: this is not a valid Cap trace containing entry")
            return (["No data"],["No data"])
        elif len(blob)==0:
            return (["No data"],["No data"])

        if xoffset == None: #CenterPosition was invalid for some reason - possibly an earlier file format issue
            xoffset = 0
            
        while i<len(blob):
            a = blob[i:i+16]
            c = unpack('ffff',a)

            x.append(c[0]-xoffset)
            y.append(c[3])

            i = i + 16
        
        if norm==True: # Normalise to 1.0
            height = max(y)
            y = [1.0*j/height for j in y]
        return (x,y)

def checkisSingle(array, arrayname, ident, identname):
    if len(array)==0:
        print("ERROR: ",arrayname,"NOT FOUND for ",identname,"=",ident)
        return True
    elif len(array)>1:
        print("ERROR: Multiple ",arrayname,"found for ",identname,"=",ident)
        return True
    return 0

##def getScanCapIDs(cursor, expID):
##    # returns a list of scancap IDs from the mcapScan (capillary scanning) table for each experiment (expID)
##    cursor.execute("SELECT ID FROM mCapScan WHERE ParentAction='"+expID+"'") #3596fa4f-ee82-423d-8c30-4de23aaef8dd
##    mCapIDs = [x[0] for x in cursor.fetchall()]
##    return mCapIDs

##def getMSTIDs(cursor, expID):
##    cursor.execute("SELECT ID FROM mMST WHERE ParentAction='"+expID+"'")
##    mMSTIDs = [x[0] for x in cursor.fetchall()]
##    return mMSTIDs

def getAnnotationbyID(cursor, annon_ID):
    # AnnotationRole, AnnotationType, Caption, NumericValue, TextValue
    fields = ["AnnotationRole", "AnnotationType", "Caption", "NumericValue", "TextValue"]
    fieldstr = ", ".join(fields)
    cursor.execute("SELECT "+fieldstr+" FROM Annotation WHERE ID='"+annon_ID+"'")
    annotation = {fields[i]:x for i,x in enumerate(cursor.fetchall()[0])}
    return annotation

class MSTCapillary:
    def __init__(self):
        self.capID = None
        self.scancapids = []
        self.capinfo = {} # Capillary meta data
        self.MSTinfo = {} # MST experiment meta data
        
        self.MST_trace = None # MST trace data here
        self.preMST = None # scan cap data here
        self.postMST = None # scan cap data here

        self.CenterPosition = 0

        self.capillary_annotation = []
        
class MSTExperiment:
    def __init__(self):
        self.aSeriesID = None
        self.tContID = None
        self.caption = ""
        self.expert = False # Expert Mode experiment type
        self.bindingaffinity = False # Binding Affinity experiment type
        self.info = {} # extracted misc data at experiment level
        self.cursor = None
        self.capillary = []
        self.experiment_annotation = []

    def getCapillaryData(self):
        if self.aSeriesID==None:
            print("ERROR: aSeriesID not valid. run getAllExperiments() on MSTFile object to extract experiments")

        # Use tContID
        CapIDs = self._getCapillaryIDs(self.tContID)
        for i,capid in enumerate(CapIDs):
            self.capillary.append(MSTCapillary())
            
            capinfo = self.getCapillaryInfo(capid)
            #print("capid=",capid," capinfo=",capinfo)

            self.capillary[i].capID = capid
                    
            # CapScan handing
            capScanIDs = self.getCapScanIDsFromCapID(capid) # Get a list of capscan IDs for this capillary
            self.capillary[i].scancapids = capScanIDs
               
            # Note there are usually 2 ScanCap entries, one for pre-MST and one for post-MST
            for cs in capScanIDs:
                scancapinfo = self.getScanCapInfo(cs)
                
                if scancapinfo["UpdateCapillaryCenterPosition"]==1:
                    # Pre-MST capscan
                    self.capillary[i].preMST = scancapinfo["CapScanTrace"]
                    capinfo["ExcitationPower"] = scancapinfo["ExcitationPower"]
                    capinfo["CenterPosition"] = scancapinfo["CenterPosition"]
                    self.capillary[i].CenterPosition = capinfo["CenterPosition"]
                else:
                    # Post_MST capscan
                    self.capillary[i].postMST = scancapinfo["CapScanTrace"]

            if self.expert == True:
                buffer = self.getExpertModeBuffer(capid)
                capinfo["BufferName"]=buffer
                
            self.capillary[i].capinfo = capinfo
            
            # MST trace handling
            mMSTID = self.getMSTIDsfromCapID(capid) # should be only a single MST entry per capillary experiment
            MSTinfo = self.getMSTInfo(mMSTID)
            if len(MSTinfo)>0: # we have a valid return from getMSTInfo
                self.capillary[i].MST_trace = MSTinfo["MstTrace"]
                MSTinfo.pop("MstTrace")
            self.capillary[i].MSTinfo = MSTinfo

            annons = self.getAnnotations(self.capillary[i].capinfo)
            self.capillary[i].capillary_annotation = annons
            self.capillary[i].capinfo.pop("Annotations")
            
        return

    def getExpertModeBuffer(self, capID):
        self.cursor.execute("SELECT BufferName FROM ExpertModeCapillarySettings WHERE Capillary='"+capID+"'")
        buffer = self.cursor.fetchall()[0]
        return buffer[0]

    def getMSTInfo(self, MSTcapID):
    # Use an experiment from mMST to fetch interesting data: ExcitationPower, MstPower, MstTrace, NominalDurationOfPhase1, NominalDurationOfPhase2, NominalDurationOfPhase3, NominalDurationOfPhase1,State,
    # If State=0, data is valid, else State=1 seems to mean some kind of non-data table entry

        fields = ["ExcitationPower", "MstPower", "NominalDurationOfPhase1", "NominalDurationOfPhase2", "NominalDurationOfPhase3", "State","MstTrace"]
        fieldstr = ", ".join(fields)
        self.cursor.execute("SELECT "+fieldstr+" FROM mMST WHERE ID='"+MSTcapID+"'")
       
        MSTinfo = self.cursor.fetchall()

        if len(MSTinfo)==0: # If we didn't find the MSTinfo from MSTcapID
            MSTinfo = {}
        else:
            MSTinfo = {fields[i]:x for i,x in enumerate(MSTinfo[0])}
            
        return MSTinfo

    def getMSTIDsfromCapID(self, capID):
        self.cursor.execute("SELECT ID FROM mMST WHERE Container='"+capID+"'")
        capids = [x[0] for x in self.cursor.fetchall()]

        if checkisSingle(capids, "capids", capID, "capID"): capids = [""] # Return nothing if failed
        return capids[0]

    def getScanCapInfo(self, scancapID):
    # ID from mCap to extract CenterPosition, ExcitationPower, UpdateCapillaryCenterPosition, CapScanTrace
    # If UpdateCapillaryCenterPosition is '1', then is the pre-MST capScan, if '0', then is the post-MST capScan
    # Container is our lookup index into the tCapillary table
        fields = ["CenterPosition", "ExcitationPower", "UpdateCapillaryCenterPosition", "CapScanTrace"]
        fieldstr = ", ".join(fields)
        self.cursor.execute("SELECT "+fieldstr+" FROM mCapScan WHERE ID='"+scancapID+"'")
        capInfo = {fields[i]:x for i,x in enumerate(self.cursor.fetchall()[0])}
        return capInfo
    
    def getCapillaryInfo(self, capid):
        # Annotations, Caption, IndexOnParentContainer
        fields = ["Annotations", "Caption", "IndexOnParentContainer"]
        fieldstr = ", ".join(fields)
        self.cursor.execute("SELECT "+fieldstr+" FROM tCapillary WHERE ID='"+capid+"'")
        capinfo = {fields[i]:x for i,x in enumerate(self.cursor.fetchall()[0])}
        return capinfo

    def getCapScanIDsFromCapID(self, capID):
        self.cursor.execute("SELECT ID FROM mCapScan WHERE Container='"+capID+"'")
        capids = [x[0] for x in self.cursor.fetchall()]
        return capids

    def _getCapillaryIDs(self, cont_id):
        self.cursor.execute("SELECT ID from tCapillary WHERE ParentContainer='"+cont_id+"'")
        capids = [x[0] for x in self.cursor.fetchall()]
        return capids
    
    def getAnnotations(self, data):
        annon_dicts = []
        if "Annotations" in data.keys():
            if data["Annotations"]!="":
                annon = data["Annotations"].split(";")
                for a in annon:
                    text = getAnnotationbyID(self.cursor, a) # returns a dictionary of associated text
                    annon_dicts.append(text) # build an array of dictionaries as there are multiple annotation lines
        return annon_dicts

def openMOCFile(filename):
    if not os.path.isfile(filename):
        print("ERROR:",filename,"does not exist!") 
        return None
    else:
        return MSTFile(filename)
        
class MSTFile:
    def __init__(self, filename):
        self.conn = sqlite3.connect(filename)
        self.cursor = self.conn.cursor()
        self.DeviceType = None
        self.SerialNumber = None
        self.MacAddress = None
        self.filebase = filename.split('.')[0]

        self.experiment = []
        self.validtables = []
        
        self.getTableList()
        self.Process() # Experiment information is now extracted automatically

    def getTableList(self):
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        self.validtables = [x[0] for x in self.cursor.fetchall()]
        
    def getMachineInfo(self):
        info = []
        if "tDevice" in self.validtables:
            self.cursor.execute("SELECT DeviceType, SerialNumber, MacAddress FROM tDevice")
            info = [x for x in self.cursor.fetchall()[0]]
            self.DeviceType, self.SerialNumber, self.MacAddress = info
        return info

    def getAllExperiments(self, verbose=True):
        allExperiments = self._gettContainerIDs()
        ExpertModeaIDs = self.getExpertModeaIDs()
        BindingAffinityaIDs = self.getBindingAffinityaIDs()

        self.experiment = [] # Clear experiments
        
        for i, tcontID in enumerate(allExperiments):
            self.experiment.append(MSTExperiment())

            aSeriesID = self._getaSeriesID(tcontID)
            caption = self._getaSeriesCaption(aSeriesID)
            
            self.experiment[i].aSeriesID = aSeriesID
            self.experiment[i].tContID = tcontID
            self.experiment[i].caption = caption
            self.experiment[i].cursor = self.cursor # pass on DB reference

            info = {} 
            # Extract different bits of information depending on experiment type
            if aSeriesID in ExpertModeaIDs:
                self.experiment[i].expert = True
                info = self.getExpertModeInfo(aSeriesID)
            elif aSeriesID in BindingAffinityaIDs:
                self.experiment[i].BindingAffinity = True
                info = self.getBindingAffinityInfo(aSeriesID)

            self.experiment[i].info = info

            exp_annon = self.experiment[i].getAnnotations(self.experiment[i].info)
            self.experiment[i].experiment_annotation = exp_annon
            if "Annotations" in self.experiment[i].info.keys(): self.experiment[i].info.pop("Annotations")

            if verbose: print("Experiment #",i,":OK")

    def Process(self,verbose=True):
        if verbose: print("Extracting experimental data from file:",self.filebase+".MOC")
        self.getAllExperiments(verbose)
        if verbose: print("Number of experiments extracted:",len(self.experiment))
        if verbose: print("Extracting capillary data")
        if verbose: self.getAllCapillaryData(verbose)
        print("Done!")

    def SaveXLSX(self):
        for i, exp in enumerate(self.experiment):
            filename = self.filebase+"_exp"+str(i)+".xlsx"
            WriteExperimentToXLSX(exp,filename)       

    def getAllCapillaryData(self, verbose=True):
        for i in range(len(self.experiment)):
            self.experiment[i].getCapillaryData()
            if verbose: print("Capillary data for experiment #",i,":OK")

    def getExpertModeaIDs(self):
        expertModeIDs = []
        if "ExpertModeEntity" in self.validtables:
            self.cursor.execute("SELECT TopmostAction FROM ExpertModeEntity")
            expertModeIDs = [x[0] for x in self.cursor.fetchall()]
            
        return expertModeIDs

    def getBindingAffinityaIDs(self):
        bindingAffinityIDs = []
        if "BindingAffinityEntity" in self.validtables:
            self.cursor.execute("SELECT TopmostAction FROM BindingAffinityEntity")
            bindingAffinityIDs = [x[0] for x in self.cursor.fetchall()]
        return bindingAffinityIDs

    def getExpertModeInfo(self, aSeriesID):
        # Get expert mode useful information by ID: Name, MSTPower, IsDeleted, FirstAvailablePosition, LastCapillaryPosition, ExcitationPower, Excitation, DurationOfPhase1, DurationOfPhase2, DurationOfPhase3
        # I think cR for Excitation = is colour Red
        fields = ["Name", "MSTPower", "IsDeleted", "FirstAvailablePosition", "LastCapillaryPosition", "ExcitationPower", "Excitation", "DurationOfPhase1", "DurationOfPhase2", "DurationOfPhase3", "Annotations"]
        fieldstr = ", ".join(fields)
        self.cursor.execute("SELECT "+fieldstr+" FROM ExpertModeEntity WHERE TopmostAction='"+aSeriesID+"'")
        info = {fields[i]:x for i,x in enumerate(self.cursor.fetchall()[0])}
        return info

    def getBindingAffinityInfo(self, aSeriesID):
        # Get binding affinity mode information by ID: Name, LigandName, TargetName, IsDeleted, AssayBufferName, MSTPower, CapillaryType, Excitation, ExcitationPower, FirstAvailablePosition, LigandConcentrationInThisAssay,  TargetConcentrationInThisAssay
        # I think cR for Excitation = is colour Red
        fields = ["Name", "LigandName", "TargetName", "IsDeleted", "AssayBufferName", "MSTPower", "CapillaryType", "Excitation", "ExcitationPower", "FirstAvailablePosition", "LigandConcentrationInThisAssay", "TargetConcentrationInThisAssay","Annotations"]
        fieldstr = ", ".join(fields)
        self.cursor.execute("SELECT "+fieldstr+" FROM BindingAffinityEntity WHERE TopmostAction='"+aSeriesID+"'")
        info = {fields[i]:x for i,x in enumerate(self.cursor.fetchall()[0])}
        return info

    def _getaSeriesCaption(self, aSeriesID):
        self.cursor.execute("SELECT Caption FROM aSeries WHERE ID='"+aSeriesID+"'")
        caption = self.cursor.fetchall()#[0]

        if len(caption)==0:
            caption = ""
        else:
            caption = caption[0]
            checkisSingle(caption, 'caption', aSeriesID,"aSeriesID")
            caption = caption[0]
            
        return caption
            
    def _gettContainerIDs(self):
        self.cursor.execute("SELECT ID FROM tContainer")
        tContainerIDs = [x[0] for x in self.cursor.fetchall()]
        return tContainerIDs
    
    def _getaSeriesID(self, tcontID):
        # This appears to build an bridge between experiment ID (i.e aSeries:ID)
        # and tContainer ID.
        self.cursor.execute("SELECT ParentAction FROM aHandling WHERE Container='"+tcontID+"'")

        aSeriesID = self.cursor.fetchall()

        if len(aSeriesID)==0:
            aSeriesID=""
        else:
            aSeriesID = [x for x in aSeriesID[0]]
            checkisSingle(aSeriesID, "aSeriesID", tcontID, "tcontID")
            aSeriesID = aSeriesID[0]
        return aSeriesID

    def close(self):
        self.conn.close()


#############################################################################
def WriteExperimentToXLSX(experiment, filename):
    import xlsxwriter
    from xlsxwriter.utility import xl_rowcol_to_cell

    workbook = xlsxwriter.Workbook(filename)
    bold = workbook.add_format({'bold': True})
    centre = workbook.add_format({'align': 'center'})
    boldANDcentre = workbook.add_format({'align': 'center','bold': True})
    
    # Information sheet:
    infosheet = workbook.add_worksheet("Experiment information")
    infosheet.write(0,0,"Experiment information",bold) # row, notation (Y,X)
    infosheet.write(1,0,"Experiment type:")
    if experiment.expert==True:
        infosheet.write(1,1,"Expert mode")
    elif experiment.bindingaffinity==True:
        infosheet.write(1,1,"Binding affinity")
    
    infosheet.write(2,0,"Experiment misc info type:")
    infosheet.write(2,1,experiment.caption)

    writedictionary(infosheet,experiment.info,3,0,headerformat=bold)

    infosheet.write(5,0,"Experiment annotations:")
    row = 6
    writeheader=True

    for i, an in enumerate(experiment.experiment_annotation): # Needs to be an array of dictionaries
        writedictionary(infosheet,an,row,0,headerformat=bold,write_header=writeheader)
        row+=1
        if i==0: writeheader=False
    row+=2

    infosheet.write(row,0,"Capillary information:",bold)

    row+=1
    writeheader=True
    MSTrow = row+2+len(experiment.capillary)

    notdeleted = 1
    if "IsDeleted" in experiment.info.keys():
        notdeleted = 1-experiment.info["IsDeleted"]

    if len(experiment.capillary)==0:
        print("Empty experiment: No capillary data recovered")
        return
    
    for i,cap in enumerate(experiment.capillary):
        width = len(cap.capinfo)
        writedictionary(infosheet,cap.capinfo,row,0,headerformat=bold,write_header=writeheader)
        writedictionary(infosheet,cap.capinfo,row+2+len(experiment.capillary),0,headerformat=bold,write_header=writeheader)

        # All capillary annotations
        for j,an in enumerate(cap.capillary_annotation):
            writedictionary(infosheet,an,row,((1+len(an))*j)+width+1,headerformat=bold,write_header=writeheader)
        # MST data annotations
        writedictionary(infosheet,cap.MSTinfo,row+2+len(experiment.capillary),width+1,headerformat=bold,write_header=writeheader)
        
        row+=1
        if i==0: writeheader=False
    row+=1
    
    # MST traces
    MSTchartsheet = workbook.add_chartsheet("MST chart")
    # Cap scan chart - horizontal representation with raw data
    CapScanchartsheet = workbook.add_chartsheet("Capillary Scan")    
    # Cap scan overlays
    Overlaychartsheet = workbook.add_chartsheet("Capillary Scan overlay")
    
    # Capillary scan sheet
    scansheet = workbook.add_worksheet("CapillaryScans")
    capscanchart = workbook.add_chart({'type': "scatter",'subtype':"straight_with_markers"})
    overlayscanchart = workbook.add_chart({'type': "scatter",'subtype':"straight_with_markers"})
    
    column = 0
    for cap in experiment.capillary:
        name = cap.capinfo["Caption"]

        scansheet.merge_range(0,column,0,column+5,name, boldANDcentre)
        scansheet.merge_range(1,column,1,column+2,"Pre-MST",centre)
        scansheet.merge_range(1,column+3,1,column+5,"Post-MST",centre)
        scansheet.write(2,column,"dist",centre)
        scansheet.write(2,column+1,"raw",centre)
        scansheet.write(2,column+2,"reldist",centre)
        scansheet.write(2,column+3,"norm",centre)
     
        scansheet.write(2,column+4,"dist",centre)
        scansheet.write(2,column+5,"raw",centre)
        scansheet.write(2,column+6,"reldist",centre)
        scansheet.write(2,column+7,"norm",centre)

        cp = cap.CenterPosition
        x,y = ExtractCapTrace(cap.preMST)
        x_rel, y_norm  = ExtractCapTrace(cap.preMST, xoffset=cp, norm=True)
        
        if notdeleted:      
            for i in range(len(x)):
                scansheet.write(3+i,column,x[i])
                scansheet.write(3+i,column+1,y[i])
                scansheet.write(3+i,column+2,x_rel[i])
                scansheet.write(3+i,column+3,y_norm[i])
                
            x,y = ExtractCapTrace(cap.postMST)
            x_rel, y_norm  = ExtractCapTrace(cap.postMST, xoffset=cp, norm=True)
    
            for i in range(len(x)):
                scansheet.write(3+i,column+4,x[i])
                scansheet.write(3+i,column+5,y[i])
                scansheet.write(3+i,column+6,x_rel[i])
                scansheet.write(3+i,column+7,y_norm[i])
            
        startID = xl_rowcol_to_cell(3, column) # pre-MST
        endID = xl_rowcol_to_cell(2+len(x), column)
        distrange = '=CapillaryScans!'+startID+":"+endID 

        startID = xl_rowcol_to_cell(3, column+1) # pre-MST
        endID = xl_rowcol_to_cell(2+len(x), column+1)
        datarange = '=CapillaryScans!'+startID+":"+endID
        
        capscanchart.add_series({'name':name+"_preMST",'categories':distrange, 'values': datarange, 'marker': {'type': 'none'}, 'line':   {'color': 'black'}})

        startID = xl_rowcol_to_cell(3, column+4) # post-MST
        endID = xl_rowcol_to_cell(2+len(x), column+4)
        distrange = '=CapillaryScans!'+startID+":"+endID
        
        startID = xl_rowcol_to_cell(3, column+5)# post-MST
        endID = xl_rowcol_to_cell(2+len(x), column+5)
        datarange = '=CapillaryScans!'+startID+":"+endID

        capscanchart.add_series({'name':name+"_postMST",'categories':distrange, 'values': datarange, 'marker': {'type': 'none'}, 'line':   {'color': 'gray'}})

        startID = xl_rowcol_to_cell(3, column+2) # pre-MST
        endID = xl_rowcol_to_cell(2+len(x), column+2)
        distrange = '=CapillaryScans!'+startID+":"+endID 

        startID = xl_rowcol_to_cell(3, column+3) # pre-MST
        endID = xl_rowcol_to_cell(2+len(x), column+3)
        datarange = '=CapillaryScans!'+startID+":"+endID
        
        overlayscanchart.add_series({'name':name+"_preMST",'categories':distrange, 'values': datarange, 'marker': {'type': 'none'}, 'line':   {'color': 'black'}})
        
        startID = xl_rowcol_to_cell(3, column+6) # post-MST
        endID = xl_rowcol_to_cell(2+len(x), column+6)
        distrange = '=CapillaryScans!'+startID+":"+endID
        
        startID = xl_rowcol_to_cell(3, column+7)# post-MST
        endID = xl_rowcol_to_cell(2+len(x), column+7)
        datarange = '=CapillaryScans!'+startID+":"+endID

        overlayscanchart.add_series({'name':name+"_postMST",'categories':distrange, 'values': datarange, 'marker': {'type': 'none'}, 'line':   {'color': 'gray'}})
        
        column+=8
        
    
    # MST trace sheet

    chart = workbook.add_chart({'type': "scatter",'subtype':"straight_with_markers"})
    
    column = 0
    MSTsheet = workbook.add_worksheet("MST_trace_data")
    
    for cap in experiment.capillary:
        name = cap.capinfo["Caption"]

        MSTsheet.merge_range(0,column,0,column+2,name, boldANDcentre)
        MSTsheet.write(1,column,"time/s",centre)
        MSTsheet.write(1,column+1,"Raw",centre)
        MSTsheet.write(1,column+2,"Fnorm",centre)

        # Calculate Fnorm-100%
        x,y = ExtractMSTTrace(cap.MST_trace)
        x,y_norm = ExtractMSTTrace(cap.MST_trace, norm=True)

        if notdeleted:    
            for i in range(len(x)):
                MSTsheet.write(2+i,column,x[i])
                MSTsheet.write(2+i,column+1,y[i])
                MSTsheet.write(2+i,column+2,y_norm[i])

        startID = xl_rowcol_to_cell(2, column)
        endID = xl_rowcol_to_cell(1+len(x), column)
        timerange = '=MST_trace_data!'+startID+":"+endID 

        startID = xl_rowcol_to_cell(2, column+2)
        endID = xl_rowcol_to_cell(1+len(x), column+2)
        datarange = '=MST_trace_data!'+startID+":"+endID

        cap_index = cap.capinfo["IndexOnParentContainer"]
        startTrans = 100
        endTrans = 255
        trans = startTrans + (endTrans-startTrans)*(cap_index/15.0)

        linecolour = "#55"+hex(int(trans))[2:]+"00"
       
        chart.add_series({'name':name,'categories':timerange, 'values': datarange, 'marker': {'type': 'none'}, 'line':   {'color': linecolour}})
        column+=3   

    MSTchartsheet.set_chart(chart)
    CapScanchartsheet.set_chart(capscanchart)
    Overlaychartsheet.set_chart(overlayscanchart)
           
    workbook.close()

def writedictionary(sheet, dictionary, startrow, startcolumn, horizontal=True, headerformat=None,write_header=True):
    #  Note when write_header is false, we still write to the row below startrow - could cause unexpected behaviour
    for i,key in enumerate(dictionary):
        if horizontal:
            if write_header: sheet.write(startrow,startcolumn+i,key,headerformat)
            sheet.write(startrow+1,startcolumn+i,dictionary[key])
        else:
            if write_header: sheet.write(startrow+i,startcolumn,key,headerformat)
            sheet.write(startrow+i,startcolumn+1,dictionary[key])
            

