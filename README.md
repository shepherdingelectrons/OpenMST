# OpenMST
[Microscale Thermophoresis](https://en.wikipedia.org/wiki/Microscale_thermophoresis) (MST) is a biophysical technique that meaures how fluorescently labelled proteins, DNA and small molecules move in response to an induced thermal gradient, and is used increasingly to characterise molecular interactions.  

NanoTemper MST instruments produce data in a custom .MOC format that requires specialist (and expensive) software to open.  OpenMST is a python program that processes NanoTemper MST .MOC files and generates Excel files (.XLSX) with the extracted data.

## Installation

Simply place **MSTProcess.py** in the folder with your scripts and import it.  See the examples for Usage

## Dependencies 

**sqlite3** - OpenMST was tested in Python 3.2.2, and uses sqlite3 which comes with this version of Python.  It might need to be installed if you're using a different version of Python 

**xlsxwriter** - Used for exporting extracted data to Excel, in table format and generating the capillary scan and MST trace charts. Only used if saving to XLSX. 

## Usage

```python
import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
MOCfile.Process()
MOCfile.SaveXLSX() 
MOCfile.close()
```

The above code is sufficient for simply extracting data from .MOC for viewing in Excel. However more selective programatic control and access to the extracted MST data is possible (also see examples for details).

We can open a .MOC file using and perform processing on all the contained experiments, including populating each experiment with all the relevant capillary data. Note that 'None' is returned if the .MOC file can't be opened.
```python
import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
MOCfile.Process() # Extracts all the experiments in the MOC file, and populates with all the capillary information
```
If we wanted, we could choose to only populate the capillary data for certain experiments, i.e experiments #0 and #2:
```python
import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
MOCfile.getAllExperiments() #

MOCfile.experiment[0].getCapillaryData() # Get the capillary data for experiment #0
MOCfile.experiment[2].getCapillaryData() # Get the capillary data for experiment #2
```
The experiment array of 'MOCfile' holds various bits of experiment-level data including an array of capillary objects (.capillary), experiment annotation (.experiment_annotation) and other associated data (.info).  We can examine this in the expected fashion, i.e. for experiment #2:
```python
exp2 = MOCfile.experiment[2]
print("Annotation and data associated with Experiment #2:")
print(exp2.info)
print(exp2.experiment_annotation)
```
Once capillary data is populated (either by MOCfile.Process() or MOCfile.getAllCapillaryData()), the capillary data can be examined.  Capillary objects contain a couple of dictionaries with relevant information, .capinfo and .MSTinfo.  The centre position of the capillary (in millimetres) is stored in .CenterPosition. We can inspect this data like so, i.e. the 6th capillary in experiment #2:
```python
print("Capillary #5 in Experiment #2")
cap6 = MOCfile.experiment[2].capillary[5]
print(cap6.capinfo)
print(cap6.MSTinfo)
print(cap6.CenterPosition)
```
The raw data of the pre/post capillary scans and the MST traces themselves are extracted from the .MOC file as a byte array.  The pre-MST capillary scan, the post-MST capillary scan and the MST trace are saved in the capillary object as blobs, under the names; .preMST, .postMST and .MST_trace respectively.  The functions to convert the blobs to understandable data are MST.ExtractMSTTrace and MST.ExtractCapTrace.  So then, we can access this data directly:
```python
cap6 = MOCfile.experiment[0].capillary[5] # For readability, get a reference to the 6th capillary of experiment 0

time_s, fluorescence = MST.ExtractMSTTrace(cap6.MST_trace) # Unnormalised MST trace

distance_pre, preMST_scan = MST.ExtractCapTrace(cap6.preMST) # Raw capillary scan, pre-MST
distance_post, postMST_scan = MST.ExtractCapTrace(cap6.postMST) # Raw capillary scan, post-MST
```
The time is in seconds in the MST_trace, and the distance for the capillary scans is in millimetres.
The MST data is normalised in the Excel output files, but is UNnormalised by default.  By specifying the parameter 'norm=True' in the ExtractMSTTrace and ExtractCapTrace functions, normalised data is returned.  In the case of capillary scans, it is sometimes useful to make the distance measurement relative to the centre of capillary.  This is achieved by specifying 'xoffset', like so:
```python
cap6 = MOCfile.experiment[0].capillary[5] # For readability, get a reference to the 6th capillary of experiment 0
centerpos = cap6.CenterPosition # grab stored CenterPosition

# Normalise the MST data (Fnorm)
time_s, fluorescence = MST.ExtractMSTTrace(cap6.MST_trace,norm=True)

# Normalise the capillary scan data AND normalise it (to 1.0f)
distance_pre, preMST_scan = MST.ExtractCapTrace(cap6.preMST, xoffset=centerpos, norm=True)
distance_post, postMST_scan = MST.ExtractCapTrace(cap6.postMST, xoffset=centerpos, norm=True)
```
We can also choose which individual experiments to save, and what to call them. To save experiment #0 as 'myoutput.xlsx':
```python
MST.WriteExperimentToXLSX(MOCfile.experiment[0],"myoutput.xlsx")
```
Finally, don't forget to close the file after use:
```python
MOCfile.close()
```
