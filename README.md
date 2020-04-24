# OpenMST
[Microscale Thermophoresis](https://en.wikipedia.org/wiki/Microscale_thermophoresis) (MST) is a biophysical technique that meaures how fluorescently labelled proteins, DNA and small molecules move in response to an induced thermal gradient, and is used increasingly to characterise molecular interactions.  

NanoTemper MST instruments produce data in a custom .MOC format that requires specialist (and expensive) software to open.  OpenMST is a python program that processes NanoTemper MST .MOC files and generates Excel files (.XLSX) with the extracted data.

![OpenMST graphic](/images/OpenMST.jpg)
## Installation

Simply download **[MSTProcess.py](https://raw.githubusercontent.com/shepherdingelectrons/OpenMST/master/MSTProcess.py)** and place in the folder with your MOC files and import it with ```import MSTProcess```.  See the examples for simple usage, and also [More advanced usage](#more-advanced-usage).  

**NOTICE: The OpenMST script doesn't modify the MOC files, but work with copies of the MOC file to be safe**

More advanced users may want to setup the script as a site-package, so that it can be imported from any script folder.  To do this:
1. Locate the folder containing the version of python on your machine (i.e. ..\AppData\Local\Programs\Python\Python38)
2. In Lib\site-packages, make a folder 'OpenMST'
3. Place the MSTProcess.py script within the OpenMST folder, and make a ```__init__.py``` file (empty) and save it in the same folder. 

Open IDLE and try ```import OpenMST.MSTProcess as MST``` .  No error will indicate success!

## Dependencies 

**sqlite3** - OpenMST was tested in Python 3.2.2, and uses sqlite3 which comes with this version of Python.  It might need to be installed if you're using a different version of Python 

**xlsxwriter** - Used for exporting extracted data to Excel, in table format and generating the capillary scan and MST trace charts. Only used if saving to XLSX. 

## Help! I don't know anything about Python: Super Simple Installation & Usage Instructions
Follow these instructions if you want to keep things as simple as possible

**(1) Install Python 3**

Install Python 3 from the official website: https://www.python.org/downloads/  At the time of writing the latest release is Python 3.8.2. Download  and run the installer, making sure to check the "Add Python 3.8 to PATH" box.

![Add python to path](/images/Add_python38_to_path.jpg)

**(2) Install python module "xlsxwriter"**

Using the latest version of Python 3, and making that python is added to the PATH makes this much easier than it would otherwise be.

For reference, instructions can be found here: https://xlsxwriter.readthedocs.io/getting_started.html

The easiest method to install xlsxwriter with PIP is detailed below:
In Windows, open a command prompt by searching for "cmd" and running it as you would search for any program in Windows.
In the black box that appears, type:

```
pip install XlsxWriter
```
A warning message about the pip version might appear, but it can be safely ignored for our purposes.  You can close the command window now.
![pip install XlsxWriter](/images/pip_install_XlsxWriter.jpg)

**(3) Getting the OpenMST python library**

Download the **[MSTProcess.py](https://raw.githubusercontent.com/shepherdingelectrons/OpenMST/master/MSTProcess.py)** script (right click and click "Save link as..." and save the file to a new folder you want to work in.
Download the **[example.py](https://raw.githubusercontent.com/shepherdingelectrons/OpenMST/master/example.py)** script in the same way, to the same folder.
Copy the .MOC files you are interested in analysing into the same folder.

**(4)  Open IDLE**

IDLE is the Python editor.  Right click on example.py in the folder where you have placed it and click "Edit with IDLE", making sure it is the version of Python you just installed.
![edit example script with IDLE](/images/Edit_with_IDLE.jpg)

**(5) Edit and run the script**

You should see the following Python code in the IDLE window:
![Example script in IDLE](/images/example_py.jpg)

Simply change 'Your_filename_here.moc' to the MOC filename in the same folder that you want to process (keeping the filename within the quotation marks).  Press F5 to run the script, or from the menu at the top of the editor, Run--> Run Module.  All things being well, you should see:
```
Processing file
Saving experiments to XLSX files
Finished exporting!
```
If you see this text - congratulations - you have successfully set everything up and you have just processed your file! Look in the same folder as the MOC files, and you'll see the extracted MST data in the excel files, one per experiment.  Good luck!

## More advanced usage
We can open a .MOC file and perform processing on all the contained experiments, including populating each experiment with all the relevant capillary data with ```Process()```.

```python
import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
MOCfile.SaveXLSX() 
MOCfile.close()
```

The above code is sufficient for extracting all data from .MOC for viewing in Excel.  The experiment object of 'MOCfile' holds various bits of experiment-level data including an array of capillary objects (```.capillary```), experiment annotation ```(.experiment_annotation```) and other associated data (```.info```).  We can examine this, i.e. for experiment #2:
```python
import MSTProcess as MST
MOCfile = MST.openMOCFile('Your_filename_here.moc')

exp2 = MOCfile.experiment[2]
print("Annotation and data associated with Experiment #2:")
print(exp2.info)
print(exp2.experiment_annotation)
```
The capillary data can also be examined.  Capillary objects contain a couple of dictionaries with relevant information, .capinfo and .MSTinfo.  The centre position of the capillary (in millimetres) is stored in .CenterPosition. We can inspect this data, i.e. capillary #5 in experiment #2:
```python
import MSTProcess as MST
MOCfile = MST.openMOCFile('Your_filename_here.moc')

print("Capillary #5 in Experiment #2")
cap5 = MOCfile.experiment[2].capillary[5]
print(cap5.capinfo)
print(cap5.MSTinfo)
print(cap5.CenterPosition)
```
The raw data of the pre/post capillary scans and the MST traces themselves are extracted from the .MOC file as a byte array.  The pre-MST capillary scan, the post-MST capillary scan and the MST trace are saved in the capillary object as blobs, under the names; .preMST, .postMST and .MST_trace respectively.  The functions to convert the blobs to understandable data are ```MST.ExtractMSTTrace``` and ```MST.ExtractCapTrace```.  Two arrays are returned from each functions with the extracted data:
```python
import MSTProcess as MST
MOCfile = MST.openMOCFile('Your_filename_here.moc')

cap5 = MOCfile.experiment[0].capillary[5] # For readability, get a reference to the capillary #5 of experiment 0

time_s, fluorescence = MST.ExtractMSTTrace(cap5.MST_trace) # Unnormalised MST trace

distance_pre, preMST_scan = MST.ExtractCapTrace(cap5.preMST) # Raw capillary scan, pre-MST
distance_post, postMST_scan = MST.ExtractCapTrace(cap5.postMST) # Raw capillary scan, post-MST
```
The time is in seconds in the MST_trace, and the distance for the capillary scans is in millimetres. The MST data is normalised in the Excel output files, but is UNnormalised by default.  By specifying the parameter ```norm=True``` in the ExtractMSTTrace and ExtractCapTrace functions, normalised data is returned.  In the case of capillary scans, it is sometimes useful to make the distance measurement relative to the centre of capillary.  This is achieved by specifying ```xoffset```, like so:
```python
import MSTProcess as MST
MOCfile = MST.openMOCFile('Your_filename_here.moc')

cap5 = MOCfile.experiment[0].capillary[5] # For readability, get a reference to capillary #5 of experiment 0
centerpos = cap5.CenterPosition # grab stored CenterPosition

# Normalise the MST data (Fnorm)
time_s, fluorescence = MST.ExtractMSTTrace(cap5.MST_trace,norm=True)

# Normalise the capillary scan data AND normalise it (to 1.0f)
distance_pre, preMST_scan = MST.ExtractCapTrace(cap5.preMST, xoffset=centerpos, norm=True)
distance_post, postMST_scan = MST.ExtractCapTrace(cap5.postMST, xoffset=centerpos, norm=True)
```
We can also choose which individual experiments to save, and what to call them using ```MST.WriteExperimentToXLSX```. To save experiment #0 as 'myoutput.xlsx':
```python
import MSTProcess as MST
MOCfile = MST.openMOCFile('Your_filename_here.moc')

MST.WriteExperimentToXLSX(MOCfile.experiment[0],"myoutput.xlsx")
```
Remember we can also save all the experiments in the MOC file to separate XLSX files using ```SaveXLSX()```.
Finally, don't forget to close the file:
```python
MOCfile.close()
```
