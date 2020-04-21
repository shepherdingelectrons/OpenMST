# OpenMST
NanoTemper MST data sits in .MOC and requires specialist software to access - until now! 
This python script processes NanoTemper MST .MOC files and generates Excel files (*.XLSX) with the extracted data.

## Installation

Simply place **MSTProcess.py** in the folder with your scripts and import it.  See the examples for Usage

## Dependencies 

**sqlite3** - OpenMST was tested in Python 3.2.2, and uses sqlite3, which comes with this version of Python, but might need to be installed if you're using a diferent version of Python 

**xlsxwriter** - Used for exporting extracted data to Excel, in table format and generating the capillary scan and MST trace charts. 

## Usage

```python
import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
MOCfile.Process()
MOCfile.SaveXLSX() 
MOCfile.close()
```

Programatic control and access to the extracted MST data is possible (see examples for details)

