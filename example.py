import OpenMST.MSTProcess as MST

print("Opening file")
MOCfile = MST.openMOCFile('Your filename here.moc')

print("Saving experiments to XLSX files")
MOCfile.SaveXLSX() 

print("Finished exporting!")
MOCfile.close()
