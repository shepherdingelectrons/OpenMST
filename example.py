import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')
print("Processing file")
MOCfile.Process()

print("Saving experiments to XLSX files")
MOCfile.SaveXLSX() 

print("Finished exporting!")
MOCfile.close()
