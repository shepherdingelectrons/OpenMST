import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')

# Get machine info
machinfo = MOCfile.getMachineInfo()
print("Machine information:", machinfo)

MOCfile.getAllExperiments()
print("Get experiments:",len(MOCfile.experiment))

print("Get all capillary data from all experiments")
MOCfile.getAllCapillaryData()
print("Done")

# or
#mst.experiment[2].getCapillaryData()
# to get the capillary data for the 3rd ([2]) experiment

# Now go through and save experiments to files:

for expno in range(len(MOCfile.experiment)):
    print("Writing experiment #",expno)
    MST.WriteExperimentToXLSX(MOCfile.experiment[expno],"experiment"+str(expno)+".xlsx")
    
print("Finished exporting!")
MOCfile.close()
