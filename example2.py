import MSTProcess as MST

MOCfile = MST.openMOCFile('Your_filename_here.moc')

# Get machine info
machinfo = MOCfile.getMachineInfo()
print("Machine information:", machinfo)

print("Number of experiments:",len(MOCfile.experiment))

# Go through and save experiments to files:

for expno in range(len(MOCfile.experiment)):
    print("Writing experiment #",expno)
    MST.WriteExperimentToXLSX(MOCfile.experiment[expno],"experiment"+str(expno)+".xlsx")

# alternatively use "MOCfile.SaveXLSX()" to save all experiments (see example.py)

print("Finished exporting!")
MOCfile.close()
