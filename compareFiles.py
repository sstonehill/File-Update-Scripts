# -*- coding: utf-8 -*-
from pandas import *
import string
import os
import sys

def compareFiles():
    oldFilePath = r'C:\Users\sstonehill\Documents\Temp Files\choicehotels-yext (29).xl'
    newFilePath = r'C:\Users\sstonehill\Documents\Temp Files\choicehotels-yext (30).xls'

    #oldFilePath = raw_input("Enter old file path:")
    #newFilePath = raw_input("Enter new file path:")
    
    validFile(oldFilePath)
    validFile(newFilePath)
    
    #Reads file extension and imports file using appropriate method
    oldExt = os.path.splitext(oldFilePath)[1]
    newExt = os.path.splitext(newFilePath)[1]
    if oldExt == '.csv': old_data = read_csv(oldFilePath, keep_default_na = False)
    else: old_data = read_excel(oldFilePath, keep_default_na = False)
    if newExt == '.csv': new_data = read_csv(newFilePath, keep_default_na = False)
    else: new_data = read_excel(newFilePath, keep_default_na = False)

    
    #Sets index to the first columnm, saves copy as 'Store ID', and moves column to front
    old_data = old_data.set_index(old_data.columns[0])
    new_data = new_data.set_index(new_data.columns[0])
    old_data['Store ID'] = old_data.index
    new_data['Store ID'] = new_data.index
    old_data = moveToFront(old_data, 'Store ID')
    new_data = moveToFront(new_data, 'Store ID')
    
    #Call error checking functions and exit program if errors are found
    print "Checking files for errors..."
    errors_old = checkIDs(old_data, "old")
    errors_new = checkIDs(new_data, "new")
    if errors_old is True or errors_new is True:
        print "Please fix file errors and try again.\n"        
        sys.exit(1)
    
    #Create added & missing location dataframes
    print "File comparison in progress..."
    added_locs = new_data[~new_data.index.isin(old_data.index)]
    missing_locs = old_data[~old_data.index.isin(new_data.index)]
    
    #Filter dataframes to include only overlapping locations (filter out adds/removes)    
    comp_old_data = old_data[old_data.index.isin(new_data.index)]
    comp_new_data = new_data[new_data.index.isin(old_data.index)] 
    
    #Create changed location dataframe
    cols = ('Store ID', 'Field', 'Old Value', 'New Value')
    changed_locs = DataFrame(columns=cols)
    for col in comp_new_data.columns:
        if col not in ['Store ID']:
            temp_df = comp_new_data[comp_new_data[col]<>comp_old_data[col]]
            temp_df = temp_df.reindex(index=temp_df.index, columns = ('Store ID', col))
            temp_df = temp_df.rename(columns = {col : 'New Value'})
            if len(temp_df.index) > 0: temp_df['Old Value'] = comp_old_data[col]
            temp_df['Field'] = col
            changed_locs = concat([changed_locs, temp_df])
    changed_locs = changed_locs.ix[:, cols]

    print "Analysis complete."     
    
    #Version control on output filename (adds version number if duplicate file exists)
    v = 1
    outputFilePath = os.path.join(os.path.dirname(newFilePath), "File Compare Output.xlsx")
    while os.path.isfile(outputFilePath):
        outputFilePath = os.path.join(os.path.dirname(newFilePath), "File Compare Output ("+str(v)+").xlsx")
        v = v + 1    
    
    #Export data to Excel file
    writer = ExcelWriter(outputFilePath)
    added_locs.to_excel(writer, "New Locations", index=False)
    missing_locs.to_excel(writer, "Removed Locations", index=False)
    changed_locs.to_excel(writer, "Changed Locations", index=False)
    writer.save()
    
    print "Output complete."
    print "Output file path: " + outputFilePath

        
def checkIDs(df, fileStr):
    error_found = False
    blanks = df[df['Store ID']==""]
    
    if len(blanks) > 0:
        error_found = True
        print "Blanks found in " + fileStr + " file Store ID Column."
    
    dupes = filter(None, df.index.get_duplicates())
    if bool(dupes):
        error_found = True
        
        print "Dupes found in " + fileStr + " file Store ID column."
        print "Duplicate IDs: " + str(dupes).replace("u","")

    return error_found

def validFile(filepath):
    if os.path.isfile(filepath) == False:
        print "\nInvalid filepath: " + filepath + "\n"
        sys.exit("Invalid filepath provided. Quitting program.")

def moveToFront(df, colName):
    cols = list(df)
    cols.insert(0, cols.pop(cols.index(colName)))
    df = df.ix[:, cols]
    return df

compareFiles()
    
