# -*- coding: utf-8 -*-
"""
Created on Wed Aug 29 14:46:07 2018

@author: Brian
"""
from IPython import get_ipython

try:
    var_reset = raw_input("Reset all variables? (y/n)? (Default = n)? ")
except ValueError:
    var_reset = "n"
    
if var_reset == "y":
    get_ipython().magic('reset -sf')
    
    
import pandas as pd
import numpy as np
import time
import sys
import pickle

############################################
##### set output filename
odms_filename = 'ODMS_analysis.xlsx'    #example output file
stop_filename = 'stop_file_points.txt'

############################################
#### load regulations
if not 'regs' in globals():
    regs = pd.read_excel("tb_reg_master.xlsx")
    len_regs = len(regs)
    print 'tb_reg_master.xlsx loaded',

### find applicable regulations
if not 'app_regs' in globals():
    app_regs = regs[(regs.Applicable == True) & 
                    (regs.Header == False) & 
                    (regs.Definition == False) & 
                    (regs.Administrative == False)].reset_index()
    
    len_app_regs = len(app_regs)
    print'# of applicable regs = ', len_app_regs
    
    #prepare reg dataframe with categories only
    df1_reg = app_regs.iloc[:,1].copy()
    df2_reg = app_regs.iloc[:,12:51].copy()
    df3_reg = pd.concat([df1_reg,df2_reg], axis=1)+0
    
    regs_np = df3_reg.values   # convert to np matrix
    
############################################
### load ODMS functions
if not 'odms' in globals():
    odms = pd.read_excel("tb_odms.xlsx")
    len_odms = len(odms)
    print 'tb_odms.xlsx loaded',  
    
### find applicable odms functions
if not 'app_odms' in globals():
    app_odms = odms[(odms.APP == 1)].reset_index()
    len_app_odms = len(app_odms)
    print'# of applicable regs = ', len_app_odms
    
    #prepare odms dataframe with categories only
    df1_odms = app_odms.iloc[:,1].copy()
    df2_odms = app_odms.iloc[:,10:47].copy()
    df3_odms = pd.concat([df1_odms,df2_odms], axis=1)+0
    
    odms_np = df3_odms.values # convert to np matrix

############################################    
######## make or load odms analysis matrix
if not 'odms_analysis' in globals():
    
    try:
        odms_load = raw_input("Make new ODMS Reg matrix (y/n)? (Default = n)? ")
    except ValueError:
        odms_load = "n"
        
    ##### make a new ODMS analysis matrix (all zeroes) if yes
    if odms_load == "y":
        odms_analysis = np.zeros((len_odms, len_regs), dtype=int) ## numpy matrix!
        
        #### reset stopping points if a new matrix is created
        L2 = [0,0] 
        with open(stop_filename, 'wb') as F:
            pickle.dump(L2, F)
    
    ##### load existing matrix if 'no'
    else:    
        odms_analysis = pd.read_excel(odms_filename).values  ## must be a numpy matrix!
        print '\n' + odms_filename + ' loaded'

#******************************************************        
############### load previous stopping points
try:
    stop_pts = raw_input("Load previous stopping points? (y/n)? (Default = n)? ")
except ValueError:
    stop_pts = "n"

i_prev, j_prev = 0, 0  ### set default values 
   
if stop_pts == "y":
    with open (stop_filename, 'rb') as F:
        L2 = pickle.load(F)
        
        #### assign previous stopping point values        
        i_prev, j_prev = L2

################################################################        
############## check to see if there is a different start point

try:
    i_start =  int(raw_input("What ODMS ID (i) to start at? (default = " + str(i_prev) + ")? "))
except ValueError:
    i_start =  i_prev

try:
    j_start =  int(raw_input("What Reg ID (j) to start at? (default = " + str(j_prev) + ")? "))
except ValueError:
    j_start =  j_prev

################################################################
### set counter to check if a break is called
### breaks from loop occur if a number > 2 is typed

break_err = 0

for i in range(i_start, len_app_odms):
    
    i_app = app_odms.ID[i]-1  ## set odms index to match applicable odms function
    if break_err == 1:
        break
    
    else:
        counter = 0  ### set up counter
        for j in range(j_start, len_app_regs):
            j_app = app_regs.ID[j]-1  ## set regs index to match applicable regs
            
            #### check to see if the odms function and reg element share common features
            a = odms_np[i,1:] + regs_np[j,1:]
            b = (a == 2)+0
            if b.sum() == 0:
                print "(Func ID: " + str(app_odms.ID[i]) + ") ODMS:", app_odms.comments[i]
                print "\n(Reg ID: " + str(app_regs.ID[j]) + ")Text:", app_regs.TextEng[j]
                print"Did not match...."
                break                            
            
            ### print element information and get element status
            
            print "\n", j, "Section:", app_odms.Section[i]
            print "(Func ID: " + str(app_odms.ID[i]) + ") ODMS:", app_odms.comments[i]
            print "Current value is", odms_analysis[i_app,j_app]
            
            if counter == 0:
                print "\n", i, "************* Regulation Source:", app_regs.TitleEn[j]
                
            print "\n(Reg ID: " + str(app_regs.ID[j]) + ")Text:", app_regs.TextEng[j]

            #*****************************************************
            ##### assign odms function analysis value
            try:
                temp_odms =  int(raw_input("N/A-0, Indirect-1, or Direct-2 Ignore entire function-3 (default = 0 N/A)? "))
            except ValueError:
                temp_odms =  0
            
            ### if input = 3, break only the regs and go to the next odms function
            if temp_odms == 3:
                break
            
            ### if input > 3, break completely 
            if temp_odms > 3:
                break_err = 1
                break
            
            else:
                if odms_analysis[i_app,j_app] > 0:      ### check to see if there is a value
                    print'\nValue already exists (' + str(odms_analysis[i_app,j_app]) + ')'
                    
                    try:
                        overwrite = raw_input("Overwrite value (y/n)? (Default = n)? ")
                    except ValueError:
                        overwrite = "n"  ### if no value (other than 0) update the element
                        
                    if overwrite == 'y':
                        odms_analysis[i_app,j_app] = temp_odms
                        
                else:
                    odms_analysis[i_app,j_app] = temp_odms
                                 
            print'Value at Reg ID', app_regs.ID[j], 'and Func ID', app_odms.ID[i], ' set to ', odms_analysis[i_app,j_app]
            print'\n_________________________________'
            
            counter += 1  ### increment counter
            
        print '\n******** New Regulation element *********'
                         
####################################################################
######  write results to file
i = i-1
if i<0:
    i=0

j = j-1
if j<0:
    j=0

print '\n Stopped at i =', i , 'and j =', j

try:
    file_write = raw_input("Save ouput file (y/n)? (Default = n)? ")
except ValueError:
    file_write = "n"

if file_write == "y":   
    writer = pd.ExcelWriter(odms_filename)
    pd.DataFrame(odms_analysis).to_excel(writer,'ODMS Reg matrix',index = False) #save prediction output
    writer.save()
    print'File saved in ' + odms_filename

######### save stopping point
try:
    file_write = raw_input("Save stopping points to file (y/n)? (Default = n)? ")
except ValueError:
    file_write = "n"
    
L = [i,j]

if file_write == "y":   
    with open(stop_filename, 'wb') as F:
    # Dump the list to file
        pickle.dump(L, F)
        
    print'File saved in ' + stop_filename

