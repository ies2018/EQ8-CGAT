# -*- coding: utf-8 -*-
"""
Created on Mon Sep  3 13:23:49 2018

@author: Brian
"""
import pandas as pd
from pandas import ExcelWriter
import numpy as np
import sys


data_file = "tb_CAT.xlsx"
new_file = "tb_NewCAT.xlsx"

old_cat = pd.read_excel(data_file).fillna('')
new_cat = pd.read_excel(new_file).fillna('')

len_old_cat = len(old_cat)
print data_file + ' loaded'

for i in range (len_old_cat):
    
    if len(new_cat) != len_old_cat:
        print"Files not equal!"
        break
 
    ######################################################    
    ###### change applicability
    if old_cat.AppIndividual[i] > 1:
        new_cat.AppAll[i] = 1
        
    if old_cat.PointSource[i] > 1:
        new_cat.AppAll[i] = 2
        
    if old_cat.Facility[i] > 1:
        new_cat.AppAll[i] = 3
        
    if old_cat.Corporate[i] > 1:
        new_cat.AppAll[i] = 4
    
    ######################################################    
    ###### change Measure
    if old_cat.DirectMeasure[i] > 0:
        new_cat.MeasOverall[i] = 1
        
    if old_cat.IndirectMeasure[i] > 0:
        new_cat.MeasOverall[i] = 3
        
    if old_cat.DirectMeasure[i] or old_cat.IndirectMeasure[i] == 5:
        new_cat.MeasOverall[i] = 5
    
    ####################################################    
    ##### change risk likelihood and impact
    risk = np.array([old_cat.RiskHealth[i],
                   old_cat.RiskEnvironment[i],
                   old_cat.RiskOperation[i]]).max()
    
    #### likelihood
    if risk < 3:
        new_cat.RiskOverall[i] = 1
        
    if risk == 3:
        new_cat.RiskOverall[i] = 3

    if risk >= 4:
        new_cat.RiskOverall[i] = 5
        
    ##### impact
    if risk == 1 or risk == 4:
        new_cat.ImpactOverall[i] = 1
        
    if risk == 2 or risk == 4:
        new_cat.ImpactOverall[i] = 5
        
    if risk == 3:
        new_cat.ImpactOverall[i] = 3
    
    ####################################################    
    ##### change cost ratio
    
    new_cat.CostOverall[i] = np.array([old_cat.CostHealth[i],
               old_cat.CostEnvironment[i],
               old_cat.CostOperation[i]]).max()
    
    ####################################################    
    ##### change Permit required
    
    if old_cat.PermitIndividual[i] > 1:
        new_cat.PermitOverall[i] = 1
        
    if old_cat.PermitPointSource[i] > 1:
        new_cat.PermitOverall[i] = 2
        
    if old_cat.PermitFacility[i] > 1:
        new_cat.PermitOverall[i] = 3
        
    if old_cat.PermitCorporate[i] > 1:
        new_cat.PermitOverall[i] = 4
    
    ####################################################    
    ##### change main receptor
    
    if old_cat.ReceptorIndividual[i] > 1:
        new_cat.ReceptorOverall[i] = 1
        
    if old_cat.ReceptorPointSource[i] > 1:
        new_cat.ReceptorOverall[i] = 2
        
    if old_cat.ReceptorFacility[i] > 1:
        new_cat.ReceptorOverall[i] = 3
        
    if old_cat.ReceptorCorporate[i] > 1:
        new_cat.ReceptorOverall[i] = 4
        
    if old_cat.ReceptorPublic[i] > 1:
        new_cat.ReceptorOverall[i] = 5
 
    ####################################################    
    ##### change responsible
    
    if old_cat.RespIndividual[i] > 1:
        new_cat.RespOverall[i] = 1
        
    if old_cat.RespPointSource[i] > 1:
        new_cat.RespOverall[i] = 2
        
    if old_cat.RespFacility[i] > 1:
        new_cat.RespOverall[i] = 3
        
    if old_cat.RespCorporate[i] > 1:
        new_cat.RespOverall[i] = 4
        
    if old_cat.RespPublic[i] > 1:
        new_cat.RespOverall[i] = 5
    

    ####################################################    
    ##### merge comments
    
    new_cat.Consultant_Comments[i] = (old_cat.Consultant_Comments[i] + '\n' +
                               old_cat.Applicability_Comments[i] + '\n' + 
                               old_cat.Audit_Comments[i] + '\n' +
                               old_cat.Enforce_Comments[i] + '\n' + 
                               old_cat.Measure_Comments[i] + '\n' +
                               old_cat.Permits_Comments[i] + '\n' + 
                               old_cat.Exemptions_Comments[i] + '\n' + 
                               old_cat.CostRatio_Comments[i] + '\n' + 
                               old_cat.Receptor_Comments[i] + '\n' + 
                               old_cat.Responsible_Comments[i] )
    
    ###### print status
    sys.stdout.write('\rUpdated record: ' + str(i) + ' of ' + str(len_old_cat))
    sys.stdout.flush()

# pick column to predict
try:
    new_save = raw_input("Save NewCAT to Excel (y/n)? (default = n): "))
except ValueError:
    new_save = 'n'   

if new_save == 'y':
    
    new_cat_file = 'tb_NewCAT2.xlsx'
    writer = ExcelWriter(new_cat_file)
    new_cat.to_excel(writer,'NewCAT')
    writer.save()
    
    print 'File saved as ' + new_cat_file