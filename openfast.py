# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 13:51:48 2018

@author: Brian
"""
import pickle

stop_filename = 'stop_file_points.txt'


with open (stop_filename, 'rb') as F:
    L2 = pickle.load(F)
    
print L2
    
