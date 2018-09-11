# -*- coding: utf-8 -*-
"""
Created on 11 Sep 2018

@author: Brian
Utility to write group VBA subroutine code for Access DB
"""

varlist = ['SafetyProcess', 'SafetyIndividual', 'ProductStewardship', 
           'Documentation', 'HR', 'Procurement', 'IT', 'Finance', 'Legal', 
           'Security', 'Quality', 'Audit', 'Other']

list_len = len(varlist)

sec_text = 'Dim lngRed, lngBlack As Long\n'
thrd_text = '    lngRed = RGB(255, 0, 0)\n' 
frth_text = '    lngBlack = RGB(0, 0, 0)\n'
eig_text = '    Else\n'
ten_text = '    End If\n\n'
elev_text = 'End Sub\n\n'

output_file = 'vba-out-checkbox.txt'





text2 = '        countQ = countQ + 1\n'
text3 = '            If countQ = 1 Then\n'
text4 = '                strJoin = ""\n'
text5 = '            Else\n'
text6 = '                strJoin = "OR "\n'
text7 = '            End If\n'

####### initialize VBA strings
vba_string1 = ''
vba_string2 = ''
vba_string3 = ''
vba_string4 = ''
vba_string5 = ''


########### make VBA strings
for i in range(list_len):
    first_text = 'Private Sub Check' + varlist[i] + '_AfterUpdate() \n'
    six_text = '    If Me.Check' + varlist[i] + ' = True Then\n'
    sev_text = '        Me.Label' + varlist[i] + '.ForeColor = lngRed\n'
    nin_text = '        Me.Label' + varlist[i] + '.ForeColor = lngBlack\n'
    text1 = '    ##### ' + varlist[i] + ' ####\n'
    text8 = '        str' + varlist[i] + ' = strJoin & "((tb_reg_master.' + varlist[i] + ')=True) "\n'
    text9 = '    strTot = strTot & str' + varlist[i] + '\n\n'
    
    vba_string1 += (first_text + sec_text + thrd_text + 
                    frth_text + six_text + sev_text + eig_text + 
                    nin_text + ten_text + elev_text)
    
    vba_string2 += (six_text + sev_text + eig_text + nin_text + ten_text)
    
    vba_string3 += '    Me.Check' + varlist[i] + ' = False\n'
    
    vba_string4 += '    Me.Label' + varlist[i] + '.ForeColor = lngBlack\n'
    
    vba_string5 += (text1 + six_text + text2 + text3 + text4 + 
                   text5 + text6 + text7 + text8 + ten_text + text9)   
    

'''
# open output file and write text
with open(output_file, 'w') as file:
    #filedata = filedata.replace(varlist[i], varlist[i+1])
    file.write(filedata)
file.close()
'''
