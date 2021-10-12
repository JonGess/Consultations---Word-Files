# -*- coding: utf-8 -*-
"""
Created on Tue Oct  5 14:39:20 2021

@author: J.Gessendorfer
"""


import Consultation_Responses_Function
import pandas as pd

#WS.2 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "WS.2 Distribution of household income, consumption and wealth - Responses -  2021-10-6 14-57 35132.xlsx"
GN_name = 'WS.2 Distribution of household income, consumption and wealth'
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 35
rankcolumns = [[26,27,28]] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns= []
#Run function
Consultation_Responses_Function.responses_in_word(directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns)


#WS.3 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "WS.3 Unpaid Household Service Work - Responses -  2021-10-12 16-28 11519.xlsx"
GN_name = 'WS.3 Unpaid Household Service Work'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 45
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(12,(33+1)))]

#Run function
Consultation_Responses_Function.responses_in_word(directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns)



#WS.4 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "WS.4 Labour, Human Capital and Education - Responses -  2021-10-12 16-25 34266.xlsx"
GN_name = 'WS.4 Labour, Human Capital and Education'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 40
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(10,(17+1)))]

#Run function
Consultation_Responses_Function.responses_in_word(directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns)




#WS.6 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "WS.6 Accounting for the Economic Ownership and Depletion of Natural Resources - Responses -  2021-10-12 16-26 12722.xlsx"
GN_name = 'WS.6 Accounting for the Economic Ownership and Depletion of Natural Resources'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 31
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = []

#Run function
Consultation_Responses_Function.responses_in_word(directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns)


