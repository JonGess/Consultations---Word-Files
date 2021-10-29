# -*- coding: utf-8 -*-
"""
Created on Tue Oct  5 14:39:20 2021

@author: J.Gessendorfer
"""


import Consultation_Responses_Function
import pandas as pd
import os

#WS.2 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "WS.2 Distribution of household income, consumption and wealth - Responses -  2021-10-6 14-57 35132.xlsx"
GN_name = 'WS.2 Distribution of household income, consumption and wealth'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 35
rankcolumns = [[26,27,28]] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns= []
selectall_inmatrix = [] 

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)


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
selectall_inmatrix = [""] 

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)



#WS.4 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
os.chdir(directory)
xlsxfile = "WS.4 Labour, Human Capital and Education - Responses -  2021-10-12 16-25 34266.xlsx"
GN_name = 'WS.4 Labour, Human Capital and Education'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 40
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(10,(17+1)))]
selectall_inmatrix = [""] 

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)




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
selectall_inmatrix = [] 

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)



#DZ.5 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "DZ.5 Digital SUTs - Responses -  2021-10-12 21-54 38284.xlsx"
GN_name = 'DZ.5 Digital SUTs'
data =  pd.read_excel(xlsxfile)
name_column = 1
altname_column = 3
firstcol = 6
lastcol = 25
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = []
selectall_inmatrix = [] 

data = data.rename(columns={'3. Please provide arguments in favor of your response:': '2B. Please provide arguments in favor of your response:',
                            '6.1. If yes, what technical assistance, if any, would you need?': '6B. If yes, what technical assistance, if any, would you need?'})


#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)



#G.2 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
os.chdir(directory)
xlsxfile = "G.2 Treatment of MNE and Intra MNE Flows - Responses -  2021-10-12 22-35 50665.xlsx"
GN_name = 'G.2 Treatment of MNE and Intra MNE Flows'
data =  pd.read_excel(xlsxfile)
name_column = 3
altname_column = 1
firstcol = 6
lastcol = 86
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(15,(22+1))), list(range(33,(67+1)))] 

#select all that apply categorys in matrixquestions:
#the length of this object should be the same as matrixcolumns.
#if there is no select all that apply category in matrix 0, the string should be empty in the 0th position
#for matrices that have a select all that apply category, the name of the category should be used.
selectall_inmatrix = ["", "Main challenges"] 

data = data.rename(columns={'22. Additional comments, if any, to elaborate on the decision tree in Question 21:\xa0': '21B. Additional comments, if any, to elaborate on the decision tree in Question 21:',
                            '24. If you responded "Strongly disagree" or "Disagree" to Question 23, please specify why:': '23B. If you responded "Strongly disagree" or "Disagree" to Question 23, please specify why:',
                            '26. Please provide an explanation for your response to Question 25:': '25B. Please provide an explanation for your response to Question 25:',
                            '28. Please provide an explanation for your response to Question 27:': '27B. Please provide an explanation for your response to Question 27:',
                            '30. Please provide an explanation for your response to Question 29:':'29B. Please provide an explanation for your response to Question 29:',
                            '32. Please provide an explanation for your response to Question 31:': '31B. Please provide an explanation for your response to Question 31:',
                            '34. Please provide an explanation for your response to Question 33:': '33B. Please provide an explanation for your response to Question 33:',
                            '36. Please provide an explanation for your response to Question 35:': '35B. Please provide an explanation for your response to Question 35:',
                            '16. Please provide an explanation for your response to Question 15:': '15B. Please provide an explanation for your response to Question 15:',
                            '13. If you responded "Strongly disagree" or "Disagree" to Question 12, please specify why:': '12B. If you responded "Strongly disagree" or "Disagree" to Question 12, please specify why:',
                            '5. Please provide an explanation for your response to Question 4:':'4B. Please provide an explanation for your response to Question 4:',
                            '8. If you responded "Strongly disagree" or "Disagree" to Question 7, please specify why:': '7B. If you responded "Strongly disagree" or "Disagree" to Question 7, please specify why:',
                            '10. If you responded "Strongly disagree" or "Disagree" to Question 9, please specify why:': '9B. If you responded "Strongly disagree" or "Disagree" to Question 9, please specify why:',
                            '20. Please provide an explanation for your response to Question 19:':'19B. Please provide an explanation for your response to Question 19:'})



#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)




#G.4 User Input
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "G.4 Treatment of Special Purpose Entities and Residency - Responses -  2021-10-12 23-11 3494.xlsx"
GN_name = 'G.4 Treatment of Special Purpose Entities and Residency'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
name_column = 2
altname_column = 1
firstcol = 5
lastcol = 41
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(21,(24+1))), list(range(25,(26+1)))]
selectall_inmatrix = ["", ""] 


data = data.rename(columns={'5. Please provide an explanation for your response to Question 4:': '4B. Please provide an explanation for your response to Question 4:',
                            '9. If you responded "Strongly disagree" or "Disagree" to Question 8, please specify why:': '8B. If you responded "Strongly disagree" or "Disagree" to Question 8, please specify why:',
                            '11. If you responded "Strongly disagree" or "Disagree" to Question 10, please specify why:': '10B. If you responded "Strongly disagree" or "Disagree" to Question 10, please specify why:',
                            '13. If you responded "Strongly disagree" or "Disagree" to Question 12, please specify why:': '12B. If you responded "Strongly disagree" or "Disagree" to Question 12, please specify why:',
                            '15. If you responded "Strongly disagree" or "Disagree" to Question 14, please specify why:':'14B. If you responded "Strongly disagree" or "Disagree" to Question 14, please specify why:',
                            '17. Please specify how:': '16B. Please specify how:',
                            '26. If you responded "Strongly disagree" or "Disagree" to Question 23, please specify why:': '25B. If you responded "Strongly disagree" or "Disagree" to Question 23, please specify why:',
                            '28. If you responded "Strongly disagree" or "Disagree" to Question 25, please specify why:': '27B. If you responded "Strongly disagree" or "Disagree" to Question 25, please specify why:',
                            '30. If you responded "Strongly disagree" or "Disagree" to Question 27, please specify why:': '29B. If you responded "Strongly disagree" or "Disagree" to Question 27, please specify why:',
                            '32. If you responded "Strongly disagree" or "Disagree" to Question 29, please specify why:': '31B. If you responded "Strongly disagree" or "Disagree" to Question 29, please specify why:'})

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)



import Consultation_Responses_Function
import pandas as pd
import os
#G1
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Valuation of imports and exports of goods in the international standards_v2 - C - Responses -  2021-10-26 21-29 14212.xlsx"
GN_name = 'G.1 Valuation of imports and exports of goods in the international standards - Part C'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
name_column = 2
altname_column = 3
firstcol = 6
lastcol = 69
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = [list(range(9,(38+1))), list(range(39,(68+1)))]
selectall_inmatrix = ["", ""] 


Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)







