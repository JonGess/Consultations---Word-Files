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




#Joint AEG/BOPCOM Meeting Oct/Nov 2021 consultation
#-----------------------------------------
#G1
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - G.1'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)

data = data.drop(columns = list(data.columns[14:132]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 13
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = []
selectall_inmatrix = [] 

data.columns = ['Record ID', 
        '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. Taking into consideration the testing results, do the AEG and the Committee agree that from a practical perspective, it is not feasible to adopt invoice values (Option 3)?',
       '1B. Please elaborate:',
       '2A. Do the AEG and the Committee agree that at this stage, given the implementation difficulties, it is practical to maintain the status quo – Option 2 so as not to disturb the consistent treatment between the SNA and BPM?',
       '2B. Please elaborate:',
       '3A. Do the AEG and the Committee agree with mentioning in the update that although Option 3 is conceptually sound, the status quo is the default recommendation on account of practical implementation difficulties?',
       '3B. Please elaborate:',
       '4A. Would the AEG and the Committee recommend a future collection of Invoice data (through the International Merchandise Trade Statistics (IMTS)) to try to quality assure this series?',
       '4B. Please elaborate:',
       '5. Do you have any other comments on the guidance note?']


#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)


#G2
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - G.2'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:14]))
data.columns
data = data.drop(columns = list(data.columns[12:122]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 11
rankcolumns = [] #this expects a list of lists of columns that belong to the same rank question
matrixcolumns = []
selectall_inmatrix = [] 

data.columns = ['Record ID', 
        '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. Do the AEG and the Committee agree with the proposed definition of MNEs?',
       '1B. Please elaborate:',
       '2A. Do the AEG and the Committee agree to align the SNA with the BPM6 and BD4 on the question of control when defining foreign controlled corporations?',
       '2B. Please elaborate:',
       '3A. Do the AEG and the Committee support the inclusion of the proposed decision tree for allocating MNE units to institutional sectors? ',
       '3B. Please elaborate:',
       '4. Do you have any other comments on the guidance note?',]

#Run function
Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)


#C4
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - C.4'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:22]))
data.columns
data = data.drop(columns = list(data.columns[24:132]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 11
rankcolumns = []
matrixcolumns = []
selectall_inmatrix = [] 

data.columns= ['Record ID', 
        '1.1. ... Name: [Question: Please provide your...]',
        '1.1. ... Email: [Question: Please provide your...]',
        '1.1. ... Affiliation: [Question: Please provide your...]',
        'Do you consider this guidance note ready for publication? ',
        '1A. Do the AEG and the Committee agree with the recommendation to record FGP transactions gross, instead of the current net treatment?',
        '1B. Please elaborate:',
        '2A. Do the AEG and the Committee agree with the recommendation of treating the output of a contractor as services only in the processing-type arrangements, and as goods in FGP-type arrangements?',
        '2B. Please elaborate:',
        '3A. Do the AEG and the Committee agree with the recommendation of considering the definition of FGP activity independent of whether the contractor is an affiliated enterprise or not?',
        '3B. Please elaborate:',
        '4A. Do the AEG and the Committee agree with the recommendation to expand BPM6 coverage of goods to show distinctly the transactions related to goods traded as part of global manufacturing arrangements as supplementary item under general merchandise (with the option to record possible trade with materials inputs in the FGP setup under merchanting as well (Option 2?))',
        '4B. Please elaborate:',
        "5A. In the AEG's and the Committee's opinion, is the decision tree (Annex II) a supportive tool? ",
        '5B. Please elaborate:',
        '6A. Do the AEG and the Committee agree with the reasoning behind the recording of negative exports in merchanting of goods? If not, please specify why. ',
        '6B. Please elaborate:',
        '7A. Do the AEG and the Committee agree with the recommendation that from a pure conceptual view, “merchanting of services” (as gross recording) is impossible? However, services can only be intermediated by a third person against an explicit or implicit fee.',
        '7B. Please elaborate:',
        '8A. Do the AEG and the Committee agree with the recommendation to record this intermediation fees (net), under trade related services as a supplementary “of which” item?',
        '8B. Please elaborate:',
        '9A. Do the AEG and the Committee agree with the recommendation that bundling of services, such as tour operators, should not be recorded as new products, but instead the components should be separately recorded?',
        '9B. Please elaborate:',
        '10. Do you have any other comments on the guidance note?']


Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)




#F9

directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - F.9'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:42]))
data.columns
data = data.drop(columns = list(data.columns[12:94]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 11
rankcolumns = []
matrixcolumns = []
selectall_inmatrix = [] 



data.columns = ['Record ID', '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. What option do the AEG and the Committee favor for the valuation of loans: nominal value (Option 1) or fair value (Option 2)?  ',
       '1B. Please elaborate:',
       '2A. If nominal value (Option 1) is the preferred option, do the AEG and the Committee favor the status-quo of the existing treatment (Option 1a), or its extension allowing for value reset in extraordinary events publicly known (Option 1b)?',
       '2B. Please elaborate:',
       '3A. If fair value (Option 2) is chosen instead of nominal value, would the AEG and the Committee prefer shifting to a full fair value approach (Option 2b) or would the Committees prefer its simplified version (Option 2a) based on the measurement of nominal value less expected loan losses? ',
       '3B. Please elaborate:',
       '4. Do you have any other comments on the guidance note?']


Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)



#D2

directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - D.2'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:50]))
data.columns
data = data.drop(columns = list(data.columns[14:94]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 13
rankcolumns = []
matrixcolumns = []
selectall_inmatrix = [] 


data.columns = ['Record ID', '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. Do the AEG and the Committee agree that cross country comparability within the ESS and consistency across institutional sectors within the SNA would be enhanced by identifying some methods as preferred?',
       '1B. Please elaborate. If yes, what would be your preferred method(s)?  ',
       '2A. Do the AEG and the Committees think that it is necessary to incorporate guidelines on the issues raised in the note (negative equity, treatment of provisions, etc.)? ',
       '2B. Please elaborate. If yes, where, whether in both the revised BPM6 and 2008 SNA or in the BPM7 Compilation Guide?',
       '3A. Do the AEG and the Committee agree with the proposal (Option 2.1) of preparing a separate clarification note on the treatment of negative equity?',
       '3B. Please elaborate:',
       '4A. Do the AEG and the Committee consider a system of information sharing across countries would promote homogeneity in the valuation of unlisted shares worldwide? ',
       '4B. Please elaborate. If yes, which preconditions must be met to implement the system and how could IOs assist the implementation?',
       '5. Do you have any other comments on the guidance note?']

Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)




#D16

directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - D.16'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:60]))
data.columns
data = data.drop(columns = list(data.columns[20:74]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 19
rankcolumns = []
matrixcolumns = []
selectall_inmatrix = [] 

data.columns = ['Record ID', '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. Do the AEG and the Committee agree that BPM6 paragraphs describing retained earnings and reinvested earnings should be revised to reflect the discussion in paragraphs 15 and 16 above? ',
       '1B. Please elaborate:',
       '2A. Do the AEG and the Committee consider it relevant to clarify and address, either in the revised BPM or in the BPM Compilation Guide, shortcomings in the compilation of DI income as described in BPM6 and to include examples of calculation of RIE?',
       '2B. Please elaborate:',
       '3A. Do the AEG and the Committee agree to separate as a memorandum item the obligatory provisions for bad loans when calculating the RIE for credit institutions?',
       '3B. Please elaborate:',
       '4A. Do the AEG and the Committees agree that the recognition of all the earnings generated down the DI ownership chain as primary income is best on a conceptual basis (Alternative A)? ',
       '4B. Please elaborate. If yes, do the AEG and the Committees think the practical challenges encountered by some countries justifies deviating from the conceptually preferred basis Alternative C?',
       '5A. Do the AEG and the Committee prefer to keep the current guidance (Alternative A), do you agree that the presentation proposed in Alternative B (to report indirect income separately) would be useful to enhance transparency and data comparability across countries? ',
       '5B. Please elaborate:',
       '6A. Do the AEG and the Committee agree that RIE and net income should always be compiled regardless of the fund’s attributes? ',
       '6B. Please elaborate:',
       '7A. Do the AEG and the Committee agree with the proposed treatment of operating expenses charged either explicitly or implicitly in the compilation of investment funds’ RIE? ',
       '7B. Please elaborate:',
       '8. Do you have any other comments on the guidance note?']

Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)





#F12
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-25 20-16 41365.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - F.12'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:76]))
data.columns
data = data.drop(columns = list(data.columns[10:74]))
data.columns
name_column = 1
altname_column = 3
firstcol = 4
lastcol = 9
rankcolumns = []
matrixcolumns = []
selectall_inmatrix = [] 

data.columns = ['Record ID', '1.1. ... Name: [Question: Please provide your...]',
       '1.1. ... Email: [Question: Please provide your...]',
       '1.1. ... Affiliation: [Question: Please provide your...]',
       'Do you consider this guidance note ready for publication? ',
       '1A. What option do the AEG and the Committee favor for the classification of hybrid insurance products (Option 1, Option 2, or Option 3)?',
       '1B. Please elaborate:',
       '2A. What option do the AEG and the Committee favor for the classification of autonomous, employer-independent pension schemes (Option 1, Option 2, or Option 3)?',
       '2B. Please elaborate:',
       '3. Do the AEG and the Committee have any other comments/suggestions on the issues discussed in the GN?']

Consultation_Responses_Function.responses_in_word(data, directory, xlsxfile, GN_name, name_column, altname_column, firstcol, lastcol, rankcolumns, matrixcolumns, selectall_inmatrix)


