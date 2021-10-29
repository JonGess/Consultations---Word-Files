

#Joint AEG/BOPCOM Meeting Oct/Nov 2021 consultation
#-----------------------------------------

import Consultation_Responses_Function
import pandas as pd
import os


#G1
directory = "C:\\Users\\J.Gessendorfer\\Consultation Comments"
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
lastcol = 23
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
GN_name = 'Joint AEG BOPCOM Meeting - D.16'
os.chdir(directory)
data =  pd.read_excel(xlsxfile)
data = data.drop(columns = list(data.columns[4:60]))
data.columns
data = data.drop(columns = list(data.columns[20:76]))
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
xlsxfile = "Joint AEG_BOPCOM Meeting Documents - Oct_Nov 2021 - Responses -  2021-10-28 20-38 44734.xlsx"
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
