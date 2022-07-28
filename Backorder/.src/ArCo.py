from pickle import TRUE
from select import select
import pandas as pd


#########################################################################################
#                  Automate the backlog for mac team                                    #
#    Get the values from two excel files and output a new file with mathched data       #
#########################################################################################

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# ---------------------------------------------------------------      Data Structure    ----------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# df_1                        
# df_2
# df_x1
# df_x2
# matched_orders
# first_list
# second_list
# select_rowdf1
# select_fromdf2
# unmatched_orders

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------     Get data from files     -------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


df1 = pd.read_excel('C:\\Backorder\\Work with files\\compare.xlsx')      # Get first file 
first_list = df1['ORDER'].tolist()                                       # Need only the values of the first column
 
df2 = pd.read_excel('C:\\Backorder\\Work with files\\main.xlsx')         # Get the second file  
second_list = df2['ORDER'].tolist()                                      # Need only the values of the first column


matched_orders = []                                                             # Empty container  
unmatched_orders = []                                                           # Empty container  

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------       Built functions    ----------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

                                  
def matching_values():
    global new_list                               # define it as a gloabl variable in case i might need to work with it in a second function 
    for element in first_list:                    # get all the elements that are defined in the first list
        if element in second_list:                # check to see if they exist in the second list 
            matched_orders.append(element)        # dump them in the new container 
        else:                                     # if the order numbe doesn`t exist
            unmatched_orders.append(element)      # dump in the unmatched container 
    
    select_rowdf1 = df1.loc[df1['ORDER'].isin(matched_orders) ]     # get the rows from data frame 1 where the condition is an iterable  
    select_fromdf2 = df2.loc[df2['ORDER'].isin(matched_orders)]     # get the rows from data frame 1 where the condition is an iterable
    unmatch_row = df1.loc[df1['ORDER'].isin(unmatched_orders)]      # get all the rows from data frame 1
    # Write in a new exel file all the data that we dumped in the new container  

    df_1 = pd.DataFrame(select_rowdf1)     # Create the result from the first data frame where the promis date is 
    df_2 = pd.DataFrame(select_fromdf2)    # Create the result from the second data frame where the structure is 
    df_3 = pd.DataFrame(unmatch_row)       # create the structure 
    
    # sort the df_1 & df_2 (smallest to largest) this way columns will match on both tables
    df_x1 = df_1.sort_values(by=['ORDER'])
    df_x2 = df_2.sort_values(by=['ORDER'])
    df_x3 = df_3.sort_values(by=['ORDER'])  
    
    # I want the dataframe column that has empty values to get the values of the same column from data frame 1
    df_x2['NEW PROMISE/DELIVERY/TIME'] = df_x1['NEW PROMISE/DELIVERY/TIME'].values    
    
    #Create two excel files and write df_x2 as masterfile and df_x3 as unmatched  
    print(df_x3) 
    writer_match = pd.ExcelWriter('C:\\Backorder\\Work with files\\masterfile.xlsx')                         # specify the path 
    writer_unmatch = pd.ExcelWriter('C:\\Backorder\\Work with files\\unmatched.xlsx')
    df_x2.to_excel(writer_match, sheet_name='Sheet1', index=False)
    df_x3.to_excel(writer_unmatch, sheet_name='Sheet1', index=False)                                           # write data frame to excel file
    writer_match.save()                                                                                      # save the file
    writer_unmatch.save() 
matching_values()


#-------------------------------------------------------------      Writen by: Ardit Cuko    -------------------------------------------------------------------------------------------------#