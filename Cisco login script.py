#import modules, pandas for dataframe, pyperclip for clipboard access
import pandas as pd
import pyperclip as pc


#Importing CSV of grid for passcode
grid = pd.read_csv(r'C:\Users\Q1071018\Documents\Grid.csv', index_col = 0)

#copy password to login prompt
pw = 'Lakitu.3833'
pc.copy(pw)
print('Copied password.')

#Grab user input for code
print('Input code:')
Codeinput = input()
print('Inputted: ' + Codeinput)


#[A1] [A2] [A3]
#Assign characters of user input to variables
#Convert strings of numbers to integers for indexing with pandas
b1 = Codeinput[1]
b1_2 = int(Codeinput[2])
b2 = Codeinput[6]
b2_2 = int(Codeinput[7])
b3 = Codeinput[11]
b3_2 = int(Codeinput[12])

#Locate values at intersections of dataframe using input variables in (row, column) format
output = (str(grid.loc[b1_2, b1]) + 
          str(grid.loc[b2_2, b2]) +  
          str(grid.loc[b3_2, b3]) )

pc.copy(str.lower(output))
print('Code copied to clipboard: ' + str.lower(output))


