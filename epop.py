###############################################################################
###                                 ePop                                    ###
###                           (excel populator)                             ###
###                            Amanda Galemmo                               ###
###############################################################################

import os, docInfo as di
from pandas import read_excel
from functools import reduce
cellList = []
eDF = read_excel('testSheet.xlsx', sheet_name='Sheet1').astype(str)

                                    #######
                                    # Run #
                                    #######
while(True):
    print('Hello. Please enter \'Y\' to continue.')
    if input(): break

# Create a list of directories to work through.
dirList = []
for (dirpath, dirnames, filenames) in os.walk(os.getcwd()):
    dirList.extend(dirnames)

for dir in dirList:
    os.chdir(dir)
    print('\n##############################')
    print('\nWorking through: ' + os.getcwd())
    print('\n##############################\n')

    cellList.extend(list(map(lambda x: di.getInfo(x),
                         filter(lambda x: x.endswith('.docx'),
                                os.listdir(os.getcwd())))))
    os.chdir('..')

# Insert those cells into eDF
for c in cellList:
    try:
        row = eDF[eDF['MARKET'] == c.get('MARKET')].index[0]
    # So some of the documents in the folders are not contracts and therefore
    # don't go through getInfo() properly, this just skips over them.
    except AttributeError:
        print('AttributeError')
        continue
    insert = str(c.get('TITLE')) + ' ' + str(c.get('STATION'))

    if (eDF.at[row, c.get('TITLE')] == 'nan'):
        eDF.at[row, c.get('TITLE')] = insert
        print('\tInserted ' + insert + ' at ' + c.get('MARKET') + ', '
              + c.get('TITLE'))
    else:
        if (eDF.at[row, c.get('TITLE')].find('||') > -1):
            if (insert != eDF.at[row, c.get('TITLE')].split(' || ')[0]
                and insert != eDF.at[row, c.get('TITLE')].split(' || ')[1]):
                eDF.at[row, c.get('TITLE')] = (eDF.at[row, c.get('TITLE')]
                                                + ' || ' + insert)
                print('\tInserted ' + insert + ' at ' + c.get('MARKET') + ', '
                      + c.get('TITLE'))
        elif (eDF.at[row, c.get('TITLE')].strip() != insert):
            eDF.at[row, c.get('TITLE')] = (eDF.at[row, c.get('TITLE')]
                                            + ' || ' + insert)
            print('\tInserted ' + insert + ' at ' + c.get('MARKET') + ', '
                  + c.get('TITLE'))
        else: print('MOVING ON')

print('\nEnd of loop.')
print('epop_ouput.xlsx can be found in ' + os.getcwd())
eDF.to_excel('epop_output.xlsx')
