from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, Border, Alignment
from copy import copy

wb = load_workbook('test2.xlsx')
ws = wb['Sheet1']

#initiating indexes
cols=0
row=0
i=0
row_data = []
sheetNames = ['Sheet2', 'Sheet3', 'Sheet4'] #to name the new sheets created

#we want to iterate over the entire dataset in Sheet1
for row in range(0,39):
    for cols in range(0,10):
        #if the value of any cell is the same as the string 'ITEM1' run the following
        if ws['C5:L45'][row][cols].value=='ITEM1':
            ws2 = wb.create_sheet(sheetNames[i]) #create new sheet upon finding 'ITEM1'
            ws2 = wb[sheetNames[i]]
            i+=1
            print("New Sheet with label " +sheetNames[i]+ " created")
            #to copy the data, we must re-align the data to its original cell
            for row_2 in range(row,row+11):
                for cols_2 in range(cols,cols+10):
                    #print("Exporting data " + "("+str(row)+","+str(cols)+")" + " from Sheet1 to " + "("+str(row_2)+","+str(cols_2)+")" + " in " + sheetNames[i-1])
                    ws2['C5:L45'][row_2-row][cols_2-cols].value = ws['C5:L45'][row_2][cols_2].value
                    #copy styling
                    ws2['C5:L45'][row_2-row][cols_2-cols].font = copy(ws['C5:L45'][row_2][cols_2].font)
                    ws2['C5:L45'][row_2-row][cols_2-cols].fill = copy(ws['C5:L45'][row_2][cols_2].fill)
                    ws2['C5:L45'][row_2-row][cols_2-cols].border = copy(ws['C5:L45'][row_2][cols_2].border)
                    ws2['C5:L45'][row_2-row][cols_2-cols].number_format = copy(ws['C5:L45'][row_2][cols_2].number_format)
        #if the string 'ITEM2' is found, then we want to add this with the previous data of 'ITEM1' (with the same date)
        if ws['C5:L45'][row][cols].value=='ITEM2':
            for row_2 in range(row,row+11):
                for cols_2 in range(cols,cols+10):
                    ws2['C19:L45'][row_2-row][cols_2-cols].value = ws['C5:L45'][row_2][cols_2].value
                    #copy styling
                    ws2['C19:L45'][row_2-row][cols_2-cols].font = copy(ws['C5:L45'][row_2][cols_2].font)
                    ws2['C19:L45'][row_2-row][cols_2-cols].fill = copy(ws['C5:L45'][row_2][cols_2].fill)
                    ws2['C19:L45'][row_2-row][cols_2-cols].border = copy(ws['C5:L45'][row_2][cols_2].border)
                    ws2['C19:L45'][row_2-row][cols_2-cols].number_format = copy(ws['C5:L45'][row_2][cols_2].number_format)

#save it into a new file
wb.save('test4.xlsx')