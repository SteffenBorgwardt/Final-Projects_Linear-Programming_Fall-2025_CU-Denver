# -*- coding: utf-8 -*-

from gurobipy import *
import xlrd
import xlwt
import sys
from datetime import *
from xlrd import xldate_as_tuple
import pandas as pd

#Read excel data
data = xlrd.open_workbook("Supply chain logisitcs problem.xls")
table = data.sheets()[0]
table1 = data.sheets()[1]
table2 = data.sheets()[2]
table3 = data.sheets()[3]
table4 = data.sheets()[4]
nrows = table.nrows 
ncols = table.ncols 
nrows1 = table1.nrows 
ncols1 = table1.ncols 
nrows2 = table2.nrows 
ncols2 = table2.ncols 
nrows3 = table3.nrows 
ncols3 = table3.ncols 
nrows4 = table4.nrows 
ncols4 = table4.ncols 
Orders={}
Plant_cap_cost={}
Plant_PID={}
Plant_port={}
Carriers={}

#Put order list information in Order dictory
for i in range(1,nrows):
    row = []
    for j in range(2,ncols):
        cell = table.cell_value(i,j)
        row.append(cell)
    Orders[tuple([table.cell_value(i,0),table.cell_value(i,1)])]=row

#Group the order ID from the product ID
for i in Orders.keys():
    group = [i]
    PID_orders = 1
    PID_units = Orders[i][0]
    PID_weights = Orders[i][1]
    for j in Orders.keys():
        if i!=j and i[0] == j[0]:
            PID_orders = PID_orders+1
            PID_units = PID_units+Orders[j][0]
            PID_weights = PID_weights+Orders[j][1]
            group.append(j)
    Orders[i].append(PID_orders)
    Orders[i].append(PID_units)
    Orders[i].append(PID_weights)
    Orders[i].append(group)

#Organize the information of capacity and cost per plant
for i in range(1, nrows1):
    row = []
    for j in range(1,ncols1):
        cell = table1.cell_value(i,j)
        row.append(cell)
    Plant_cap_cost[table1.cell_value(i,0)]=row

#Specific product ID is assigned to a specific plant or plants
for i in range(1, nrows2):
    row = []
    for j in range(1,ncols2):
        cell = table2.cell_value(i,j)
        row.append(cell)
        if table2.cell_value(i,0) in Plant_PID.keys():
            row.extend(Plant_PID[table2.cell_value(i,0)])
    Plant_PID[table2.cell_value(i,0)]=row

for i in Orders.keys():
    for j in Plant_PID.keys():
        if i[0]==j:
            Orders[i].append(Plant_PID[j])

#Specific plant is assigned to a specific port or ports
for i in range(1, nrows4):
    row = []
    for j in range(1,ncols4):
        cell = table4.cell_value(i,j)
        row.append(cell)
        if table4.cell_value(i,0) in Plant_port.keys():
            row.extend(Plant_port[table4.cell_value(i,0)])
    Plant_port[table4.cell_value(i,0)]=row

#Append the port information to the each order according to the plant information 
for i in Orders.keys():
    ports=[]
    for j in Orders[i][7]:
        if len(Plant_port[j])>1:
            for k in range(len(Plant_port[j])):
                ports.append([j, Plant_port[j][k]])
        else:
            ports.append([j, Plant_port[j][0]])                
    Orders[i].append(ports)

# Organize carrier information
a = 1
for i in range(1,nrows3):
    row = []
    for j in range(4,ncols3):
        cell = table3.cell_value(i,j)
        row.append(cell)
    Carriers[tuple([table3.cell_value(i,0),table3.cell_value(i,1),table3.cell_value(i,2),table3.cell_value(i,3), a])]=row
    a = a + 1

#Calculate the min cost for different levels of weight for each of Carrier 
Carriers_min={}
for i in Carriers.keys():
    if type(Carriers[i][4]) == float and Carriers[i][4]>0:
        for j in Carriers.keys():
            if i!=j and i[0]==j[0] and i[1]==j[1] and i[2]==j[2] and i[3]==j[3]:
                    x= Carriers[i][1]+Carriers[i][2]*i[3] 
                    y= Carriers[j][1]+Carriers[j][2]*j[3]
                    if x < y:
                        Carriers_min[tuple([i[0], i[1], i[2], i[3]])]=Carriers[i]                 
                    else:
                        Carriers_min[tuple([i[0], i[1], i[2], i[3]])]=Carriers[j]

#Single-plant orders assign to the min-cost of transport                     
for i in Orders.keys():
    if len(Orders[i][7]) == 1:
        Plant_cap_cost[Orders[i][7][0]][0]=Plant_cap_cost[Orders[i][7][0]][0]-1        
        Total_weight=0
        Total_unit=0
        for j in Orders.keys():
            if i!=j and Orders[j][7] == Orders[i][7]:
                Total_weight=Total_weight+Orders[j][1]
                Total_unit=Total_unit+Orders[j][0]
        Orders[i].append(Total_unit)
        Orders[i].append(Total_weight)
min_tcost=999999
index={}
for i in Orders.keys():
    if len(Orders[i][7]) == 1:
        for j in Orders[i][8]:
            for p in Carriers_min.keys():
                if j[1]==p[1] and Orders[i][10]<= p[3] and Orders[i][10]>= p[2]:                   
                    tcost = Carriers_min[p][1] + Carriers_min[p][2] * Orders[i][10]
                    if tcost < min_tcost:
                        min_tcost = tcost
                        index[p]=Carriers_min[p]
        
        plant_cost=Orders[i][9] * Plant_cap_cost[Orders[i][7][0]][1]
        Orders[i].append(1)                                         # Indicate this order is completed
        Orders[i].append(Orders[i][7][0])                           # Append the Plant information
        Orders[i].append(p[1])                                      # Append the Port information
        Orders[i].append(p)                                         # Append the Carrier type
        Orders[i].append(Carriers_min[p])                           # Append the Carrier information
        Orders[i].append(plant_cost)                                # Append total Plant cost of grouped product
        Orders[i].append(min_tcost)                                 # Append the min cost of transportation
    else:
        Orders[i].append(" ")
        Orders[i].append(" ")
        Orders[i].append(0)                                        #Indicate this order is not completed

#Multiple-plant orders 
Orders_PID={}
for i in Orders.keys():
    Orders_PID[i[0]]= Orders[i]
C={}
N={}
T={}
W={}
for i in Orders_PID.keys():
    if Orders_PID[i][11] == 0 and Orders_PID[i][10]==' ':
        for p in Orders_PID[i][7]:
            if Plant_cap_cost[p][0]>0:
                for t in Orders_PID[i][8]:
                    if p==t[0]:
                        for m in Carriers_min.keys():
                            if t[1]==m[1] and Orders_PID[i][1]<= m[3]:
                                C[i, p, t[1], m]=Orders_PID[i][4]*Plant_cap_cost[p][1]+ Orders_PID[i][5]*Carriers_min[m][2]+ Carriers_min[m][1]
                                T[i, p, t[1], m]=Orders_PID[i][3]
                                W[i, p, t[1], m]=Orders_PID[i][5]
                                N[p, t[1], m,]=Plant_cap_cost[p][0]

#Model build to minimize the plant cost and transportation cost for each grouped product 
model = Model()
x = model.addVars(C.keys(), obj=C, vtype=GRB.BINARY, name='x')
y = model.addVars(N.keys(), vtype=GRB.BINARY, name='y')
        
#Constraint 1: all orders have to be placed and delivered.
for i in Orders_PID.keys():
    if Orders_PID[i][11] == 0:
        model.addConstr(x.sum(i, '*', '*', '*') == 1)
      
#Constraint 2: the capacity limitation of the plant.
for i in N.keys():            
    model.addConstr(x.prod(T, '*', 'i[0]', '*', '*') <= N[i]*y[i])

#Constraint 3: the capacity limitation of m-th Carrier and transporter 
for i in N.keys():
    model.addConstr(x.prod(W,'*', '*', '*', i[2]) <= i[2][3]*y[i])
    model.addConstr(x.prod(W,'*', '*', '*', i[2]) >= i[2][2]*y[i])      
        
      
#model.setParam(GRB.Param.LogToConsole, 0)    
model.setParam(GRB.Param.TimeLimit, 10) 
model.optimize() 

#Print optimized Plant, port, Carrier and related information
for i in C.keys():
    if x[i].x != 0:
        Orders_PID[i[0]][11] = 2
        Orders_PID[i[0]].append(i[1])
        Orders_PID[i[0]].append(i[2])
        Orders_PID[i[0]].append(i[3])
        Orders_PID[i[0]].append(Carriers_min[i[3]])
        Orders_PID[i[0]].append(Orders_PID[i[0]][4] * Plant_cap_cost[i[1]][1])
        Orders_PID[i[0]].append(Orders_PID[i[0]][5] * Carriers_min[i[3]][2]+Carriers_min[i[3]][1])

for i in Orders.keys():
    Orders[i]=Orders_PID[i[0]]

#Write out information to excel
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Logistic solution')
worksheet.write(0, 0, label = 'Product ID')
worksheet.write(0, 1, label = 'Order ID')
worksheet.write(0, 2, label = 'Unit quantity')
worksheet.write(0, 3, label = 'weight')
worksheet.write(0, 4, label = 'Destination port')
worksheet.write(0, 5, label = 'Plant code')
worksheet.write(0, 6, label = 'Origin Port')
worksheet.write(0, 7, label = 'Carrier type')
worksheet.write(0, 8, label = 'minm_wgh_qty')
worksheet.write(0, 9, label = 'max_wgh_qty')
worksheet.write(0, 10, label = 'Service Level')
worksheet.write(0, 11, label = 'Plant cost for group of Product')
worksheet.write(0, 12, label = 'Transportation cost for group of Product')

j = 1   
for i in Orders.keys():
    worksheet.write(int(j), 0, label=i[0])
    worksheet.write(int(j), 1, label=i[1])
    worksheet.write(int(j), 2, label=Orders[i][0])
    worksheet.write(int(j), 3, label=Orders[i][1])
    worksheet.write(int(j), 4, label=Orders[i][2])
    worksheet.write(int(j), 5, label=Orders[i][12])
    worksheet.write(int(j), 6, label=Orders[i][13])
    worksheet.write(int(j), 7, label=Orders[i][14][0])
    worksheet.write(int(j), 8, label=Orders[i][14][2])
    worksheet.write(int(j), 9, label=Orders[i][14][3])
    worksheet.write(int(j), 10, label=Orders[i][15][0])
    worksheet.write(int(j), 11, label=Orders[i][16])
    worksheet.write(int(j), 12, label=Orders[i][17])       
    j=j+1

workbook.save('Logistic solution.xls')



 

























