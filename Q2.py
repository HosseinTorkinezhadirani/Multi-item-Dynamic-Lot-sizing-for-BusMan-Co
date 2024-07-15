#!/usr/bin/env python
# coding: utf-8

# In[8]:


#------------------------------------------------------- Question 2 ------------------------------------------------------------
from pyomo.environ import *
import pandas as pd

parameters_df = pd.read_excel("Data.xlsx", sheet_name='Parameter',index_col=0)
demand_df = pd.read_excel("Data.xlsx", sheet_name='Demand',index_col=0)


# In[9]:


parameters_df.head(50)


# In[10]:


demand_df.head(50)


# In[11]:


model = ConcreteModel()

product_indx = parameters_df.index.tolist()
period_indx = demand_df.columns

model.x = Var(product_indx,period_indx,within = NonNegativeIntegers)
x = model.x

model.I = Var(product_indx,period_indx,within = NonNegativeIntegers)
I = model.I

model.y = Var(period_indx,within = NonNegativeIntegers)
y = model.y

model.obj = Objective(expr = sum(160*y[t] for t in period_indx) + sum(parameters_df.loc[i,"Holding Cost (€)"]*I[i,t] 
                                                                      for i in product_indx for t in period_indx))

model.constraints = ConstraintList()

print("Prodcuts = " , product_indx)
print("Periods = " , period_indx)
print("\n ----------------------------------------------------------------- \n")

for i in product_indx:
    model.constraints.add(expr =  I[i,1] == x[i,1] - demand_df.loc[i,1] + parameters_df.loc[i,"Initial Stock"])

for i in product_indx:
    for t in period_indx:
        if t > 1:
            model.constraints.add(expr =  I[i,t] == x[i,t] - demand_df.loc[i,t] + I[i,t-1])
            
for i in product_indx:
    for t in period_indx:   
         model.constraints.add(expr = I[i,t] >= parameters_df.loc[i,"Safety Stock"])

for t in period_indx:   
         model.constraints.add(expr = sum((1/(parameters_df.loc[i,"Number of items that fit each pallet"]*
                                                parameters_df.loc[i,"Number of pallets that fit a truck"]))
                                           *x[i,t] for i in product_indx) <= y[t])
for i in product_indx:
    for t in period_indx:   
         model.constraints.add(expr = x[i,t] <= parameters_df.loc[i,"Supplier limit"])

            
solver = SolverFactory('gurobi')
solver.solve(model)
#display(model)
model.pprint()


# In[12]:


x_solution = pd.DataFrame(index=product_indx, columns=period_indx, dtype=int)
I_solution = pd.DataFrame(index=product_indx, columns=period_indx, dtype=int)

for i in product_indx:
    for t in period_indx:
        x_solution.at[i, t] = x[i, t].value
        I_solution.at[i, t] = I[i, t].value

y_solution = pd.Series(index=period_indx, dtype=int)
for t in period_indx:
    y_solution[t] = y[t].value

with pd.ExcelWriter('model_solutions_Q2.xlsx') as writer:
    x_solution.to_excel(writer, sheet_name='X_Solution')
    I_solution.to_excel(writer, sheet_name='I_Solution')
    y_solution.to_excel(writer, sheet_name='Y_Solution', header=True)

print("Solutions saved to model_solutions_Q2.xlsx")


# In[13]:


for i in product_indx:
    print("\n")
    print("product" , i , ":")
    for t in period_indx:
        print("Week ", t ,":")
        print("\t",f'x[{i},{t}] =', x[i,t].value)
        print("\t",f'I[{i},{t}] =', I[i,t].value)
        print("\t",f'y[{t}] =', y[t].value)
        print("\n")
    print("-----------------------------------------------")


# In[14]:


print("Total Cost:" ,round(value(model.obj),3))
print("Total Transportation Cost: " , sum(160*value(y[t]) for t in period_indx))
print("Total Holding Cost" , round(sum(parameters_df.loc[i,"Holding Cost (€)"]*value(I[i,t]) 
                                 for i in product_indx for t in period_indx),3))
print("Total Truck Utilization: " , round(sum((100/(parameters_df.loc[i,"Number of items that fit each pallet"]
                                              *parameters_df.loc[i,"Number of pallets that fit a truck"]))
                    *value(x[i,t]) for i in product_indx for t in period_indx)/(sum(value(y[t]) for t in period_indx)),3),"%")

b = []
for t in period_indx:
    b.append(round(sum((100/(parameters_df.loc[i,"Number of items that fit each pallet"]
                                              *parameters_df.loc[i,"Number of pallets that fit a truck"]))
                    *value(x[i,t]) for i in product_indx)/value(y[t]),3))

print("\nTruck utilization of each week: " , b )   

a = 0
for t in period_indx:
    a = a + b[t-1]

print("\nAverage Utilizarion is: " , round(a/12,3), "%")    


# In[ ]:




