#!/usr/bin/env python
# coding: utf-8

# In[263]:


import numpy as np
import pandas as pd
from math import pi
import matplotlib.pyplot as plt


# In[276]:


key = pd.read_excel("7 Forms of Respect Key.xlsx")
df = pd.read_csv("Workplace Communication Assessment_3.2.2021.csv")


# In[265]:


qNum = len(key["Question"].dropna())
dim1 = key["Dimension #1 Key"].dropna().unique().tolist()
dim2 = key["Dimension #2 Key"].dropna().unique().tolist()
dim3 = key["Dimension #3 Key"].dropna().unique().tolist()
outKey = key["Output"].dropna().tolist()
outVal = key["Value"].dropna().tolist()
userNum = len(df)


# In[266]:


def tableTemplate(dim3Val):
    output = pd.DataFrame()
    output[dim3Val] = dim1
    for i in range(len(dim2)):
        output[dim2[i]] = 0
    output["Total"] = 0
    return output


# In[267]:


def toTable(index):
    if (len(dim3) > 0):
        output = [pd.DataFrame()] * len(dim3)
        for i in range(len(dim3)):
            output[i] = tableTemplate(dim3[i])
    else: 
        output = [pd.DataFrame()]
        output[0] = tableTemplate("")
    
    for i in range(qNum):
        val = outVal[outKey.index(df[key.Question[i]][index])]
        x1 = dim1.index(key["Dimension #1"][i])
        if (len(dim2) > 0):
            x2 = dim2.index(key["Dimension #2"][i])
        else:
            x2 = 0
        if (len(dim3) > 0):
            x3 = dim3.index(key["Dimension #3"][i])
        else:
            x3 = 0
        
        output[x3].iloc[x1, x2 + 1] = output[x3].iloc[x1, x2 + 1] + val
        if (len(dim2) > 1):
            output[x3].iloc[x1, len(dim2) + 1] = output[x3].iloc[x1, len(dim2) + 1] + val

    return output


# In[268]:


def toRadar(df):
    maxVal = int(key.iloc[2,2]) + 1
    if (maxVal >= 8):
        ticks = [maxVal / 5] * 4
        ticks = [round(num) for num in ticks]
        ticks = np.multiply(ticks, [1,2,3,4]).tolist()
    else:
        ticks = list(range(1,maxVal))

    angles = [n / float(len(dim1)) * 2 * pi for n in range(len(dim1))]
    angles += angles[:1]
    ax = plt.subplot(111, polar=True)

    ax.set_theta_offset(pi / 2)
    ax.set_theta_direction(-1)
    plt.xticks(angles[:-1], dim1)

    # Draw ylabels
    ax.set_rlabel_position(1)
    plt.yticks(ticks, ticks, color="grey", size=7)
    plt.ylim(0, maxVal)

    colors = ['orange', 'b', 'm', 'r', 'g', '#00A9BD', '#EF9526']
    
    if (len(dim2) > 0):
        plotNum = len(dim2)
    else:
        plotNum = 1
    for i in range(plotNum):
        values=df.T.iloc[i+1].values.flatten().tolist()
        values += values[:1]
        if (len(dim2) > 0):
            ax.plot(angles, values, colors[i], linewidth=2, linestyle='solid', label=dim2[i])
        else:
            ax.plot(angles, values, colors[i], linewidth=2, linestyle='solid', label="Results")
        ax.fill(angles, values, colors[i], alpha=.3)
        
    count = 0
    for label, rot in zip(ax.get_xticklabels(), angles):
        if count != 0:
            if count < 4:
                label.set_horizontalalignment("left")
            else:
                label.set_horizontalalignment("right")
        else:
            label.set_horizontalalignment("center")
        count = count + 1

    plt.legend(loc='upper right', bbox_to_anchor=(0, 0.3))
    return plt


# In[332]:


def getDetail(dim1Name, dim2Name = "", dim3Name = ""):
    output = ["", ""]
    d1 = (key["Dimension #1 Key"]) == dim1Name
    if (dim2Name != ""):
        d2 = (key["Dimension #2 Key"]) == dim2Name
    else: 
        d2 = True
    if (dim3Name != ""):
        d3 = (key["Dimension #3 Key"]) == dim3Name
    else: 
        d3 = True
    output[0] = key["Detail #1"][d1 & d2 & d3][1]
    output[1] = key["Detail #2"][d1 & d2 & d3][1]
    return output

