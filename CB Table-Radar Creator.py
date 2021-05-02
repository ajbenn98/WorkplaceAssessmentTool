#!/usr/bin/env python
# coding: utf-8

# In[7]:


import numpy as np
import pandas as pd
from math import pi
import matplotlib.pyplot as plt


# In[216]:


def respOut(argument):
    response = {
        "rarely": 1, 
        "sometimes": 2, 
        "almost always": 3
    }
    return response.get(argument, argument)


# In[220]:


key = pd.read_csv("7 Forms of Respect Master Key.csv")
df = pd.read_csv("Workplace Communication Assessment_3.2.2021.csv")


# In[144]:


qNum = len(key["Question"].dropna())
dim1 = key["Dimension #1 Key"].dropna().unique().tolist()
dim2 = key["Dimension #2 Key"].dropna().unique().tolist()
dim3 = key["Dimension #3 Key"].dropna().unique().tolist()
userNum = len(df)


# In[137]:


def tableTemplate(dim3Val):
    output = pd.DataFrame()
    output[dim3Val] = dim1
    for i in range(len(dim2)):
        output[dim2[i]] = 0
    if (len(dim2) > 1):
        output["Total"] = 0
    return output


# In[153]:


def toTable(index):
    output = [pd.DataFrame()] * len(dim3)
    for i in range(len(dim3)):
        output[i] = tableTemplate(dim3[i])
    
    for i in range(qNum):
        x1 = dim1.index(key["Dimension #1"][i])
        x2 = dim2.index(key["Dimension #2"][i])
        x3 = dim3.index(key["Dimension #3"][i])
        val = respOut(df[key.Question[i]][index])
        output[x3].iloc[x1, x2 + 1] = output[x3].iloc[x1, x2 + 1] + val  
        if (len(dim2) > 1):
            output[x3].iloc[x1, len(dim2) + 1] = output[x3].iloc[x1, len(dim2) + 1] + val

    return output


# In[245]:


def toRadar(df):
    maxVal = int(key.iloc[2,2]) + 1
    ticks = [maxVal / 5] * 4
    ticks = [round(num) for num in ticks]
    ticks = np.multiply(ticks, [1,2,3,4]).tolist()

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

    colors = ['#00A9BD', '#EF9526', 'g', 'r', 'b', 'm', 'y']

    for i in range(3):
        values=df.T.iloc[i+1].values.flatten().tolist()
        values += values[:1]
        ax.plot(angles, values, colors[i], linewidth=2, linestyle='solid', label=dim2[i])
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

