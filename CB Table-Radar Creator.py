#!/usr/bin/env python
# coding: utf-8

# In[7]:


import numpy as np
import pandas as pd
from math import pi
import matplotlib.pyplot as plt
import PySimpleGUI as sg
import dataframe_image as dfi
from bs4 import BeautifulSoup
import pdfkit
from PyPDF2 import PdfFileMerger

# In[216]:

sg.ChangeLookAndFeel('GreenTan')

# ------ Menu Definition ------ #
menu_def = [['File', ['Open', 'Save', 'Exit', 'Properties']],
            ['Edit', ['Paste', ['Special', 'Normal', ], 'Undo'], ],
            ['Help', 'About...'], ]

layout = [
    [sg.Text('Choose the Data File', size=(35, 1))],
    [sg.Text('Data CSV', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default File'), sg.FileBrowse()],
    [sg.Text('Choose the Answer Key File', size=(35, 1))],
    [sg.Text('Answer Key', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default File'), sg.FileBrowse()],
    [sg.Text('Choose the Preface Page(s) File', size=(35, 1))],
    [sg.Text('Preface PDF', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default File'), sg.FileBrowse()],
    [sg.Text('Choose the End Page(s) File', size=(35, 1))],
    [sg.Text('Endpage PDF', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default File'), sg.FileBrowse()],
    [sg.Text('Input Location of Build Folder', size=(35, 1))],
    [sg.Text('Build Folder', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default Folder'), sg.FolderBrowse()],
    [sg.Text('Choose a Destination Folder for the Final Results', size=(75, 1))],
    [sg.Text('Destination Folder', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default Folder'), sg.FolderBrowse()],
    [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
]

window = sg.Window('Workplace Assessment Tool', layout, default_element_size=(40, 1), grab_anywhere=False)

event, dirs = window.read()

window.close()

def respOut(argument):
    response = {
        "rarely": 1, 
        "sometimes": 2, 
        "almost always": 3
    }
    return response.get(argument, argument)


# In[220]:


key = pd.read_csv(dirs[1])
df = pd.read_csv(dirs[0])
# key = pd.read_excel(dirs[1], "AnswerSheet")
# df = pd.read_excel(dirs[0], "RawData")

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

i = 0
for user in df["First and Last Name"]:
    giv_path = "{}/{} GIVE.png".format(dirs[3], user)
    table = toTable(i)[0]
    # dfi.export(table, "{}/{} GIVE table.png".format(dirs[2], user), table_conversion='matplotlib')
    giv_fig = toRadar(table)
    giv_fig.savefig(giv_path)
    giv_fig.clf()
    i = i + 1

i = 0
for user in df["First and Last Name"]:
    final_path = "{}/{} GET.png".format(dirs[3], user)
    table = toTable(i)[1]
    # dfi.export(table, "{}/{} GET table.png".format(dirs[2], user), table_conversion='matplotlib')
    fig = toRadar(table)
    fig.savefig(final_path)
    fig.clf()
    i = i + 1

# for user, email in (df["First and Last Name"], df["Email"]):
#     results = open("{}/results.html".format(dirs[3]))
#     soup = BeautifulSoup(results)
#     final_path = "{}/{}_{}.pdf".format(dirs[2], email, user)
#     giv_path = "{}/{} GIVE.png".format(dirs[3], user)
#     get_path = "{}/{} GIVE.png".format(dirs[3], user)
#     giv_chart = soup.find(id="give-plot")
#     giv_chart['src'] = giv_path
#     get_chart = soup.find(id="get-plot")
#     get_chart['src'] = get_path
#     soup.find(id="give-detail").string.replace_with("INSERT GIVE DETAIL HERE")
#     soup.find(id="get-detail").string.replace_with("INSERT GET DETAIL HERE")
#     desc = soup.find_all(id="lang-descript")
#     for d in desc:
#         d.string.replace_with("INSERT DESCRIPTIONS HERE")
#     results.close()
#     html = soup.prettify("utf-8")
#     with open("{}/output.html".format(dirs[3]), "wb") as file:
#         file.write(html)
#     pdfkit.from_file("{}/output.html".format(dirs[3]), final_path)
