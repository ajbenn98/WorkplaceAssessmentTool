#!/usr/bin/env Python3
# coding: utf-8

# In[2]:


import numpy as np
import pandas as pd
import dataframe_image as dfi
from math import pi
# import matplotlib as mpl
import matplotlib.pyplot as plt
import PySimpleGUI as sg
from bs4 import BeautifulSoup
import pdfkit
from PyPDF2 import PdfFileMerger
# import matplotlib.font_manager as font_manager
# import os

# path = os.path.join(mpl.rcParams["datapath"], "fonts/ttf/RobotoCondensed.ttf")
# prop = font_manager.FontProperties(fname=path)
# plt.rcParams['font.family'] = prop.get_name()

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
    return response.get(argument, lambda: "Invalid Response Name")


def langOut(argument):
    language = {
        "Engagement": 1,
        "Consideration": 2,
        "Candor": 3,
        "Order": 4,
        "Recognition": 5,
        "Time Commitment": 6,
        "Transparency": 7
    }
    return language.get(argument, lambda: "Invalid Language Name")


def langIn(argument):
    language = {
        1: "Engagement",
        2: "Consideration",
        3: "Candor",
        4: "Order",
        5: "Recognition",
        6: "Time Commitment",
        7: "Transparency"
    }
    return language.get(argument, lambda: "Invalid Language ID")


def levOut(argument):
    level = {
        "Senior": 10,
        "Peer": 20,
        "Junior": 30
    }
    return level.get(argument, lambda: "Invalid Level Name")


def levIn(argument):
    level = {
        1: "Senior",
        2: "Peer",
        3: "Junior"
    }
    return level.get(argument, lambda: "Invalid Level ID")


def dirOut(argument):
    direction = {
        "Give": 100,
        "Get": 200
    }
    return direction.get(argument, lambda: "Invalid Direction")


def dirIn(argument):
    response = {
        1: "Give",
        2: "Get"
    }
    return response.get(argument, lambda: "Invalid Direction ID")


# In[3]:


key = pd.read_excel(dirs[1], "AnswerSheet")
df = pd.read_excel(dirs[0], "RawData")

question_num = len(key)
users = len(df)

output = [0] * question_num
for q in range(question_num):
    output[q] = levOut(key.Level[q]) + dirOut(key.Direction[q]) + langOut(key.Language[q])

# In[4]:


group = dict()

for i in range(len(output)):
    if output[i] in group:
        group[output[i]].append(i)
    else:
        group[output[i]] = [i]

allVal = []
for u in range(users):
    userVal = dict()
    for key in group:
        localTotal = 0
        for i in range(len(group[key])):
            localTotal += respOut(df.iloc[:, group[key][i] + 6][u])
        userVal[key] = localTotal
    allVal.append(userVal)


# In[5]:


def toTable(index):
    give = pd.DataFrame()
    get = pd.DataFrame()
    languages = []
    for i in range(7):
        languages.append(langIn(i + 1))
    give["Give"] = languages
    get["Get"] = languages
    giveValues = []
    getValues = []

    for lev in range(3):
        totals = []
        for lang in range(7):
            totals.append(allVal[index][100 + (lev + 1) * 10 + (lang + 1)])
        giveValues.append(totals)
    give["Seniors"] = giveValues[0]
    give["Peers"] = giveValues[1]
    give["Juniors"] = giveValues[2]
    give["Total"] = [sum(col) for col in zip(*giveValues)]

    for lev in range(3):
        totals = []
        for lang in range(7):
            totals.append(allVal[index][200 + (lev + 1) * 10 + (lang + 1)])
        getValues.append(totals)
    get["Seniors"] = getValues[0]
    get["Peers"] = getValues[1]
    get["Juniors"] = getValues[2]
    get["Total"] = [sum(col) for col in zip(*getValues)]

    return give, get


# In[6]:


def toRadar(df):
    categories = df.iloc[:, 0]
    types = ["Senior", "Peer", "Junior"]
    N = len(categories)

    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    ax = plt.subplot(111, polar=True)

    ax.set_theta_offset(pi / 2)
    ax.set_theta_direction(-1)
    plt.xticks(angles[:-1], categories)

    # Draw ylabels
    ax.set_rlabel_position(1)
    plt.yticks([2, 4, 6, 8], [2, 4, 6, 8], color="grey", size=7)
    plt.ylim(0, 10)

    colors = ['b', 'or', 'g']

    for i in range(3):
        values = df.T.iloc[i + 1].values.flatten().tolist()
        values += values[:1]
        ax.plot(angles, values, linewidth=2, linestyle='solid', label=types[i])
        ax.fill(angles, values, colors[i], alpha=.1)

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

    plt.legend(loc='upper right', bbox_to_anchor=(0, .3))
    return plt


# In[60]:


# index = 7
# print(df["First and Last Name"][index])
# print(toTable(index)[0])
# print(toTable(index)[1])
#
# # In[12]:
#
#
# print(df["First and Last Name"])
#
# # In[16]:
#
#
# i = 7
# print(df["First and Last Name"][i])
# tables = toTable(i)
#
# # In[59]:
#
#
# print(toTable(i)[0])
#
# # In[18]:
#
#
# print(toTable(i)[1])

i = 0
for user, email in zip(df["First and Last Name"], df['Email']):
    giv_path = "{}/{}_{} GIVE.png".format(dirs[4], email, user)
    table = toTable(i)[0]
    dfi.export(table, "{}/{}_{} GIVE table.png".format(dirs[4], email, user), table_conversion='matplotlib')
    giv_fig = toRadar(table)
    giv_fig.savefig(giv_path)
    giv_fig.clf()
    i = i + 1

i = 0
for user, email in zip(df["First and Last Name"], df['Email']):
    final_path = "{}/{}_{} GET.png".format(dirs[4], email, user)
    table = toTable(i)[1]
    dfi.export(table, "{}/{}_{} GET table.png".format(dirs[4], email, user), table_conversion='matplotlib')
    fig = toRadar(table)
    fig.savefig(final_path)
    fig.clf()
    i = i + 1

options = {
    'minimum-font-size': "20",
    'header-font-size': "24",
    'page-size': "Letter",
}

for user, email in zip(df["First and Last Name"], df['Email']):
    results = open("{}/results.html".format(dirs[4]))
    soup = BeautifulSoup(results)
    build_path = "{}/{}_{} GET.pdf".format(dirs[4], email, user)
    img_path = "{}/{}_{} GET.png".format(dirs[4], email, user)
    tbl_path = "{}/{}_{} GET table.png".format(dirs[4], email, user)
    title = "{}'s 7 Forms of Respect™ Results".format(user)
    subtitle = "How You Expect to Get Respect"
    giv_chart = soup.find(id="plot")
    giv_chart['src'] = img_path
    giv_chart = soup.find(id="table")
    giv_chart['src'] = tbl_path
    soup.find(id="title").string.replace_with(title)
    soup.find(id="subtitle").string.replace_with(subtitle)
    # soup.find(id="give-detail").string.replace_with("INSERT GIVE DETAIL HERE")
    # soup.find(id="get-detail").string.replace_with("INSERT GET DETAIL HERE")
    # desc = soup.find_all(id="lang-descript")
    # for d in desc:
    #     d.string.replace_with("INSERT DESCRIPTIONS HERE")
    results.close()
    html = soup.prettify("utf-8")
    with open("{}/output.html".format(dirs[4]), "wb") as file:
        file.write(html)
    pdfkit.from_file("{}/output.html".format(dirs[4]), build_path, options=options)

for user, email in zip(df["First and Last Name"], df['Email']):
    results = open("{}/results.html".format(dirs[4]))
    soup = BeautifulSoup(results)
    build_path = "{}/{}_{} GIVE.pdf".format(dirs[4], email, user)
    img_path = "{}/{}_{} GIVE.png".format(dirs[4], email, user)
    tbl_path = "{}/{}_{} GIVE table.png".format(dirs[4], email, user)
    title = "{}'s 7 Forms of Respect™ Results".format(user)
    subtitle = "How You Give Respect"
    giv_chart = soup.find(id="plot")
    giv_chart['src'] = img_path
    giv_chart = soup.find(id="table")
    giv_chart['src'] = tbl_path
    soup.find(id="title").string.replace_with(title)
    soup.find(id="subtitle").string.replace_with(subtitle)
    # soup.find(id="give-detail").string.replace_with("INSERT GIVE DETAIL HERE")
    # soup.find(id="get-detail").string.replace_with("INSERT GET DETAIL HERE")
    # desc = soup.find_all(id="lang-descript")
    # for d in desc:
    #     d.string.replace_with("INSERT DESCRIPTIONS HERE")
    results.close()
    html = soup.prettify("utf-8")
    with open("{}/output.html".format(dirs[4]), "wb") as file:
        file.write(html)
    pdfkit.from_file("{}/output.html".format(dirs[4]), build_path, options=options)

for user, email in zip(df["First and Last Name"], df['Email']):
    pdfs = [dirs[2], "{}/{}_{} GIVE.pdf".format(dirs[4], email, user),
            "{}/{}_{} GET.pdf".format(dirs[4], email, user), dirs[3]]
    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write("{}/{}_{}.pdf".format(dirs[5], email, user))
    merger.close()