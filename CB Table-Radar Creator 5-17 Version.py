#!/usr/bin/env python
# coding: utf-8

# In[263]:


import numpy as np
import pandas as pd
from math import pi
import matplotlib.pyplot as plt
import PySimpleGUI as sg
from bs4 import BeautifulSoup
import pdfkit
from PyPDF2 import PdfFileMerger

# In[276]:

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
    [sg.Text('Input Location of Source File Folder', size=(35, 1))],
    [sg.Text('Source Folder', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default Folder'), sg.FolderBrowse()],
    [sg.Text('Choose a Destination Folder for the Final Results', size=(75, 1))],
    [sg.Text('Destination Folder', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default Folder'), sg.FolderBrowse()],
    [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
]

window = sg.Window('Workplace Assessment Tool', layout, default_element_size=(40, 1), grab_anywhere=False)

event, dirs = window.read()

window.close()

# key = pd.read_excel("7 Forms of Respect Key.xlsx")
# df = pd.read_csv("Workplace Communication Assessment_3.2.2021.csv")
key = pd.read_excel(dirs[1])
df = pd.read_csv(dirs[0])

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
    maxVal = int(key.iloc[2, 2]) + 1
    if (maxVal >= 8):
        ticks = [maxVal / 5] * 4
        ticks = [round(num) for num in ticks]
        ticks = np.multiply(ticks, [1, 2, 3, 4]).tolist()
    else:
        ticks = list(range(1, maxVal))

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
        values = df.T.iloc[i + 1].values.flatten().tolist()
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


def getDetail(dim1Name, dim2Name="", dim3Name=""):
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


def render_mpl_table(data, col_width=3.0, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in mpl_table._cells.items():
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0] % len(row_colors)])
    return ax.get_figure(), ax


# generate all radar plots
i = 0
for user, email in zip(df["First and Last Name"], df['Email']):
    giv_path = "{}/{}_{} GIVE.png".format(dirs[4], email, user)
    get_path = "{}/{}_{} GET.png".format(dirs[4], email, user)
    table = toTable(i)
    giv_fig = toRadar(table[0])
    giv_fig.savefig(giv_path)
    giv_fig.clf()
    get_fig = toRadar(table[1])
    get_fig.savefig(get_path)
    get_fig.clf()
    i = i + 1

# generate all info tables
i = 0
for user, email in zip(df["First and Last Name"], df['Email']):
    table = toTable(i)
    give_fig, give_ax = render_mpl_table(table[0], header_columns=0, col_width=2.0)
    give_fig.savefig("{}/{}_{} GIVE table.png".format(dirs[4], email, user))
    give_fig.clf()
    get_fig, get_ax = render_mpl_table(table[1], header_columns=0, col_width=2.0)
    get_fig.savefig("{}/{}_{} GET table.png".format(dirs[4], email, user))
    get_fig.clf()

options = {
    'minimum-font-size': "20",
    'page-size': "A4",
}

# print GET results pages
for user, email in zip(df["First and Last Name"], df['Email']):
    results = open("{}/results.html".format(dirs[5]))
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
    with open("{}/output.html".format(dirs[5]), "wb") as file:
        file.write(html)
    pdfkit.from_file("{}/output.html".format(dirs[5]), build_path, options=options)

# print GIVE results pages
for user, email in zip(df["First and Last Name"], df['Email']):
    results = open("{}/results.html".format(dirs[5]))
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
    with open("{}/output.html".format(dirs[5]), "wb") as file:
        file.write(html)
    pdfkit.from_file("{}/output.html".format(dirs[5]), build_path, options=options)

# merge all pages together for each user
for user, email in zip(df["First and Last Name"], df['Email']):
    pdfs = [dirs[2], "{}/{}_{} GIVE.pdf".format(dirs[4], email, user),
            "{}/{}_{} GET.pdf".format(dirs[4], email, user), dirs[3]]
    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write("{}/{}_{}.pdf".format(dirs[6], email, user))
    merger.close()
