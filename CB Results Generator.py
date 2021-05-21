#!/usr/bin/env python
# coding: utf-8


import numpy as np
import pandas as pd
from math import pi
import matplotlib.pyplot as plt
import PySimpleGUI as sg
from bs4 import BeautifulSoup
import pdfkit
from PyPDF2 import PdfFileMerger

sg.ChangeLookAndFeel('GreenTan')

# ------ Menu Definition ------ #
menu_def = [['File', ['Open', 'Save', 'Exit', 'Properties']],
            ['Edit', ['Paste', ['Special', 'Normal', ], 'Undo'], ],
            ['Help', 'About...'], ]

layout = [
    [sg.Text('Choose the Data File', size=(35, 1))],
    [sg.Text('Data Excel File', size=(15, 1), auto_size_text=False, justification='right'),
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

window = sg.Window('CB Results Generator', layout, default_element_size=(40, 1), grab_anywhere=False)

event, dirs = window.read()

window.close()

key = pd.read_excel(dirs[1])
df = pd.read_excel(dirs[0])


qNum = len(key["Question"].dropna())
dim1 = key["Dimension #1 Key"].dropna().unique().tolist()
dim2 = key["Dimension #2 Key"].dropna().unique().tolist()
dim3 = key["Dimension #3 Key"].dropna().unique().tolist()
outKey = key["Output"].dropna().tolist()
outVal = key["Value"].dropna().tolist()
userNum = len(df)


def tableTemplate(dim3Val):
    output = pd.DataFrame()
    output[dim3Val] = dim1
    for i in range(len(dim2)):
        output[dim2[i]] = 0
    output["Total"] = 0
    return output


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
    output[0] = key["Detail #1"][d1 & d2 & d3].iloc[0]
    output[1] = key["Detail #2"][d1 & d2 & d3].iloc[0]
    return output


def get_top_lang(df):
    return df.iloc[:, 0][df["Total"].idxmax()]


def get_lang_defs():
    output = [[""] * len(dim1), [""] * len(dim1)]
    for i in range(len(dim1)):
        output[0][i] = dim1[i]
        d1 = (key["Dimension #1 Key"]) == dim1[i]
        output[1][i] = next(item for item in key["Description"][d1])
    return output


def get_titles():
    output = [""] * 5
    output[0] = key["Assessment Details"][0]
    for i in range(4):
        output[i + 1] = key["Assessment Details"][i + 4]
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
def generate_radar_imgs(dim3_num):
    z = 0
    for user_radar, email_radar in zip(df["First and Last Name"], df['Email']):
        plot_path = "{}/{}_{} {}.png".format(dirs[4], email_radar, user_radar, dim3_num)
        table = toTable(z)
        plot_fig = toRadar(table[dim3_num])
        plot_fig.savefig(plot_path)
        plot_fig.clf()
        plt.close()
        z = z + 1


# generate all numeric table images
def generate_table_imgs(dim3_num):
    b = 0
    for user_t, email_t in zip(df["First and Last Name"], df['Email']):
        table = toTable(b)
        table_fig, give_ax = render_mpl_table(table[dim3_num], header_columns=0, col_width=2.0)
        table_fig.savefig("{}/{}_{} {} table.png".format(dirs[4], email_t, user_t, dim3_num))
        table_fig.clf()
        b = b + 1


options = {
    'minimum-font-size': "20",
    'page-size': "A4",
    'enable-local-file-access': "",
}


# generates a single result page
def generate_pdf_result(dim3_num):
    w = 0
    for user_p, email_p in zip(df["First and Last Name"], df['Email']):
        print("Printing result page #{} for user {} out of {} ({})".format(dim3_num + 1, w + 1,
                                                                   len(df["First and Last Name"]), user_p))
        results = open("{}/results.html".format(dirs[5]))
        soup = BeautifulSoup(results)
        build_path = "{}/{}_{} {}.pdf".format(dirs[4], email_p, user_p, dim3_num)
        img_path = "{}/{}_{} {}.png".format(dirs[4], email_p, user_p, dim3_num)
        tbl_path = "{}/{}_{} {} table.png".format(dirs[4], email_p, user_p, dim3_num)
        all_titles = get_titles()
        title = "{}'s {}".format(user_p, all_titles[0])
        subtitle = all_titles[dim3_num + 1]
        giv_chart = soup.find(id="plot")
        giv_chart['src'] = img_path
        giv_chart = soup.find(id="table")
        giv_chart['src'] = tbl_path
        top_lang = get_top_lang(toTable(w)[dim3_num])
        details = getDetail(dim1Name=top_lang, dim3Name=dim3[dim3_num])
        soup.find(id="title").string.replace_with(title)
        soup.find(id="subtitle").string.replace_with(subtitle)
        soup.find(id="top-lang").string.replace_with("Your top language for this section is <strong>{}</strong>"
                                                     .format(top_lang.upper()))
        soup.find(id="top-desc").string.replace_with(details[0])
        soup.find(id="top-question").string.replace_with(details[1])
        desc = get_lang_defs()
        for d in range(0, len(desc[0])):
            html_id = "lang-descript-{}".format(d)
            lang_def = "<strong>{}:</strong> {}".format(desc[0][d], desc[1][d])
            soup.find(id=html_id).string.replace_with(lang_def)
        results.close()
        w = w + 1
        html = soup.prettify("utf-8", formatter=None)
        with open("{}/output.html".format(dirs[5]), "wb") as file:
            file.write(html)
        pdfkit.from_file("{}/output.html".format(dirs[5]), build_path, options=options)


# generate all radar plots
for g in range(0, len(dim3)):
    generate_radar_imgs(g)

# generate all info tables and generate single-page result PDFs
for g in range(0, len(dim3)):
    generate_table_imgs(g)
    generate_pdf_result(g)

# merge all pages together for each user
for user, email in zip(df["First and Last Name"], df['Email']):
    pdfs = [dirs[2], dirs[3]]
    for k in range(0, len(dim3)):
        pdfs.insert(1, "{}/{}_{} {}.pdf".format(dirs[4], email, user, k))
    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(pdf, import_bookmarks=False)
    merger.write("{}/{}_{}.pdf".format(dirs[6], email, user))
    merger.close()

print("Results for all users have been generated! Operation complete.")
