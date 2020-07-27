import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter import Tk
import numpy as np
import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
import matplotlib.pyplot as plt

start = pd.Timestamp.now()
# Read in file

Tk().withdraw() #hide tk window
master_df_filename = askopenfilename(initialdir='W:/storyboards/',
                               filetypes=(("Excel File", "*.xlsx"), ("All Files", "*.*")),
                               title="Choose the relevant master file."
                               )

master_df = pd.read_excel(master_df_filename)
master_df['Report Date'] = pd.to_datetime(master_df['Report Date'], format='%Y-%m-%d')
print(master_df['Report Date'].value_counts())


#TODO separate to individual directorate
def df_cutter(df, directorate):
    """Chops up the master_df into specific chunks for given part of org structure"""
    df = df[df['Sector/Directorate/HSCP'] == directorate]
    return df


def page_1_piv(df):

    piv = pd.pivot_table(df, index='Report Date', values=['Head Count', 'WTE', 'Bank WTE', 'Overtime WTE', 'Excess WTE',
                                                          'Additional WTE'], aggfunc=np.sum).round(0)

    piv.reset_index(inplace=True)
    piv.sort_values('Report Date', ascending=False, inplace=True)
    piv['Report Date'] = piv['Report Date'].dt.strftime('%b-%y')
    return piv

def page_2_piv(df):
    piv = pd.pivot_table(df, index='Report Date', values=['WTE', 'Starters WTE', 'Leavers WTE'], aggfunc=np.sum).round(0)
    piv = piv.round(1)

    piv['Starters %'] = (piv['Starters WTE'] / piv['WTE'] * 100).round(1)
    piv['Leavers %'] = (piv['Leavers WTE'] / piv['WTE'] * 100).round(1)
    piv.reset_index(inplace=True)
    piv.sort_values('Report Date', ascending=False, inplace=True)
    piv['Report Date'] = piv['Report Date'].dt.strftime('%b-%y')
    return piv

def page_3_piv(df):
    print(df.columns)
    piv = pd.pivot_table(df, index='Report Date', values=['Absence Percent Short', 'Absence Percent Long',
                                                          'Absence Percent', 'Annual Percent', 'Maternity Percent',
                                                          'Paternal Percent', 'Parental Percent',
                                                          'Public Holiday Percent', 'Study Percent', 'Special Percent',
                                                          'Other Percent'
                                                          ], aggfunc = np.sum).round(2)
    piv = piv.rename(columns={
        'Absence Percent Short':'Sickabs Short %', 'Absence Percent Long':'SickAbs Long %',
        'Absence Percent':'Sickness Abs %', 'Annual Percent':'Annual Leave %', 'Maternity Percent':'Maternity %',
        'Paternal Percent':'Paternity %', 'Parental Percent': 'Parental %',
        'Public Holiday Percent':'Public Holiday %', 'Study Percent':'Study %', 'Special Percent':'Special %',
        'Other Percent':'Other %'
    })
    piv.reset_index(inplace=True)
    piv.sort_values('Report Date', ascending=False, inplace=True)
    piv['Report Date'] = piv['Report Date'].dt.strftime('%b-%y')
    return piv

def page_1_bargraph(piv, sector):
    print(piv.columns)
    headcount = piv['Head Count'].tolist()
    WTE = piv['WTE'].astype(int).tolist()
    months = piv['Report Date'].tolist()
    WTE_avg = sum(WTE) / len(WTE)
    hc_avg = sum(headcount) / len(headcount)
    for i in [months, headcount, WTE]:
        i.reverse()


    x = np.arange(len(months))
    width = 0.35

    fig, ax = plt.subplots()
    rects1 = ax.bar(x - width/2, headcount, width, label = 'Headcount')
    rects2 = ax.bar(x + width/2, WTE, width, label='WTE')

    ax.set_title(sector + ' WTE & Headcount')
    ax.set_xticks(x)
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    ax.legend()

    def autolabel(rects, avg, colour, label):
        for rect in rects:
            height = rect.get_height()
            ax.annotate('{}'.format(height),
                        xy = (rect.get_x() + rect.get_width() / 2, height),
                        xytext = (0,3),
                        textcoords = "offset points",
                        ha='center', va='bottom')
        plt.hlines(avg, -1, len(months), colors=colour, linestyles='dashed', label=label)
    autolabel(rects1, hc_avg, 'C0', 'Headcount Average')
    autolabel(rects2, WTE_avg, 'C1', 'WTE Average')
    fig.tight_layout()
    plt.savefig('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-bar.png')

def page_1_stackedbar(piv, sector):
    bank = piv['Bank WTE'].tolist()
    excess = piv['Excess WTE'].tolist()
    ot = piv['Overtime WTE'].tolist()
    additional = piv['Additional WTE'].tolist()
    months = piv['Report Date'].tolist()



    for i in [bank, excess, ot, additional, months]:
        i.reverse()

    bars1 = np.add(bank, ot).tolist()
    bars2 = np.add(bars1, excess).tolist()


    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, bank, width, label='Bank WTE')
    ax.bar(months, ot, width, label='Overtime WTE', bottom=bank)
    ax.bar(months, excess, width, label='Excess WTE', bottom=bars1)
    ax.bar(months, additional, width, label='Additional WTE', bottom=bars2)
    ax.set_title(sector + ' Supplementary WTE')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    ax.legend()
    plt.savefig('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-stackedbar.png')



df = df_cutter(master_df, 'eHealth')
print(page_1_piv(df))
print(page_2_piv(df))
print(page_3_piv(df))



#TODO page 1 tables



#TODO page 1 graphs


def pdfmaker(sector):
    df = df_cutter(master_df, sector)
    doc = SimpleDocTemplate("W:/Storyboards/Test_Boards/"+sector+'.pdf', rightMargin=10, leftMargin=10,
                            topMargin=10, bottomMargin=10, pagesize=(A4[1], A4[0]))
    main_pdf = []
    styles = getSampleStyleSheet()
    page_1_table = page_1_piv(df)
    page_1_bargraph(page_1_table, sector)
    page_1_stackedbar(page_1_table, sector)
    p1_barchart_file = Image('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-bar.png')
    p1_stackedbar_file = Image('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-stackedbar.png')

    p1tab = Table(np.vstack((list(page_1_table), np.array(page_1_table))).tolist())
    q = (len(page_1_table.columns) - 1, len(page_1_table))
    p1tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
                                    ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
                                    ('FONTSIZE', (0, 1), (q[0], q[1]), 8),
                                    ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
                                    ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
                                    ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
                                    ]))




    page_2_table = page_2_piv(df)
    p2tab = Table(np.vstack((list(page_2_table), np.array(page_2_table))).tolist())
    q = (len(page_2_table.columns) - 1, len(page_2_table))
    p2tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
                               ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
                               ('FONTSIZE', (0, 1), (q[0], q[1]), 8),
                               ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
                               ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
                               ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
                               ]))

    page_3_table = page_3_piv(df)
    p3tab = Table(np.vstack((list(page_3_table), np.array(page_3_table))).tolist())
    q = (len(page_3_table.columns) - 1, len(page_3_table))
    p3tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
                               ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
                               ('FONTSIZE', (0, 1), (q[0], q[1]), 8),
                               ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
                               ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
                               ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
                               ]))
    main_pdf.append(p1tab)
    main_pdf.append(p1_barchart_file)
    main_pdf.append(p1_stackedbar_file)
    main_pdf.append(PageBreak())
    main_pdf.append(p2tab)
    main_pdf.append(PageBreak())
    main_pdf.append(p3tab)
    main_pdf.append(PageBreak())
    doc.build(main_pdf)

for i in master_df['Sector/Directorate/HSCP'].unique():
    pdfmaker(i)

    #TODO create styles

finish = pd.Timestamp.now()
print(f'Total time elapsed: {(finish - start).seconds} seconds')