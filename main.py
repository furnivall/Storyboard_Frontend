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
from PIL import Image as img2

start = pd.Timestamp.now()
# Read in file

Tk().withdraw() #hide tk window
master_df_filename = askopenfilename(initialdir='/media/wdrive/storyboards/',
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
    piv = pd.pivot_table(df, index='Report Date', values=['Absence WTE Short', 'Absence WTE Long',
                                                          'Absence WTE', 'Annual WTE', 'Maternity WTE',
                                                          'Paternal WTE', 'Parental WTE',
                                                          'Public Holiday WTE', 'Study WTE', 'Special WTE',
                                                          'Other WTE', 'WTE'
                                                          ], aggfunc = np.sum).round(2)
    for column in ['Absence WTE Short', 'Absence WTE Long','Absence WTE', 'Annual WTE', 'Maternity WTE',
              'Paternal WTE', 'Parental WTE','Public Holiday WTE', 'Study WTE', 'Special WTE', 'Other WTE', 'WTE']:
        piv[column] = (piv[column] / piv['WTE'] * 100).round(1)
    piv = piv.rename(columns={
        'Absence WTE Short':'SickAbs Short %', 'Absence WTE Long':'SickAbs Long %',
        'Absence WTE':'Sickness Abs %', 'Annual WTE':'Annual Leave %', 'Maternity WTE':'Maternity %',
        'Paternal WTE':'Paternity %', 'Parental WTE': 'Parental %',
        'Public Holiday WTE':'Public Holiday %', 'Study WTE':'Study %', 'Special WTE':'Special %',
        'Other WTE':'Other %'
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
    plt.savefig('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'-page1-bar.png', dpi=300)

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
    plt.savefig('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'-page1-stackedbar.png', dpi=300)

def page3_sickabsGraph(piv, sector):

    WTE_Long = piv['SickAbs Long %'].tolist()
    WTE_Short = piv['SickAbs Short %'].tolist()
    months = piv['Report Date'].tolist()
    for i in [WTE_Long, WTE_Short, months]:
        i.reverse()
    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, WTE_Short, width, label='Absence WTE Long %')
    ax.bar(months, WTE_Long, width, label='Absence WTE Short %', bottom = WTE_Short)
    ax.set_title(sector + ' Sick Absence')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    plt.hlines(4, -1, len(months), colors='C0', linestyles='dashed', label='Target (4%)')
    ax.legend()
    plt.savefig('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'-page3-sickabsgraph.png', dpi=300)

def page_3_otherleavegraph(piv, sector):
    special = piv['Special %'].tolist()
    maternity = piv['Maternity %'].tolist()
    study = piv['Study %'].tolist()
    parental = piv['Parental %'].tolist()
    public = piv['Public Holiday %'].tolist()
    other = piv['Other %'].tolist()
    paternal = piv['Paternity %'].tolist()
    months = piv['Report Date'].tolist()
    for i in [special, maternity, study, parental, public, other, paternal, months]:
        i.reverse()
    bars1 = np.add(special, maternity).tolist()
    bars2 = np.add(bars1, study).tolist()
    bars3 = np.add(bars2, parental).tolist()
    bars4 = np.add(bars3, public).tolist()
    bars5 = np.add(bars4, other).tolist()

    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, special, width, label='Special WTE %')
    ax.bar(months, maternity, width, label='Maternity WTE %', bottom=special)
    ax.bar(months, study, width, label='Study WTE %', bottom=bars1)
    ax.bar(months, parental, width, label='Parental WTE %', bottom=bars2)
    ax.bar(months, public, width, label='Public Holiday WTE %', bottom=bars3)
    ax.bar(months, other, width, label='Other Leave WTE %', bottom=bars4)
    ax.bar(months, paternal, width, label='Paternal WTE %', bottom=bars5)
    ax.set_title(sector + ' Other Leave')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    ax.legend()
    plt.savefig('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page3-otherleave.png', dpi=300)

def page3_annualleavegraph(piv, sector):
    annual = piv['Annual Leave %'].tolist()
    months = piv['Report Date'].tolist()
    annual.reverse()
    months.reverse()
    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, annual, width, label = 'Annual Leave WTE %')
    ax.set_title(sector + ' Annual Leave')
    ax.set_xticklabels(months, rotation=45, fontsize=8)

    plt.hlines(10.5, -1, len(months), colors='C0', linestyles='dashed', label='Target (10.5%)')
    ax.legend()
    plt.savefig('/media/wdrive/Storyboards/Test_Boards/graphs/'+ sector + '-page3-annualleave.png', dpi=300)






df = df_cutter(master_df, 'eHealth')
print(page_1_piv(df))
print(page_2_piv(df))
print(page_3_piv(df))



#TODO page 1 tables



#TODO page 1 graphs


def pdfmaker(sector):
    df = df_cutter(master_df, sector)
    doc = SimpleDocTemplate("/media/wdrive/Storyboards/Test_Boards/"+sector+'.pdf', rightMargin=10, leftMargin=10,
                            topMargin=10, bottomMargin=10, pagesize=(A4[1], A4[0]))
    main_pdf = []
    styles = getSampleStyleSheet()

    ###    PAGE 1 DATA    ###
    page_1_table = page_1_piv(df)

    # page 1 table
    p1tab = Table(np.vstack((list(page_1_table), np.array(page_1_table))).tolist(), colWidths=2 * cm)
    q = (len(page_1_table.columns) - 1, len(page_1_table))
    p1tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
                                    ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
                                    #('FONTSIZE', (0, 1), (q[0], q[1]), 8),
                                    ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
                                    ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
                                    ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
                                    ]))
    w, h = p1tab.wrap(0, 0)

    page_1_bargraph(page_1_table, sector)
    page_1_stackedbar(page_1_table, sector)
    image1 = img2.open('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page1-bar.png')
    image2 = img2.open('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page1-stackedbar.png')

    page_1_graphs_combined = img2.new('RGB',(image1.width, image1.height + image2.height))
    page_1_graphs_combined.paste(image1, (0,0))
    page_1_graphs_combined.paste(image2, (0, image1.height))

    page_1_graphs_combined = page_1_graphs_combined.resize((300, 500), img2.ANTIALIAS)
    page_1_graphs_combined.save('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + 'page1_graphs.png')
    p1graphs = Image('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'page1_graphs.png')

    p1_barchart_file = Image('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page1-bar.png', 0.8 * w, h * 0.8)
    p1_stackedbar_file = Image('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page1-stackedbar.png', w *0.8, h *0.8)


    p1_wrapper = Table([[p1graphs, p1tab]])
    print(p1_wrapper.wrap(0,0))
    #p1_wrapper2 = Table([[p1_stackedbar_file]])


    ###    PAGE 2 DATA    ###
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

    ###    PAGE 3 DATA    ###
    page_3_table = page_3_piv(df)
    print(page_3_table.columns)
    page3_sickabsGraph(page_3_table, sector)
    page_3_otherleavegraph(page_3_table, sector)
    page3_annualleavegraph(page_3_table, sector)
    p3_othergraph = Image('/media/wdrive/Storyboards/Test_Boards/graphs/' + sector + '-page3-otherleave.png')
    p3_sickabsgraph = Image('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'-page3-sickabsgraph.png')
    p3_annualgraph = Image('/media/wdrive/Storyboards/Test_Boards/graphs/'+sector+'-page3-annualleave.png')

    p3tab = Table(np.vstack((list(page_3_table), np.array(page_3_table))).tolist())
    q = (len(page_3_table.columns) - 1, len(page_3_table))
    p3tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
                               ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
                               ('FONTSIZE', (0, 1), (q[0], q[1]), 8),
                               ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
                               ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
                               ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
                               ]))

    ### PAGE 1 REPORT ###
    #main_pdf.append(p1tab)
    main_pdf.append(p1_wrapper)
    #main_pdf.append(p1_wrapper2)
    main_pdf.append(PageBreak())

    ### PAGE 2 REPORT ###
    main_pdf.append(p2tab)
    main_pdf.append(PageBreak())

    ### PAGE 3 REPORT ###
    # main_pdf.append(p3_sickabsgraph)
    # main_pdf.append(p3_othergraph)
    # main_pdf.append(p3_annualgraph)
    # main_pdf.append(p3tab)
    # main_pdf.append(PageBreak())
    doc.build(main_pdf)

for i in master_df['Sector/Directorate/HSCP'].unique():
    pdfmaker(i)

    #TODO create styles

finish = pd.Timestamp.now()
print(f'Total time elapsed: {(finish - start).seconds} seconds')