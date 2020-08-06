import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter import Tk
import numpy as np
import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors, utils
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus.flowables import HRFlowable
import matplotlib.pyplot as plt
import winsound



from PIL import Image as img2



start = pd.Timestamp.now()
# Read in file

Tk().withdraw() #hide tk window
master_df_filename = askopenfilename(initialdir='W://Workforce Monthly Reports/monthly_reports/',
                               filetypes=(("Excel File", "*.xlsx"), ("All Files", "*.*")),
                               title="Choose the relevant master file."
                               )
print(f'Reading file - {master_df_filename}')
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
                                                          'Additional WTE'], aggfunc=np.sum).round(1)

    piv.reset_index(inplace=True)
    piv.sort_values('Report Date', ascending=False, inplace=True)
    piv['Report Date'] = piv['Report Date'].dt.strftime('%b-%y')
    return piv

def page_2_piv(df):
    piv = pd.pivot_table(df, index='Report Date', values=['WTE', 'Starters WTE', 'Leavers WTE'], aggfunc=np.sum).round(1)

    piv['Starters %'] = (piv['Starters WTE'] / piv['WTE'] * 100).round(1)
    piv['Leavers %'] = (piv['Leavers WTE'] / piv['WTE'] * 100).round(1)
    piv.reset_index(inplace=True)
    piv.sort_values('Report Date', ascending=False, inplace=True)
    piv['Report Date'] = piv['Report Date'].dt.strftime('%b-%y')
    return piv

def page_3_piv(df):
    piv = pd.pivot_table(df, index='Report Date', values=['Absence WTE Short', 'Absence WTE Long',
                                                          'Absence WTE', 'Annual WTE', 'Maternity WTE',
                                                          'Paternal WTE', 'Parental WTE',
                                                          'Public Holiday WTE', 'Study WTE', 'Special WTE',
                                                          'Other WTE', 'WTE'
                                                          ], aggfunc = np.sum).round(2)
    for column in ['Absence WTE Short', 'Absence WTE Long','Absence WTE', 'Annual WTE', 'Maternity WTE',
              'Paternal WTE', 'Parental WTE','Public Holiday WTE', 'Study WTE', 'Special WTE', 'Other WTE', 'WTE']:
        piv[column] = (piv[column] / piv['WTE'] * 100).round(1)


    piv.reset_index(inplace=True)
    piv = piv.rename(columns={
        'Absence WTE Short': 'Short %', 'Absence WTE Long': 'Long %',
        'Absence WTE': 'SickAbs %', 'Annual WTE': 'Annual %', 'Maternity WTE': 'Maternity %',
        'Paternal WTE': 'Paternity %', 'Parental WTE': 'Parental %',
        'Public Holiday WTE': 'Pub Hol %', 'Study WTE': 'Study %', 'Special WTE': 'Special %',
        'Other WTE': 'Other %', 'Report Date': 'Month'
    })

    piv.sort_values('Month', ascending=False, inplace=True)
    piv['Month'] = piv['Month'].dt.strftime('%b-%y')
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
    plt.style.use('seaborn-pastel')
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
    # autolabel(rects1, hc_avg, 'C0', 'Headcount Average')
    # autolabel(rects2, WTE_avg, 'C1', 'WTE Average')
    fig.tight_layout()
    plt.savefig('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-bar.png', dpi=300)

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

    plt.style.use('seaborn-pastel')
    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, bank, width, label='Bank WTE')
    ax.bar(months, ot, width, label='Overtime WTE', bottom=bank)
    ax.bar(months, excess, width, label='Excess WTE', bottom=bars1)
    ax.bar(months, additional, width, label='Additional WTE', bottom=bars2)
    ax.set_title(sector + ' Supplementary WTE')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    ax.legend()
    plt.savefig('W:/Storyboards/Test_Boards/graphs/'+sector+'-page1-stackedbar.png', dpi=300)

def page_2_twoaxes(piv, sector):
    wte = piv['WTE'].tolist()
    starters = piv['Starters WTE'].tolist()
    leavers = piv['Leavers WTE'].tolist()
    months = piv['Report Date'].tolist()
    for i in [wte, starters, leavers, months]:
        i.reverse()
    plt.style.use('seaborn-pastel')
    width = 0.35
    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()
    x = np.arange(len(months))
    ax1.bar(x + width / 3, wte, width / 3, label='WTE', color='C2')
    ax2.bar(x - width / 3, starters, width/3, label='Starters WTE', color = 'C0')
    ax2.bar(x, leavers, width/3, label='Leavers WTE', color= 'C1')


    ax1.set_title(sector + ' Starters & Leavers')
    ax1.set_xticks(x)
    ax1.set_xticklabels(months, rotation=45, fontsize=8)

    ax1.set_ylabel('WTE')
    ax2.set_ylabel('Starter/Leaver WTE')

    lines, labels = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines + lines2, labels + labels2)
    plt.savefig('W://Storyboards/Test_Boards/graphs/'+sector+'-page2-twoaxes.png', dpi=300)



def page3_sickabsGraph(piv, sector):

    WTE_Long = piv['Long %'].tolist()
    WTE_Short = piv['Short %'].tolist()
    months = piv['Month'].tolist()
    for i in [WTE_Long, WTE_Short, months]:
        i.reverse()
    width = 0.35
    plt.style.use('seaborn-pastel')
    fig, ax = plt.subplots()
    ax.bar(months, WTE_Short, width, label='Absence WTE Long %')
    ax.bar(months, WTE_Long, width, label='Absence WTE Short %', bottom = WTE_Short)
    ax.set_title(sector + ' Sick Absence')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    plt.hlines(4, -1, len(months), colors='C0', linestyles='dashed', label='Target (4%)')
    ax.legend()
    plt.savefig('W://Storyboards/Test_Boards/graphs/'+sector+'-page3-sickabsgraph.png', dpi=300)

def page_3_otherleavegraph(piv, sector):
    special = piv['Special %'].tolist()
    maternity = piv['Maternity %'].tolist()
    study = piv['Study %'].tolist()
    parental = piv['Parental %'].tolist()
    public = piv['Pub Hol %'].tolist()
    other = piv['Other %'].tolist()
    paternal = piv['Paternity %'].tolist()
    months = piv['Month'].tolist()
    for i in [special, maternity, study, parental, public, other, paternal, months]:
        i.reverse()
    bars1 = np.add(special, maternity).tolist()
    bars2 = np.add(bars1, study).tolist()
    bars3 = np.add(bars2, parental).tolist()
    bars4 = np.add(bars3, public).tolist()
    bars5 = np.add(bars4, other).tolist()
    plt.style.use('seaborn-pastel')
    width = 0.35
    fig, ax = plt.subplots()
    ax.bar(months, special, width, label='Special%')
    ax.bar(months, maternity, width, label='Maternity %', bottom=special)
    ax.bar(months, study, width, label='Study %', bottom=bars1)
    ax.bar(months, parental, width, label='Parental %', bottom=bars2)
    ax.bar(months, public, width, label='Public Hol %', bottom=bars3)
    ax.bar(months, other, width, label='Other %', bottom=bars4)
    ax.bar(months, paternal, width, label='Paternal %', bottom=bars5)
    ax.set_title(sector + ' Other Leave')
    ax.set_xticklabels(months, rotation=45, fontsize=8)
    ax.legend()
    plt.savefig('W://Storyboards/Test_Boards/graphs/' + sector + '-page3-otherleave.png', dpi=300)

def page3_annualleavegraph(piv, sector):
    annual = piv['Annual %'].tolist()
    months = piv['Month'].tolist()
    annual.reverse()
    months.reverse()
    width = 0.35
    plt.style.use('seaborn-pastel')
    fig, ax = plt.subplots()
    ax.bar(months, annual, width, label = 'Annual Leave WTE %')
    ax.set_title(sector + ' Annual Leave')
    ax.set_xticklabels(months, rotation=45, fontsize=8)

    plt.hlines(10.5, -1, len(months), colors='C0', linestyles='dashed', label='Target (10.5%)')
    ax.legend()
    plt.savefig('W://Storyboards/Test_Boards/graphs/'+ sector + '-page3-annualleave.png', dpi=300)

def abs_analysis_pivs(df):
    df_annual = df[df['Report Date'].isin(['2020-06-01', '2020-05-01'])]
    df = df[df['Report Date'] == '2020-06-01']

    # Short Term pivot
    df['Short Term Sick %'] = df['Absence WTE Short'] / df['WTE'] * 100
    short_term_mean = df['Short Term Sick %'].mean()
    piv_shortterm = pd.pivot_table(df, index='Sub-Directorate 2',
                                   values=['Short Term Sick %', 'WTE'],
                                   aggfunc={'Short Term Sick %':np.mean, 'WTE':np.sum}).round(2)
    piv_shortterm = piv_shortterm[(piv_shortterm['WTE'] > 15) & (piv_shortterm['Short Term Sick %'] > short_term_mean)]
    piv_shortterm.reset_index(inplace=True)
    # Long Term pivot
    df['Long Term Sick %'] = df['Absence WTE Long'] / df['WTE'] * 100
    long_term_mean = df['Long Term Sick %'].mean()
    piv_longterm = pd.pivot_table(df, index='Sub-Directorate 2',
                                  values=['Long Term Sick %', 'WTE'],
                                  aggfunc={'Long Term Sick %':np.mean, 'WTE':np.sum}).round(2)
    piv_longterm = piv_longterm[(piv_longterm['WTE'] > 15) & (piv_longterm['Long Term Sick %'] > long_term_mean)]
    piv_longterm.reset_index(inplace=True)
    # TODO 25% winter pivot

    # Annual Leave Pivot
    #df_annual['Report Date'] = df_annual['Report Date'].dt.strftime('%b-%y')
    df_annual['Annual Leave %'] = df_annual['Annual WTE'] / df_annual['WTE'] * 100
    piv_annual = pd.pivot_table(df_annual,
                                index ='Sub-Directorate 1',
     #                           columns = 'Report Date',
                                values=['Annual Leave %', 'WTE'],
                                aggfunc={'Annual Leave %': np.mean, 'WTE': np.sum}).round(2)

    piv_annual = piv_annual[(piv_annual['WTE'] > 50) & (piv_annual['Annual Leave %'] < 9)]
    piv_annual.reset_index(inplace=True)
    print(piv_annual)

    return piv_shortterm, piv_longterm, piv_annual





#TODO page 1 tables



#TODO page 1 graphs

def verticalimage_combine(image1, image2):
    image1 = img2.open(image1)
    image2 = img2.open(image2)
    final = img2.new('RGB', (image1.width, image1.height + image2.height))
    final.paste(image1, (0, 0))
    final.paste(image2, (0, image1.height))
    return final

def get_image(path, width=1*cm):
    img = utils.ImageReader(path)
    iw, ih = img.getSize()
    aspect = ih / float(iw)
    return Image(path, width=width, height=(width * aspect))

def excel_maker(sector, df):

    date = pd.Timestamp.now().strftime('%b-%y')
    shortterm, longterm, annual = abs_analysis_pivs(df)
    starterleavers = page_2_piv(df)
    absWTE = page_3_piv(df)
    bankOT = page_1_piv(df)
    with pd.ExcelWriter('W:/storyboards/excel/'+sector+'-'+date+'.xlsx') as writer:
        shortterm.to_excel(writer, sheet_name="ShortTerm>15WTE", index=False)
        longterm.to_excel(writer, sheet_name='LongTerm>15WTE', index=False)
        annual.to_excel(writer, sheet_name='AnnualLeave>50WTE', index=False)
        absWTE.to_excel(writer, sheet_name='AbsenceTypes', index=False)
        starterleavers.to_excel(writer, sheet_name='Starters&Leavers', index=False)
        bankOT.to_excel(writer, sheet_name='Bank-OT-Excess', index=False)
    writer.save()



# def pdfmaker(sector):
    # styles = getSampleStyleSheet()
    # styles.add(ParagraphStyle(name="subtitle", alignment=TA_CENTER, fontSize=14, textColor=colors.HexColor("#003087")))
    # styles.add(ParagraphStyle(name="heading-blue", alignment= TA_LEFT, fontSize=32, textColor =colors.HexColor('#003087')))
    # print(f'Building {sector} PDF.')
    # df = df_cutter(master_df, sector)
    # doc = SimpleDocTemplate("W://Storyboards/Test_Boards/"+sector+'.pdf', rightMargin=10, leftMargin=10,
    #                         topMargin=10, bottomMargin=10, pagesize=(A4[1], A4[0]))
    # main_pdf = []
    #
    # blue_line = HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.HexColor('#003087'),
    #                        spaceBefore=1, spaceAfter=1, hAlign='CENTER', vAlign='BOTTOM', dash=None)
    #
    # ### COVER PAGE DATA   ###
    #
    # logo = Image('W:/Storyboards/logo.jpg', 150, 122, hAlign='LEFT') #this image is 300x245px
    # words1 = Paragraph(f"Workforce Analytics Storyboard - June 2020", styles["heading-blue"])
    # words2 = Paragraph(f'{sector}', styles['heading-blue'])
    #
    # ###    PAGE 1 DATA    ###
    # page_1_table = page_1_piv(df)
    #
    # # page 1 table
    # p1tab = Table(np.vstack((list(page_1_table), np.array(page_1_table))).tolist(), colWidths=2 * cm)
    # q = (len(page_1_table.columns) - 1, len(page_1_table))
    # p1tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                                 ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                                 ('FONTSIZE', (0, 0), (q[0], q[1]), 8),
    #                                 ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                                 ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                                 ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                                 ]))
    # w, h = p1tab.wrap(0, 0)
    #
    # page_1_bargraph(page_1_table, sector)
    # page_1_stackedbar(page_1_table, sector)
    # p1i1 = 'W://Storyboards/Test_Boards/graphs/' + sector + '-page1-bar.png'
    # p1i2 = 'W://Storyboards/Test_Boards/graphs/' + sector + '-page1-stackedbar.png'
    #
    #
    # page_1_graphs_combined = verticalimage_combine(p1i1, p1i2)
    #
    # page_1_graphs_combined = page_1_graphs_combined.resize((300, 500), img2.ANTIALIAS)
    # page_1_graphs_combined.save('W://Storyboards/Test_Boards/graphs/' + sector + 'page1_graphs.png', quality=95)
    # p1graphs = Image('W://Storyboards/Test_Boards/graphs/'+sector+'page1_graphs.png')
    #
    #
    # p1_wrapper = Table([[p1graphs, p1tab]])
    # print(p1_wrapper.wrap(0,0))
    # #p1_wrapper2 = Table([[p1_stackedbar_file]])
    #
    #
    # ###    PAGE 2 DATA    ###
    # page_2_table = page_2_piv(df)
    # p2tab = Table(np.vstack((list(page_2_table), np.array(page_2_table))).tolist(), colWidths=2 * cm)
    # q = (len(page_2_table.columns) - 1, len(page_2_table))
    # p2tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                            ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                            ('FONTSIZE', (0, 0), (q[0], q[1]), 8),
    #                            ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                            ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                            ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                            ]))
    # page_2_twoaxes(page_2_table, sector)
    # p2graph = get_image('W://Storyboards/Test_Boards/graphs/'+sector+'-page2-twoaxes.png', width = 4.5 * inch)
    # p2wrapper = Table([[p2graph, p2tab]])
    #
    # ###    PAGE 3 DATA    ###
    # page_3_table = page_3_piv(df)
    # print(page_3_table.columns)
    # page3_sickabsGraph(page_3_table, sector)
    # page_3_otherleavegraph(page_3_table, sector)
    # page3_annualleavegraph(page_3_table, sector)
    # p3i1 = 'W://Storyboards/Test_Boards/graphs/' + sector + '-page3-otherleave.png'
    # p3i2 = 'W://Storyboards/Test_Boards/graphs/'+sector+'-page3-sickabsgraph.png'
    # p3i3 = 'W://Storyboards/Test_Boards/graphs/' + sector + '-page3-annualleave.png'
    # p3_graphs = verticalimage_combine(p3i1, p3i2)
    # p3_graphs.save('W://Storyboards/Test_Boards/graphs/' + sector + 'page3_graphs.png')
    # p3_graphs = verticalimage_combine('W://Storyboards/Test_Boards/graphs/' + sector + 'page3_graphs.png', p3i3)
    # p3_graphs = p3_graphs.resize((300, 500), img2.BILINEAR)
    # p3_graphs.save('W://Storyboards/Test_Boards/graphs/' + sector + 'page3_graphs.png', quality=95)
    # p3_graphs = Image('W://Storyboards/Test_Boards/graphs/' + sector + 'page3_graphs.png')
    #
    #
    #
    # p3tab = Table(np.vstack((list(page_3_table), np.array(page_3_table))).tolist(), colWidths=1.3 * cm)
    # q = (len(page_3_table.columns) - 1, len(page_3_table))
    # p3tab.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                            ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                            ('FONTSIZE', (0, 0), (q[0], q[1]), 7),
    #                            ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                            ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                            ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                            ]))
    # p3_wrapper = Table([[p3_graphs, p3tab]])
    #
    # ### Absence Analysis Data ###
    # abs_shortterm, abs_longterm, abs_annual = abs_analysis_pivs(df)
    #
    # q = (len(abs_shortterm.columns) - 1, len(abs_shortterm))
    # st_table = Table(np.vstack((list(abs_shortterm), np.array(abs_shortterm))).tolist(), hAlign='RIGHT')
    #
    # st_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                        ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                        ('FONTSIZE', (0, 0), (q[0], q[1]), 7),
    #                        ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                        ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                        ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                        ]))
    #
    # q = (len(abs_longterm.columns) - 1, len(abs_longterm))
    # lt_table = Table(np.vstack((list(abs_longterm), np.array(abs_longterm))).tolist(), hAlign='RIGHT')
    #
    # lt_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                               ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                               ('FONTSIZE', (0, 0), (q[0], q[1]), 7),
    #                               ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                               ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                               ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                               ]))
    #
    # q = (len(abs_annual.columns) - 1, len(abs_annual))
    # al_table = Table(np.vstack((list(abs_annual), np.array(abs_annual))).tolist(), hAlign='RIGHT')
    #
    # al_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (q[0], 0), colors.HexColor("#003087")),
    #                               ('TEXTCOLOR', (0, 0), (q[0], 0), colors.HexColor("#E8EDEE")),
    #                               ('FONTSIZE', (0, 0), (q[0], q[1]), 7),
    #                               ('ALIGN', (1, 0), (q[0], q[1]), 'CENTER'),
    #                               ('BOX', (0, 1), (q[0], q[1]), 0.006 * inch, colors.HexColor("#003087")),
    #                               ('BOX', (0, 0), (q[0], 0), 0.006 * inch, colors.HexColor("#003087"))
    #                               ]))
    #
    #
    #
    # ### COVER PAGE ###
    # main_pdf.append(logo)
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(words1)
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(words2)
    # main_pdf.append(PageBreak())
    #
    # ### SUMMARY PAGE ###
    # main_pdf.append(Paragraph("Sub-Directorate Summary - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0, 36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### MEASURABLE OUTCOMES PAGE ###
    # main_pdf.append(Paragraph("Measurable Outcomes - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    #
    # ### PAGE 1 REPORT ###
    # #main_pdf.append(p1tab)
    # main_pdf.append(Spacer(0,48))
    # main_pdf.append(blue_line)
    # main_pdf.append(p1_wrapper)
    # #main_pdf.append(p1_wrapper2)
    # main_pdf.append(PageBreak())
    #
    # ### PAGE 2 REPORT ###
    # main_pdf.append(Spacer(0,48))
    # main_pdf.append(blue_line)
    # main_pdf.append(p2wrapper)
    # main_pdf.append(PageBreak())
    #
    # ### PAGE 3 REPORT ###
    # main_pdf.append(Spacer(0, 48))
    # main_pdf.append(blue_line)
    # main_pdf.append(p3_wrapper)
    #
    #
    #
    # ### ABS ANALYSIS P1 ###
    # main_pdf.append(Paragraph("Absence Analysis - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(Spacer(0, 24))
    # main_pdf.append(st_table)
    # main_pdf.append(Spacer(0,24))
    # main_pdf.append(lt_table)
    #
    # main_pdf.append(PageBreak())
    #
    #
    # ### ABS ANALYSIS P2 ###
    # main_pdf.append(Paragraph("Absence Analysis - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(al_table)
    # main_pdf.append(PageBreak())
    #
    #
    # ### ABS REASONS ###
    # main_pdf.append(Paragraph("Absence Reasons - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### STAT MAND ###
    # main_pdf.append(Paragraph("Statutory/Mandatory Compliance - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HEALTH & SAFETY TRAINING 1 ###
    # main_pdf.append(Paragraph("Health & Safety - M&H - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HEALTH & SAFETY TRAINING 2 ###
    # main_pdf.append(Paragraph("Health & Safety - Falls & Sharps "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### KSF ###
    # main_pdf.append(Paragraph("KSF - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HR POLICY ACTIVITY - OUTCOME OF EMPLOYEE RELATIONS ###
    # main_pdf.append(Paragraph("HR - Employee Relations Cases - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HR POLICY ACTIVITY - OUTCOME OF ABSENCE MANAGEMENT CASES ###
    # main_pdf.append(Paragraph("HR - Absence Management Cases"+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HR POLICY ACTIVITY - OUTCOME OF ABSENCE MANAGEMENT CASES (not work related) ###
    # main_pdf.append(Paragraph("Absence - non-work related"+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### HR POLICY ACTIVITY - OUTCOME OF ABSENCE MANAGEMENT CASES (work related) ###
    # main_pdf.append(Paragraph("Absence - Work related"+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### EMPLOYEE BASE DEMOGRAPHIC - AGE RANGE, BY PAY BAND ###
    # main_pdf.append(Paragraph("Demographics 1 - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### EMPLOYEE BASE DEMOGRAPHIC - AGE RANGE, BY PAY BAND ###
    # main_pdf.append(Paragraph("Demographics 2 - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # ### EQUALITY DATASET ###
    # main_pdf.append(Paragraph("Equal Opportunities - "+sector, styles['heading-blue']))
    # main_pdf.append(Spacer(0,36))
    # main_pdf.append(blue_line)
    # main_pdf.append(PageBreak())
    #
    # doc.build(main_pdf)


excel_maker('NHS GGC', master_df)

excel_maker('Acute', master_df[master_df['Area'] == 'Acute'])


for i in master_df['Sector/Directorate/HSCP'].unique():
    # pdfmaker(i)
    df = df_cutter(master_df, i)
    excel_maker(i, df)
    #TODO create styles



finish = pd.Timestamp.now()
print(f'Total time elapsed: {(finish - start).seconds} seconds')
winsound.Beep(440, 250)