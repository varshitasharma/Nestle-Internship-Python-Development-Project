from tkinter import *
import xlrd
import os
import matplotlib.pyplot as plt
import math

#creating main widget
root = Tk()
root.title("Nestle")
root.configure(background='whitesmoke')

global tables
tables = {}
global sheetnames
sheetnames = []
#button_row = 6

class rec(object):
    def __init__(self,mon, p1, p2, p3,f,t,y):
        self.M = mon
        self.P1 = p1
        self.P2 = p2
        self.P3 = p3
        self.F = f
        self.T = t
        self.Y = y


def create_button(sheetname ):
    btn = Button(root, text ="Show Graphs - " + sheetname  , bg="whitesmoke" , fg = "blueviolet", command = lambda: graph(sheetname,tables[sheetname]))
    btn.pack()
    space_label = Label(root, text="\n").pack()

def create_lists(sheet_index, workbook):
    sheetname = sheetnames[sheet_index]
    sheet = workbook.sheet_by_index(sheet_index)
    mon = []
    p1 = []
    p2 = []
    p3 = []
    f =[]
    t = []
    y = []

    for i in range(1,5):
        t.append(sheet.cell_value(1,i))
        y.append(sheet.cell_value(14,i))


    if(sheet.cell_value(2, 1) == ''):
        p1.append('')
    else:
        for i in range(2, 14):
            if(sheet.cell_value(i, 1) == ''):
                break
            p1.append(sheet.cell_value(i, 1))


    if(sheet.cell_value(2, 2) == ''):
        p2.append('')
    else:
        for i in range(2, 14):
            if(sheet.cell_value(i, 2) == ''):
                break
            p2.append(sheet.cell_value(i, 2))

    if(sheet.cell_value(2, 3) == ''):
        p3.append('')
    else:
        for i in range(2, 14):
            if(sheet.cell_value(i, 3) == ''):
                break
            p3.append(sheet.cell_value(i,3))

    if(sheet.cell_value(2, 4) == ''):
        f.append('')
    else:
        for i in range(2, 14):
            if(sheet.cell_value(i, 4) == ''):
                break
            f.append(sheet.cell_value(i, 4))
            mon.append(sheet.cell_value(i, 0))


    tables[sheetname] = rec(mon, p1, p2, p3, f, t, y)

def autolabel(rects, ax):
# Get y-axis height to calculate label position from.
    (y_bottom, y_top) = ax.get_ylim()
    y_height = y_top - y_bottom

    for rect in rects:
        height = rect.get_height()

        # Fraction of axis height taken up by this rectangle
        p_height = (height / y_height)

        # If we can fit the label above the column, do that;
        # otherwise, put it inside the column.
        if p_height > 0.95: # arbitrary; 95% looked good to me.
            label_position = height - (y_height * 0.05)
        else:
            label_position = height + (y_height * 0.01)

        ax.text(rect.get_x() + rect.get_width()/2., label_position,round(float(height),2),
                ha='center', va='bottom')


def open_excel_file():
    wb = xlrd.open_workbook(os.getcwd()+"\\Costing.xlsx")
    num_of_sheets = len(wb.sheet_names())
    sheetnames.extend(wb.sheet_names())

    for i in range(0, num_of_sheets):
        create_lists(i, wb)
        create_button(sheetnames[i])

def graph1(head,data):
    if(data.T[0] == ''):
        h1 = 'PHASE-2'
        h2 = 'PHASE-3'
        p1 = data.P2
        p2 = data.P3
        f =  data.F
        mon = data.M
        target = [data.T[1], data.T[2], data.T[3]]
        ytd = [data.Y[1], data.Y[2], data.Y[3]]
    elif(data.T[1] == ''):
        print(data.T[1],"1")
        h1 = 'PHASE-1'
        h2 = 'PHASE-3'
        p1 = data.P1
        p2 = data.P3
        f =  data.F
        mon = data.M
        target = [data.T[0], data.T[2], data.T[3]]
        ytd = [data.Y[0], data.Y[2], data.Y[3]]
    elif(data.T[2] == ''):
        target = [data.T[0], data.T[1], data.T[3]]
        ytd = [data.Y[0], data.Y[1], data.Y[3]]
        h1 = 'PHASE-1'
        h2 = 'PHASE-2'
        p1 = data.P1
        p2 = data.P2
        f =  data.F
        mon = data.M

    len1 = len(p1)
    colors1 = []
    colors2 = []
    colors3 = []
    for i in range(0,len1):
        if(head == 'Var vs 2019(INR Mio)'):
            if(p1[i] < target[0] ):
                colors1.append('red')
            elif(p1[i] >= target[0]):
                colors1.append('lawngreen')
            if(p2[i] < target[1] ):
                colors2.append('red')
            elif(p2[i] >= target[1]):
                colors2.append('lawngreen')
            if(f[i] < target[2] ):
                colors3.append('red')
            elif(f[i] >= target[2]):
                colors3.append('lawngreen')
        else:
            if(p1[i] > target[0] ):
                colors1.append('red')
            elif(p1[i] <= target[0]):
                colors1.append('lawngreen')
            if(p2[i] > target[1] ):
                colors2.append('red')
            elif(p2[i] <= target[1]):
                colors2.append('lawngreen')
            if(f[i] > target[2] ):
                colors3.append('red')
            elif(f[i] <= target[2]):
                colors3.append('lawngreen')

    u1 = math.floor( max(max(p1),target[0])) + 8
    u2 = math.floor(max(max(p2),target[1])) + 8
    u3 = math.floor(max(max(f),target[2])) + 8

    if( min(min(p1), target[0]) < 0):
        l1 = math.ceil( min(min(p1), target[0]) - 1)
    else:
        l1 = 0
    if( min(min(p2), target[1]) < 0):
        l2 = math.ceil( min(min(p2), target[1]) - 1)
    else:
        l2 = 0
    if( min(min(f), target[2]) < 0):
        l3 = math.ceil( min(min(f), target[2]) - 1)
    else:
        l3 = 0

    fig = plt.figure(figsize=(12,8))
    ax1 = fig.add_subplot(221)
    ax2 = fig.add_subplot(222)
    ax3 = fig.add_subplot(212)

    r1 = ax1.bar(mon, p1, color = colors1 )
    r2 = ax2.bar(mon, p2, color = colors2 )
    r3 = ax3.bar(mon, f, color = colors3 )

    ax1.set_ylim(l1,u1)
    ax2.set_ylim(l2,u2)
    ax3.set_ylim(l3,u3)

    ax1.set_ylabel(head )
    ax2.set_ylabel(head )
    ax3.set_ylabel(head )

    ax1.set_title("MONTHLY "+ h1 +" " +head)
    ax2.set_title("MONTHLY " + h2 +" " +head)
    ax3.set_title("MONTHLY FACTORY " + head)

    autolabel(r1, ax1)
    autolabel(r2, ax2)
    autolabel(r3, ax3)

    if head == 'Var vs 2019(INR Mio)':
        c = '>'
    else:
        c = '<'

    ax1.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
    ax2.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
    ax3.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)

    ax1.text(0,u1-2.8, "Target : " +str(round(target[0],2)) + "\n YTD : " + str(round(ytd[0],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
    ax2.text(0,u2-2.8, "Target : " +str(round(target[1],2)) + "\n YTD : " + str(round(ytd[1],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
    ax3.text(0,u3-2.8, "Target : " + str(round(target[2],2)) + "\n YTD : " + str(round(ytd[2],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})

    ax1.axhline( color='crimson', linestyle='-',linewidth = 0.9)
    ax2.axhline( color='crimson', linestyle='-',linewidth = 0.9)
    ax3.axhline( color='crimson', linestyle='-',linewidth = 0.9)

    plt.tight_layout(pad = 5)
    plt.show()

def graph2(head, data):
    if(data.T[0] == '' and data.T[1] == ''):
        h1 = 'PHASE-3'
        p1 = data.P3
        f =  data.F
        mon = data.M
        target = [data.T[2], data.T[3]]
        ytd = [data.Y[2], data.Y[3]]
    elif(data.T[1] == '' and data.T[2] == ''):
        h1 = 'PHASE-1'
        p1 = data.P1
        f =  data.F
        mon = data.M
        target = [data.T[0], data.T[3]]
        ytd = [data.Y[0], data.Y[3]]
    elif(data.T[2] == '' and data.T[0]):
        target = [data.T[1], data.T[3]]
        ytd = [data.Y[1], data.Y[3]]
        h1 = 'PHASE-2'
        p1 = data.P2
        f =  data.F
        mon = data.M

    len1 = len(p1)
    colors1 = []
    colors2 = []
    for i in range(0,len1):
        if(head == 'Var vs 2019(INR Mio)'):
            if(p1[i] < target[0] ):
                colors1.append('red')
            elif(p1[i] >= target[0]):
                colors1.append('lawngreen')
            if(f[i] < target[1] ):
                colors2.append('red')
            elif(f[i] >= target[1]):
                colors2.append('lawngreen')
        else:
            if(p1[i] > target[0] ):
                colors1.append('red')
            elif(p1[i] <= target[0]):
                colors1.append('lawngreen')
            if(f[i] > target[1] ):
                colors2.append('red')
            elif(f[i] <= target[1]):
                colors2.append('lawngreen')

    u1 = math.floor( max(max(p1),target[0])) + 8
    u2 = math.floor(max(max(f),target[1])) + 8

    if( min(min(p1), target[0]) < 0):
        l1 = math.ceil( min(min(p1), target[0]) - 1)
    else:
        l1 = 0
    if( min(min(f), target[1]) < 0):
        l2 = math.ceil( min(min(f), target[1]) - 1)
    else:
        l2 = 0

    fig = plt.figure(figsize=(12,5))
    ax1 = fig.add_subplot(221)
    ax2 = fig.add_subplot(222)

    r1 = ax1.bar(mon, p1, color = colors1 )
    r2 = ax2.bar(mon, f, color = colors2 )

    ax1.set_ylim(l1,u1)
    ax2.set_ylim(l2,u2)

    ax1.set_ylabel(head )
    ax2.set_ylabel(head )

    ax1.set_title("MONTHLY "+ h1 +" " +head)
    ax2.set_title("MONTHLY FACTORY " + head)

    autolabel(r1, ax1)
    autolabel(r2, ax2)

    if head == 'Var vs 2019(INR Mio)':
        c = '>'
    else:
        c = '<'

    ax1.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
    ax2.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)

    ax1.text(0,u1-2.8, "Target : " +str(round(target[0],2)) + "\n YTD : " + str(round(ytd[0],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
    ax2.text(0,u2-2.8, "Target : " +str(round(target[1],2)) + "\n YTD : " + str(round(ytd[1],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})

    ax1.axhline( color='crimson', linestyle='-',linewidth = 0.9)
    ax2.axhline( color='crimson', linestyle='-',linewidth = 0.9)

    plt.tight_layout(pad = 5)
    plt.show()

def graph3(head, data):
    f =  data.F
    mon = data.M
    target = data.T[3]
    ytd = data.Y[3]
    len1 = len(f)
    colors = []

    for i in range(0,len1):
        if(head == 'Var vs 2019(INR Mio)'):
            if(f[i] < target ):
                colors.append('red')
            elif(f[i] >= target):
                colors.append('lawngreen')
        else:
            if(f[i] > target ):
                colors.append('red')
            elif(f[i] <= target):
                colors.append('lawngreen')

    u = math.floor( max(max(f),target)) + 8
    if( min(min(f), target) < 0):
        l = math.ceil( min(min(f), target) - 1)
    else:
        l = 0

    figure = plt.figure()
    ax = figure.add_subplot(111)

    r = ax.bar(mon, f, color = colors  )

    ax.set_ylim(l,u)

    ax.set_ylabel(head )

    ax.set_title("MONTHLY FACTORY " + head)

    autolabel(r, ax)

    if head == 'Var vs 2019(INR Mio)':
        c = '>'
    else:
        c = '<'
    ax.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)

    ax.text(0,u-4, "Target : " +  str(round(target,2)) + "\n YTD : " + str(round(ytd,2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})

    ax.axhline( color='crimson', linestyle='-',linewidth = 0.9)

    plt.tight_layout()
    plt.show()


def graph(head,data):
    p1 = data.P1
    p2 = data.P2
    p3 = data.P3
    f =  data.F
    mon = data.M
    target = data.T
    ytd = data.Y
    if(target[0] == '' and target[1] == '' and target[2] == ''):
        graph3(head,data)
    elif((target[0] == '' and target[1] == '' ) or (target[0] =='' and target[2] == '') or (target[1] == '' and target[2] == '')):
        graph2(head,data)
    elif( target[0] == '' or target[1] == '' or target[2] == ''):
        graph1(head,data)
    else:
        len1 = len(p1)
        colors1 = []
        colors2 = []
        colors3 = []
        colors4 = []

        for i in range(0,len1):
            if(head == 'Var vs 2019(INR Mio)'):
                if(p1[i] < target[0] ):
                        colors1.append('red')
                elif(p1[i] >= target[0]):
                        colors1.append('lawngreen')
                if(p2[i] < target[1] ):
                        colors2.append('red')
                elif(p2[i] >= target[1]):
                        colors2.append('lawngreen')
                if(p3[i] < target[2] ):
                        colors3.append('red')
                elif(p3[i] >= target[2]):
                        colors3.append('lawngreen')
                if(f[i] < target[3] ):
                        colors4.append('red')
                elif(f[i] >= target[3]):
                        colors4.append('lawngreen')
            else:

                if(p1[i] > target[0] ):
                        colors1.append('red')
                elif(p1[i] <= target[0]):
                        colors1.append('lawngreen')
                if(p2[i] > target[1] ):
                        colors2.append('red')
                elif(p2[i] <= target[1]):
                        colors2.append('lawngreen')
                if(p3[i] > target[2] ):
                        colors3.append('red')
                elif(p3[i] <= target[2]):
                        colors3.append('lawngreen')
                if(f[i] > target[3] ):
                        colors4.append('red')
                elif(f[i] <= target[3]):
                        colors4.append('lawngreen')


        u1 = math.floor( max(max(p1),target[0])) + 8
        u2 = math.floor(max(max(p2),target[1])) + 8
        u3 = math.floor(max(max(p3),target[2])) + 8
        u4 = math.floor(max(max(f),target[3])) + 8

        if( min(min(p1), target[0]) < 0):
            l1 = math.ceil( min(min(p1), target[0]) - 1)
        else:
            l1 = 0
        if( min(min(p2), target[1]) < 0):
            l2 = math.ceil( min(min(p2), target[1]) - 1)
        else:
            l2 = 0
        if( min(min(p3), target[2]) < 0):
            l3 = math.ceil( min(min(p3), target[2]) - 1)
        else:
            l3 = 0
        if( min(min(f), target[3]) < 0):
            l4 = math.ceil( min(min(f), target[3]) - 1)
        else:
            l4 = 0


        fig, ((ax1,ax2),(ax3,ax4)) = plt.subplots(figsize=(12,8),nrows=2, ncols=2)
        plt.subplots_adjust(bottom = 0.2, left = 0.1)


        r1 = ax1.bar(mon, p1, color = colors1 )
        r2 = ax2.bar(mon, p2, color = colors2 )
        r3 = ax3.bar(mon, p3, color = colors3 )
        r4 = ax4.bar(mon, f, color = colors4  )


        ax1.set_ylim(l1,u1)
        ax2.set_ylim(l2,u2)
        ax3.set_ylim(l3,u3)
        ax4.set_ylim(l4,u4)

        ax1.set_ylabel(head )
        ax2.set_ylabel(head )
        ax3.set_ylabel(head )
        ax4.set_ylabel(head )

        ax1.set_title("MONTHLY PHASE-1 " + head)
        ax2.set_title("MONTHLY PHASE-2 " + head)
        ax3.set_title("MONTHLY PHASE-3 " + head)
        ax4.set_title("MONTHLY FACTORY " + head)


        autolabel(r1, ax1)
        autolabel(r2, ax2)
        autolabel(r3, ax3)
        autolabel(r4, ax4)

        if head == 'Var vs 2019(INR Mio)':
            c = '>'
        else:
            c = '<'


        ax1.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
        ax2.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
        ax3.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)
        ax4.grid( color='dodgerblue', linestyle='--',linewidth = 0.9)

        ax1.text(0,u1-2.8, "Target : " +str(round(target[0],2)) + "\n YTD : " + str(round(ytd[0],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
        ax2.text(0,u2-2.8, "Target : " +str(round(target[1],2)) + "\n YTD : " + str(round(ytd[1],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
        ax3.text(0,u3-2.8, "Target : " + str(round(target[2],2)) + "\n YTD : " + str(round(ytd[2],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})
        ax4.text(0,u4-4, "Target : " +  str(round(target[3],2)) + "\n YTD : " + str(round(ytd[3],2)), bbox={'facecolor': 'cyan', 'alpha': 0.3, 'pad': 3})

        ax1.axhline( color='crimson', linestyle='-',linewidth = 0.9)
        ax2.axhline( color='crimson', linestyle='-',linewidth = 0.9)
        ax3.axhline( color='crimson', linestyle='-',linewidth = 0.9)
        ax4.axhline( color='crimson', linestyle='-',linewidth = 0.9)



        plt.tight_layout(pad = 5)
        plt.show()

heading =Label(root,text = "Costing Operation Review",pady=5, bg="whitesmoke" , fg = "mediumorchid", font=("Helvetica", 18))
heading.pack()

space_label = Label(root, text="\n").pack()

open_excel_file()

root.mainloop()
