
######################################################
## This script doesn't affect any CC/CQ artifacts.
## It only delivers data to your fingertips very fast.
## It allows you to analyze either whole freeze list,
## one or more activities or make delivery preview.
## This script will be used mainly for determination
## of dependencies on different activities.
##
##
## Created by Ondrej MICHAL (H157043)
## Updated: 10/2017
######################################################

import os
import re
import Tkinter
import tkFileDialog
from subprocess import Popen, PIPE
import openpyxl

# pylint: disable=C0103, C0301, C0326
def FindView(inFileTxt):
    view = tkFileDialog.askdirectory(title="Choose view", initialdir="C:")
    inFileTxt.delete(0, 'end')
    inFileTxt.insert(0, view)

def FindFreezeFile(outFileTxt):
    freezeList = tkFileDialog.askopenfilename(title="Choose freeze list", initialdir="//az18st2701/FCOPS/C919_NLR/C919_Software_NLR/SIR", defaultextension=".xlsx")
    outFileTxt.delete(0, 'end')
    outFileTxt.insert(0, freezeList)

def Analyze(freezeList, view, VOB, intView, project):
    w = openpyxl.load_workbook(freezeList.get())
    ws = w.get_sheet_by_name(name='SIR List')
    ARdict = {}
    HonFei_Facri = [] # list of all HonFei/FACRI ARs
    PVOB = VOB.get()

    for row in ws.iter_rows():
        if row[0].value is None or row[3].value is None or row[5].value is  None:
            continue
        if row[0].value == "Building" and ('Brno' not in row[3].value) and ('India' not in row[3].value) and ('C919' in row[5].value):
            AR = re.search(r'C919[A-Z]\d{8}', row[5].value)
            if AR is None:
                continue
            AR = AR.group(0)
            com = Popen(r"cleartool descr -fmt '%[stream]Xp' activity:"+AR+"@\\"+PVOB, stdout=PIPE, universal_newlines=True, creationflags=0x08000000)
            ARvob, stderrdata = com.communicate()
            ARvob = ARvob.strip("'")
            if ARvob == "":
                # DEV_HonFei/DEV_FACRI
                HonFei_Facri.append(AR)
            elif ARvob in ARdict:
                ARdict[ARvob].append(AR)
            else:
                ARdict[ARvob] = [AR]

    # DEV_HonFei/DEV_FACRI #
    if HonFei_Facri != []:
        # cleartool deliver -preview -long -stream DEV_HonFei@\C919_FC_RC_HI_PVOB -to H157043_C919_FC_SW_Int_Voltron -target C919_FC_SW_Int@\C919_FC_RC_HI_PVOB
        com = Popen(r'cleartool lsactivity -fmt "%[crm_record_id]p %[headline]p\n" -in stream:DEV_HonFei@\C919_FC_RC_HI_PVOB')
        AllHonFeiAct, stderrdata = com.communicate()
        # cleartool deliver -preview -long -stream DEV_FACRI@\C919_FC_RC_HI_PVOB -to H157043_C919_FC_SW_Int_Voltron -target C919_FC_SW_Int@\C919_FC_RC_HI_PVOB
        com = Popen(r'cleartool lsactivity -fmt "%[crm_record_id]p %[headline]p\n" -in stream:DEV_FACRI@\C919_FC_RC_HI_PVOB')
        AllFacriAct, stderrdata = com.communicate()
        HonFeiList = []
        FACRIList = []
        for hf_ar in HonFei_Facri:
            for hf in AllHonFeiAct:
                if hf_ar in hf:
                    HonFeiList.append(hf[:12])
            for fa in AllFacriAct:
                if hf_ar in fa:
                    FACRIList.append(fa[:12])
        if HonFeiList != []:
            ARdict[r"stream:DEV_HonFei@\C919_FC_RC_HI_PVOB"] = []
            for hf in HonFeiList:
                ARdict[r"stream:DEV_HonFei@\C919_FC_RC_HI_PVOB"].append(hf)
        if FACRIList != []:
            ARdict[r"stream:DEV_FACRI@\C919_FC_RC_HI_PVOB"] = []
            for fa in FACRIList:
                ARdict[r"stream:DEV_FACRI@\C919_FC_RC_HI_PVOB"].append(fa)
    # DEV_HonFei/DEV_FACRI #

    for d in ARdict:
        print ARdict[d]
    # go through all ARs and check, if latest versions have hyperlinks with r"/C919_FC_SW_Int/" (i.e are delivered)
    exit()
    # to be able to execute cleartool command and don't have to do it more then once
    os.chdir(r"M:\%s" % (view.get()))
    OutputData = []

    for i in ARdict:
        if i == "DEV_HonFei" or i == "Dev_FACRI":
            continue
        SecondOrMore = False
        CtDeliverStrm = r"cleartool deliver -preview -stream %s@\%s -to %s -target %s@\%s -activities " % (i, PVOB, view.get(), intView.get(), PVOB)
        for y in ARdict:
            if y[0] == i:
                if SecondOrMore is True:
                    CtDeliverStrm += ","
                else:
                    CtDeliverStrm += " "
                    SecondOrMore = True
                CtDeliverStrm += r"activity:%s@\%s" % (y[1], PVOB)
        # make deliver preview
        print "Delivering preview of: %s" % (i)
        com = Popen(CtDeliverStrm, stdout=PIPE, universal_newlines=True, creationflags=0x08000000)
        stdoutdata, stderrdata = com.communicate()  # lines = files.split('N:')[:] # [:] makes mutable list from un-mutable tuple; see: ClearToolPy.py
        print stdoutdata

        if stderrdata != None:
            print "ERROR!!!"
            print stderrdata
            break
        if stdoutdata is None or "No activities to deliver" in stdoutdata: # stdoutdata==None
            continue

        while "C919" in stdoutdata:
            AR = re.search(r'C919[A-Z]\d{8}', stdoutdata)
            if AR is None or AR.group(0) in ARdict:
                break
            AR = AR.group(0)
            print AR
            OutputData.append(AR)
            stdoutdata = stdoutdata.replace(AR, "")
    print OutputData
    print "DONE..."

    # oznacit activity, u kterych neni co deliverovat (i Building AR mohou byt uz deliverovane), ty pak v dalsich krocich ignorovat
    # dulezite! odfiltrovat nasledne dependent activity, ve kterych neni co deliverovat (tzn. ty AR, ktere se pak neobjevi v listingu)
    # teprve pak vylistovat dependent activity uzivateli
    # k nim pak vylistovat change set
    # filtrovat podle komponenty SW_ - kodove komponenty - brat jako kod, textove zmeny v DOC
    # dokumentacni zmeny prijmout automaticky
    # kodove zmeny vyhodit uzivateli
    # od 'delivery' dal implementovat do Perl skriptu
    '''
    Finalize ActivityAnalyzer as standalone project:
    - find dependent activities (try mechanical way by looking for 'h-link'),
    - go through change-set,
    - ignore document changes under DOC folder, highlight code changes under SW_ folders,
    - test it on build machine.
    Add functionality into deliver_activities.pl (part of EBT):
    - find dependent activities (try mechanical way by looking for 'h-link'),
    - go through change-set,
    - add changes DOC folder, highlight code changes under SW_ folders,
    - test it on build machine. 
    '''
    
if __name__ == '__main__':
    form = Tkinter.Tk()
    form.minsize(width=735, height=365)
    form.maxsize(width=735, height=365)
    getFld = Tkinter.IntVar()
    statTxt = "In progress:"
    form.wm_title(' CC/CQ Activity Analyzer')

    stepOne = Tkinter.LabelFrame(form, width=550, height=160, text=" Analyze Freeze List and Make Delivery Pre-view")
    stepOne.grid(row=0, columnspan=7, sticky='w', padx=5, pady=5, ipadx=5, ipady=5)
    stepOne.grid_propagate(0)

    helpLf = Tkinter.LabelFrame(form, text=" Description ")
    helpLf.grid(row=0, column=9, columnspan=2, rowspan=8, sticky='NS', padx=3, pady=5)
    helpLbl = Tkinter.Label(helpLf, wraplength=150, anchor="n", \
                justify='left', text="This script doesn't affect any CC/CQ artefacts. It only delivers data to your fingertips very fast.\
                \n\nIt allows you to analyze either whole freeze list, one or more activities or make delivery pre-view.\
                \n\nThis script will be used mainly for determination of dependencies on different activities.")
    helpLbl.grid(row=0)

    stepTwo = Tkinter.LabelFrame(form, width=550, height=70, text=" Analyze Dependent Activit(y)ies: ")
    stepTwo.grid(row=2, columnspan=7, sticky='w', padx=5, pady=5, ipadx=5, ipady=5)
    stepTwo.grid_propagate(0)


    stepThree = Tkinter.LabelFrame(form, width=550, height=40, text=" 3. Delivery Pre-view: ")
    stepThree.grid(row=3, columnspan=7, sticky='w', padx=5, pady=5, ipadx=5, ipady=5)
    stepThree.grid_propagate(0)

    VOBLbl = Tkinter.Label(stepOne, text="PVOB:")
    VOBLbl.grid(row=0, column=0, sticky='E', padx=5, pady=2)

    VOBTxt = Tkinter.Entry(stepOne, width=25)
    VOBTxt.insert(0, "C919_FC_RC_HI_PVOB")
    VOBTxt.grid(row=0, column=1, sticky="WE", pady=3)

    outFileLbl1 = Tkinter.Label(stepOne, text="Int. Stream:")
    outFileLbl1.grid(row=0, column=5, sticky='E', padx=5, pady=2)

    outFileTxt1 = Tkinter.Entry(stepOne, width=25)
    outFileTxt1.insert(0, "C919_FC_SW_Int")
    outFileTxt1.grid(row=0, column=7, sticky="WE", pady=2)

    ProjLbl = Tkinter.Label(stepOne, text="Project:")
    ProjLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)

    ProjTxt = Tkinter.Entry(stepOne, width=25)
    ProjTxt.insert(0, "C919_FC_SW")
    ProjTxt.grid(row=1, column=1, sticky="WE", pady=3)

    inFileLbl1 = Tkinter.Label(stepOne, text="Integration View:")
    inFileLbl1.grid(row=2, column=0, sticky='E', padx=5, pady=2)

    inFileTxt1 = Tkinter.Entry(stepOne)
    inFileTxt1.insert(0, "H157043_C919_FC_SW_Int_Voltron")
    inFileTxt1.grid(row=2, column=1, columnspan=7, sticky="WE", pady=3)

    inFileBtn1 = Tkinter.Button(stepOne, text="Browse...", command=lambda i=inFileTxt1: FindView(i))
    inFileBtn1.grid(row=2, column=8, sticky='W', padx=5, pady=2)

    outFileLbl2 = Tkinter.Label(stepOne, text="Select Freeze List:")
    outFileLbl2.grid(row=3, column=0, sticky='E', padx=5, pady=2)

    outFileTxt2 = Tkinter.Entry(stepOne)
    outFileTxt2.grid(row=3, column=1, columnspan=7, sticky="WE", pady=2)

    outFileBtn2 = Tkinter.Button(stepOne, text="Browse...", command=lambda i=outFileTxt2: FindFreezeFile(i))
    outFileBtn2.grid(row=3, column=8, sticky='W', padx=5, pady=2)

    anaButton1 = Tkinter.Button(stepOne, text="      Analyze...      ", command=lambda d=outFileTxt2, e=inFileTxt1, f=VOBTxt, g=outFileTxt1, h=ProjTxt: Analyze(d, e, f, g, h))
    anaButton1.grid(row=4, column=7, columnspan=2, sticky='ne', padx=5, pady=5)

    outTblLbl = Tkinter.Label(stepTwo, text="Enter activities:")
    outTblLbl.grid(row=5, column=0, sticky='W', padx=5, pady=2)

    outTblTxt = Tkinter.Entry(stepTwo, width=80)
    outTblTxt.insert(0, " format 'id,id,...' (no spaces)")
    outTblTxt.grid(row=6, column=0, columnspan=3, padx=5, pady=2, sticky='E')
    outTblTxt.grid_propagate(0)

    getFldChk = Tkinter.Checkbutton(stepTwo, text="Get fields", onvalue=1, offvalue=0)
    getFldChk.grid(row=5, column=1, columnspan=3, pady=2, sticky='WE')

    anaButton1 = Tkinter.Button(stepTwo, text="Analyze...", command=Analyze)
    anaButton1.grid(row=6, column=7, columnspan=2, sticky='ne', padx=0, pady=5)

    cancelBtn = Tkinter.Button(form, text="      Quit      ", command=form.destroy)
    cancelBtn.grid(row=9, column=9, columnspan=2, sticky='E', padx=5, pady=2)

    form.mainloop()
