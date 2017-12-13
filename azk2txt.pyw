#
# ask2txt.pyw v. 1.2.3
# Athanassios Protopapas
# 9 December 2011 (save trial order; keep last folder on Mac; fonts)
#
# This program will convert DMDX data from .azk to .txt,
# one row (or column) per subject ID, separated by tab, comma, or space
#
import sys
if (sys.platform=="win32"):
    _WINDOWS_ = True
    _MAC_ = False
elif (sys.platform=="darwin"):
    _WINDOWS_ = False
    _MAC_ = True
else:
    _WINDOWS_ = False
    _MAC_ = False
import os
import string
import math
import datetime
import time
import operator
if _WINDOWS_: import _winreg
elif _MAC_ : import plistlib
from Tkinter import *
import tkFileDialog, tkMessageBox, tkSimpleDialog
import tkFont

## FIXED PARAMETERS set in GlobVariables
DEFAULT_C_DIST=10 # pixels for vertical panel separation
DEFAULT_SUBJ_COL=25 # subjects per column in subject selection frame
DEFAULT_ONE_COL_MAX=20 # maximum number of subjects in a single column
DEFAULT_SUBJ_MAX=100 # above this number, only subject IDs are shown for selection

## PARAMETERS ADJUSTED IN OPENING WINDOW
##
DEFAULT_SAVEFMT=1 # save output in one row per subject
DEFAULT_SEP=0 # separator for output data file defaults to Tab (1=space, 2=comma)
DEFAULT_ENCODING="Latin (ISO-8859-1)" # Encoding for answer string display
DEFAULT_SAVEDATE=1
DEFAULT_SAVETIME=1
DEFAULT_SAVECOMPUTER=0
DEFAULT_SAVEREFRESH=0
DEFAULT_SAVETRIALORDER=0

## BASE FONTS
DEFAULT_SCALE=9 # main font size under Windows, used as a scaling factor
DEFAULT_FONTFAMILY="Helvetica"
if _WINDOWS_: DEFAULT_FONTSIZE=9
elif _MAC_: DEFAULT_FONTSIZE=12
else: DEFAULT_FONTSIZE=9

## MICSELLANEOUS PARAMETERS
##
UBPATH='~/Library/Application Support'
APPNAME="CheckVocal"
PLISTFNAME="azk2txt_param.plist"
CV_REGISTRY_KEY="SOFTWARE\\azk2txt\\" # location of last folder key
PRINT_ENCODING="iso-8859-1" # for printing messages on console window
ENCODING_OPTIONS=["Latin (ISO-8859-1)",
                  "European (ISO-8859-2)",
                  "Esperanto (ISO-8859-3)",
                  "Baltic (ISO-8859-4)",
                  "Cyrillic (ISO-8859-5)",
                  "Arabic (ISO-8859-6)",
                  "Greek (ISO-8859-7)",
                  "Hebrew (ISO-8859-8)",
                  "Turkish (ISO-8859-9)",
                  "Chinese (EUC-CN)",
                  "Japanese (EUC-JP)",
                  "Korean (EUC-KR)"]

# Variables and settings that need to be available to various widgets and processes
#
class GlobVariables:

    def __init__(self):
        self.encoding=StringVar(root)
        self.savechoice=IntVar(root)
        self.sepchoice=IntVar(root)
        self.azkff=StringVar(root)
        self.savedate=IntVar(root)
        self.savetime=IntVar(root)
        self.savecomputer=IntVar(root)
        self.saverefresh=IntVar(root)
        self.savetrialorder=IntVar(root)
        #
        self.azkfilename=""
        self.expdir="."
        self.expname=""
        self.logfile=""
        self.Nsubj=0
        self.sub_trials={}
        self.sub_order={}
        self.done=0
        self.SRATE=0
        self.totaltrials=0
        self.trial_ind={}
        self.subj_ind={}
        self.sub_ids={}
        self.sub_ids_new={}
        self.sub_dates={}
        self.sub_nums={}
        #
        self.listofanswers=[]
        self.listoftrials=[]
        self.listoffiles=[]
        self.listoftrialsub=[]
        self.listofproblems=[]
        self.listofloaded=[]
        self.original_listoftrials=[]
        #
        ## Fixed parameters
        self._C_DIST=DEFAULT_C_DIST
        self._SUBJ_COL=DEFAULT_SUBJ_COL
        self._ONE_COL_MAX=DEFAULT_ONE_COL_MAX
        self._SUBJ_MAX=DEFAULT_SUBJ_MAX
        #
        ## Fonts
        self.fontfamily=DEFAULT_FONTFAMILY
        self.fontsize=DEFAULT_FONTSIZE
        #self.largefontsize=DEFAULT_FONTSIZE*10/9
        self.mainfont=tkFont.Font(family=self.fontfamily,size=self.fontsize,weight=tkFont.NORMAL,slant=tkFont.ROMAN)
        self.mainboldfont=tkFont.Font(family=self.fontfamily,size=self.fontsize,weight=tkFont.BOLD,slant=tkFont.ROMAN)
        #self.largefont=tkFont.Font(family=self.fontfamily,size=self.largefontsize,weight=tkFont.NORMAL,slant=tkFont.ROMAN)
        #self.largeboldfont=tkFont.Font(family=self.fontfamily,size=self.largefontsize,weight=tkFont.BOLD,slant=tkFont.ROMAN)
        
    def scale(self,w):
        return w*DEFAULT_FONTSIZE/DEFAULT_SCALE
    
    def update(self):
        tmp_enc=self.encoding.get().split()[-1]
        self.char_encoding = tmp_enc.strip('()')
        self.save_rows=self.savechoice.get()
        if (self.sepchoice.get()==0):
            self._SEP="\t"
        elif (self.sepchoice.get()==1):
            self._SEP=" "
        elif (self.sepchoice.get()==2):
            self._SEP=","

    def reset(self):
        self.encoding.set(DEFAULT_ENCODING)
        self.savechoice.set(DEFAULT_SAVEFMT)
        self.sepchoice.set(DEFAULT_SEP)
        self.savedate.set(DEFAULT_SAVEDATE)
        self.savetime.set(DEFAULT_SAVETIME)
        self.savecomputer.set(DEFAULT_SAVECOMPUTER)
        self.saverefresh.set(DEFAULT_SAVEREFRESH)
        self.savetrialorder.set(DEFAULT_SAVETRIALORDER)
        self.azkff.set("")
        self.read_lastfolder()

    def read_lastfolder(self): # Read the last selected folder from the registry if possible
        if not (_WINDOWS_ or _MAC_):
            self.lastfolder="."
            # we may add here something to keep track of things in a file...
            return

        if _MAC_:
            basepath = os.path.expanduser(UBPATH)
            cvpath = os.path.join(basepath,APPNAME)
            try:
                pl=plistlib.readPlist(os.path.join(cvpath,PLISTFNAME))
            except IOError: # nonexistent file
                self.lastfolder="."
            else:
                self.lastfolder=pl["lastfolder"]
                if (not os.path.exists(self.lastfolder)): # path not available (deleted, on removable media..)
                    self.lastfolder="."
            return

        try: # WINDOWS #
            rkey=_winreg.CreateKey(_winreg.HKEY_LOCAL_MACHINE,CV_REGISTRY_KEY)
            self.lastfolder=_winreg.QueryValueEx(rkey,"LastFolder")[0]
            _winreg.CloseKey(rkey)
        except WindowsError: # nonexistent key or error reading registry
            self.lastfolder=""    # an empty string to return false from os.path.exists()
            
        if (not os.path.exists(self.lastfolder)): # path not available (deleted, on removable media..)
            try: # retrieve user "home" from the registry
                dkey=_winreg.OpenKey(_winreg.HKEY_CURRENT_USER,"Volatile Environment\\")
                homedrive=_winreg.QueryValueEx(dkey,"HOMEDRIVE")[0]
                homepath=_winreg.QueryValueEx(dkey,"HOMEPATH")[0]
                _winreg.CloseKey(dkey)
                self.lastfolder=homedrive+homepath
            except WindowsError: # nonexistent key or error reading registry
                self.lastfolder="."


# General message display in log file
#
def logmsg(logtext):
    # added encoding to deal with printing out non-ascii path components
    gv.logfile.write(logtext.encode(gv.char_encoding, 'replace')+"\n")
    gv.logfile.flush()

# Fatal program failure, logging and displaying message before exiting
#
def exiterror(errtext):
    logmsg("**ERROR: "+errtext)
    tkMessageBox.showerror("azk2txt error",errtext)
    global_quit()

# Deal with non-terminated text files by adding a final newline
#
def myreadlines(f):
    rlines=f.readlines()
    if (rlines[-1][-1] != "\n"):
        last=rlines[-1]+"\n"
        rlines=rlines[:-1]
        rlines.append(last)
    return rlines


# The parameter setup/configure window
#
class SetupWindow(Toplevel):

    def __init__(self,parent):
        
        Toplevel.__init__(self,parent)
##        self.transient(parent) ### NO NO NO
        self.parent=parent

        self.title(u"azk2txt setup")
        self.geometry("+%d+%d" % (gv.scale(100),gv.scale(150)))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)

        self.contF0=Frame(self)
        self.cbutton0=Button(self.contF0,text="Proceed",command=self.proceed)
        self.cbutton0.config(width=12,font=gv.mainboldfont,anchor="s")
        self.cbutton0.pack(side="right",padx=24)
        self.cbutton01=Button(self.contF0,text="Reset",command=self.reset)
        self.cbutton01.config(width=8,font=gv.mainfont,anchor="s")
        self.cbutton01.pack(side="right",padx=5)
        self.cbutton02=Button(self.contF0,text="Cancel",command=self.mquit)
        self.cbutton02.config(width=8,font=gv.mainfont,anchor="s")
        self.cbutton02.pack(side="right",padx=5)
        self.contF0.pack(side="bottom",pady=10)

        self.encF2=Frame(self)
        self.rlabel2=Label(self.encF2,text="Character encoding:")
        self.rlabel2.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel2.pack(side="left")
        self.encodingmenu2=apply(OptionMenu,(self.encF2,gv.encoding)+tuple(ENCODING_OPTIONS)) # from http://effbot.org/tkinterbook/optionmenu.htm
        self.encodingmenu2.config(width=15,font=gv.mainfont)
        self.encodingmenu2.pack(side="left")
        self.encF2.pack(side="bottom",anchor="w")

        self.sepF4=Frame(self)
        self.rlabel4=Label(self.sepF4,text="Separator:")
        self.rlabel4.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel4.pack(side="left")
        self.sepchoice1=Radiobutton(self.sepF4,text="Tab",variable=gv.sepchoice,value=0)
        self.sepchoice1.config(font=gv.mainfont)
        self.sepchoice1.pack(side="left",anchor="s")
        self.sepchoice2=Radiobutton(self.sepF4,text="Space",variable=gv.sepchoice,value=1)
        self.sepchoice2.config(font=gv.mainfont)
        self.sepchoice2.pack(side="left",anchor="s")
        self.sepchoice3=Radiobutton(self.sepF4,text="Comma",variable=gv.sepchoice,value=2)
        self.sepchoice3.config(font=gv.mainfont)
        self.sepchoice3.pack(side="left",anchor="s")
        self.sepF4.pack(side="bottom",anchor="w",pady=1)

        self.saveF9=Frame(self)
        self.rlabel9=Label(self.saveF9,text="Include:")
        self.rlabel9.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel9.pack(side="left")
        self.savebox1=Checkbutton(self.saveF9,text="date",variable=gv.savedate)
        self.savebox1.config(font=gv.mainfont)
        self.savebox1.pack(side="left",anchor="s")
        self.savebox2=Checkbutton(self.saveF9,text="time",variable=gv.savetime)
        self.savebox2.config(font=gv.mainfont)
        self.savebox2.pack(side="left",anchor="s")
        self.savebox3=Checkbutton(self.saveF9,text="computer",variable=gv.savecomputer)
        self.savebox3.config(font=gv.mainfont)
        self.savebox3.pack(side="left",anchor="s")
        self.savebox4=Checkbutton(self.saveF9,text="refresh",variable=gv.saverefresh)
        self.savebox4.config(font=gv.mainfont)
        self.savebox4.pack(side="left",anchor="s")
        self.savebox5=Checkbutton(self.saveF9,text="trial order",variable=gv.savetrialorder)
        self.savebox5.config(font=gv.mainfont)
        self.savebox5.pack(side="left",anchor="s")
        self.saveF9.pack(side="bottom",anchor="w",pady=1)

        self.saveF8=Frame(self)
        self.rlabel8=Label(self.saveF8,text="Save subject data in:")
        self.rlabel8.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel8.pack(side="left")
        self.savechoice1=Radiobutton(self.saveF8,text="rows",variable=gv.savechoice,value=1,command=self.savenew)
        self.savechoice1.config(font=gv.mainfont)
        self.savechoice1.pack(side="left",anchor="s")
        self.savechoice2=Radiobutton(self.saveF8,text="columns",variable=gv.savechoice,value=0,command=self.savenew)
        self.savechoice2.config(font=gv.mainfont)
        self.savechoice2.pack(side="left",anchor="s")
        self.savechoice3=Radiobutton(self.saveF8,text="AZK file",variable=gv.savechoice,value=-1,command=self.savenew)
        self.savechoice3.config(font=gv.mainfont)
        self.savechoice3.pack(side="left",anchor="s")
        self.savechoice4=Radiobutton(self.saveF8,text="long format",variable=gv.savechoice,value=2,command=self.savenew)
        self.savechoice4.config(font=gv.mainfont)
        self.savechoice4.pack(side="left",anchor="s")
        self.saveF8.pack(side="bottom",anchor="w")

        self.azkF1=Frame(self)
        self.rlabel1=Label(self.azkF1,text="DMDX results file:")
        self.rlabel1.config(width=20,font=gv.mainboldfont,anchor="e")
        self.rlabel1.pack(side="left")
        self.rmessage1=Label(self.azkF1,textvariable=gv.azkff,anchor="w")
        self.rmessage1.config(width=60,font=gv.mainfont,background="white",foreground="black")
        self.rmessage1.bind("<Button-1>",self.get_filename)
        self.rmessage1.pack(side="left",pady=5)
        self.rlabel10=Label(self.azkF1,text=" ")
        self.rlabel10.config(font=gv.mainboldfont)
        self.rlabel10.pack(side="left")
        self.azkF1.pack(side="bottom",anchor="w")

##    def startup(self):
        self.initial_focus = self
        self.initial_focus.focus_set()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.mquit)
        self.bind("<Return>", self.proceed)
        self.bind("<Escape>", self.mquit)
        self.focus_force()
        self.get_filename()
        self.wait_window(self)

    def savenew(self):
        if gv.savechoice.get()==-1:
            self.savebox1.config(state=DISABLED)
            self.savebox2.config(state=DISABLED)
            self.savebox3.config(state=DISABLED)
            self.savebox4.config(state=DISABLED)
            self.savebox5.config(state=DISABLED)
        else:
            self.savebox1.config(state=NORMAL)
            self.savebox2.config(state=NORMAL)
            self.savebox3.config(state=NORMAL)
            self.savebox4.config(state=NORMAL)
            self.savebox5.config(state=NORMAL)

    def proceed(self,c_event=None):
        gv.update()
        self.destroy()
        cv_process.run()

    def mquit(self,c_event=None):
        global_quit()
    
    def reset(self):
        gv.reset()

    def save_expdir(self):
        if _WINDOWS_:
            try: # save selected expdir as last folder and close registry key
                rkey=_winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,CV_REGISTRY_KEY,0,_winreg.KEY_SET_VALUE)
                _winreg.SetValueEx(rkey,"LastFolder",0,_winreg.REG_SZ,gv.expdir)
                _winreg.CloseKey(rkey)
            except: # failed to update registry
                pass # fail silently; msgwindow not available and logfile not yet open
        elif _MAC_:
            basepath = os.path.expanduser(UBPATH)
            cvpath = os.path.join(basepath,APPNAME)
            if not os.path.exists(cvpath):
                os.mkdir(cvpath)
            pl=dict(lastfolder=gv.expdir)
            try:
                plistlib.writePlist(pl,os.path.join(cvpath,PLISTFNAME))
            except: # failed to update
                pass # fail silently; msgwindow not available and logfile not yet open
        
    def get_filename(self,c_event=None): 
        filename=tkFileDialog.askopenfilename(parent=self.parent,initialdir=gv.lastfolder,filetypes=[('DMDX data files','*.azk')] ,title="Choose a DMDX results file")
        if (len(filename)>0):
            gv.expdir=os.path.dirname(filename)+"/"
            gv.lastfolder=gv.expdir
            try:
                os.chdir(gv.expdir)
            except OSError:
                exiterror("Unable to work in this folder -- check security permissions")
            gv.expname=os.path.splitext(os.path.basename(filename))[0]
            gv.azkff.set(filename)
            self.cbutton0.config(state="normal")
            #self.fileOK.set(True)
            self.save_expdir()
        elif (len(gv.azkff.get())==0):
            self.cbutton0.config(state="disabled")


# Select subjects to process using a simple CheckButton GUI
#
class SubjectSelect(Toplevel):

    def subj_select_ok(self): # end-of-selection callback
        self.quit()
        self.destroy()

    def subj_select_all(self): # select all available subjects
        for subject in gv.sub_trials.keys():
            if (cv_process.subject_select[subject]):
                self.c[subject].select()

    def subj_select_none(self): # deselect all available subjects
        for subject in gv.sub_trials.keys():
            if (cv_process.subject_select[subject]):
                self.c[subject].deselect()
        
    def __init__(self,parent):

        Toplevel.__init__(self,parent)
        self.parent=parent
        self.title(u"DMDX subject selection")
        self.geometry("+%d+%d" % (100,175))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)

        self.sublabel=Label(self, text="Select subjects to process data", font=gv.mainboldfont)
        self.sublabel.pack(padx=30,pady=5)
        self.newframe=Frame(self) 
        self.subframe=Frame(self.newframe)
        self.sub_buttons={}
        multi_column = 0
        sub_incol = 0
        self.c={}

        gv.Nsubj = len(gv.sub_trials.keys())
        if (gv.Nsubj > 100):
            logmsg( "Number of subjects (%i) probably too large to fit date/PC info" % (gv.Nsubj))
        if (gv.Nsubj > gv._ONE_COL_MAX):
            multi_column = 1
            column_length = math.ceil (gv.Nsubj / math.ceil(float(gv.Nsubj)/float(gv._SUBJ_COL)) )
        ##
        ##for subject in gv.sub_trials.keys():
        live_subjects=gv.sub_ids_new.keys()
        for subjnum in live_subjects:
            subject=gv.sub_ids_new[subjnum]
        ##
            self.sub_buttons[subject]=IntVar()
            if (gv.Nsubj > gv._SUBJ_MAX): # do not display dates with so many subjects, to fit more columns
                self.c[subject]=Checkbutton(self.subframe,text=subject,variable=self.sub_buttons[subject])
            else:
                self.c[subject]=Checkbutton(self.subframe,text=subject+" "+gv.sub_dates[subject],variable=self.sub_buttons[subject])
            if (not cv_process.subject_select[subject]):
                self.sub_buttons[subject].set(0)
                self.c[subject].deselect()
                self.c[subject].configure(state=DISABLED) # do not allow selection of subjects with data problems
            else:
                self.c[subject].select()
            self.c[subject].pack(side='top',anchor=W)
            sub_incol += 1
            if (multi_column==1 and sub_incol>=column_length):
                self.subframe.pack(pady=5,side='left')
                self.subframe=Frame(self.newframe) 
                sub_incol=0
        self.subframe.pack(pady=5,side='left',anchor=N)
        self.newframe.pack(side='top')
        self.buttonframe=Frame(self)
        self.allbutton=Button(self.buttonframe,text="All",width=10,font=gv.mainfont,command=self.subj_select_all)
        self.nonebutton=Button(self.buttonframe,text="None",width=10,font=gv.mainfont,command=self.subj_select_none)
        self.okbutton=Button(self.buttonframe,text="OK",width=20,font=gv.mainboldfont,command=self.subj_select_ok)
        self.allbutton.pack(padx=2,side='left')
        self.nonebutton.pack(padx=2,side='left')
        self.okbutton.pack(padx=4,side='left')
        self.buttonframe.pack(pady=10,side='top')


# This is where all of the preparation and track-keeping is done
#
class azkConvertClass:

    def __init__(self):
        # we need the instance created so that it is globally available
        pass

    def run(self):

        logfilename=gv.expname+"-msg.txt"
        try:
            gv.logfile=open(gv.expdir+logfilename,"w")
        except:
            # this is the same as exiterror but is called directly
            # because exiterror also writes to the logfile,
            # which here cannot be opened
            tkMessageBox.showerror("Log file error",
                      "Could not open %s to write processing log" % (logfilename))
            global_quit()

        logmsg("Session started : "+str(datetime.datetime.now()))
        gv.azkfilename=gv.expname+".azk"
        try:
            azkfile=open(gv.expdir+gv.azkfilename)
        except:
            exiterror("Could not open %s to read experiment data" % (gv.azkfilename))

        logmsg("Working folder is %s" % gv.expdir)
        azklines=myreadlines(azkfile)
        logmsg("Opened %s" % gv.azkfilename)
        azkfile.close()

        logmsg("Processing experiment "+gv.expname)

        Nlines=len(azklines)
        subjno=0
        self.ntrials=0
        line=0

        sub_origlines={}
        self.subject_select={}

        _COT_ = 0

        # skip initial blank lines
        while (len(azklines[line])<3):
            line += 1

        # read total subject information
        if (azklines[line][:31]=="Subjects incorporated to date: "):
            gv.Nsubj = string.atoi(azklines[line][31:-1]) # discard newline
        else:
            exiterror("General subject information not found")
            
        for subj in range (gv.Nsubj):
            #
            # parse file to extract trial information for each of Nsubj subjects
            #
            while (len(azklines[line])<11 or azklines[line][:10] != "**********"):
                line += 1
                if (line >= len(azklines)):
                    logmsg("File ended at line %i before processing %i subjects" % (line,gv.Nsubj))
                    # perhaps data have been manually removed
                    subj=gv.Nsubj
                    break

            try:
                subj_,date_,refresh_,ids_ = string.split(azklines[line+1][:-1],',') # discard newline
            except ValueError:
                logmsg("Insufficient fields for subject %i (probably missing ID)"%(subjno+1))
                # useless data without ID, try to skip this subject
                # assume there will always be subject, date, and refresh fields
                subj_,date_,refresh_ = string.split(azklines[line+1][:-1],',') # discard newline
                ids_="xxx xxx" # so that a dummy ID will be made up below
            s_ = string.atoi(string.split(subj_)[1])
            if (subj_[:7]!="Subject" or subjno+1 != s_):
                logmsg("Unexpected subject number "+`s_`+" (expected "+`subjno+1`+") at line "+`line`)
                # perhaps a subject's data have been manually removed from the .azk
            subjno += 1
            idtmp=string.split(unicode(ids_,gv.char_encoding)) # ID might be in non-latin characters...
            if (len(idtmp)>1):
                s_id = idtmp[1]
            elif (len(ids_)>3 and ids_[1:3]=="ID"):
                s_id = ids_[3:]
            elif (len(idtmp[0])>1):
                s_id = idtmp[0]
                logmsg("Malformed ID information for subject %i at line %i, will use %s" % (s_,line+1,s_id))
            else:
            ##if (ids_[1:3]!="ID" or len(s_id)<1):
                s_id = "SID%04i" % s_
                logmsg("Could not determine ID for subject %i at line %i, will use %s" % (s_,line+1,s_id))

            gv.sub_ids_new[subjno]=s_id
            for s_temp in gv.sub_ids.keys():
                if (s_id == gv.sub_ids[s_temp]):
                    try:
                        logmsg( "Duplicate subject ID %s (subjects %i and %i)" % (s_id,gv.sub_nums[s_id],s_))
                    except KeyError:
                        logmsg( "Duplicate subject ID %s for subject %i (and probably two or more other subjects)" % (s_id,s_))
                    gv.sub_ids_new[subjno]=s_id+"_S#%03i"%(s_)
                    logmsg( "Will use ID %s for subject %i"%(gv.sub_ids_new[subjno],s_) )
                    if ((s_id in gv.sub_nums.keys()) and # If two or more previous subjects have the same ID
                                                         # then sub_nums is indexed by the modified IDs already
                                                         # and it is a mess to figure out which IDs go with which
                                                         # experimental runs. The good news is that in such a case
                                                         # no pre-existing IDs need be modified -- the bad news is
                                                         # that it becomes problematic to check things properly if
                                                         # subject IDs are repeated in different experimental runs
                        (gv.sub_ids_new[gv.sub_nums[s_id]][-6:-3]!="_S#")): # don't append again if more repetitions of an ID are found
                        other_id=gv.sub_ids[s_temp]
                        other_s=gv.sub_nums[other_id]
                        new_other_id=gv.sub_ids[other_s] + "_S#%03i"%(other_s)
                        gv.sub_ids_new[other_s]=new_other_id
                        logmsg ( "Also changing ID from %s to %s for subject %i"%(gv.sub_ids[other_s],gv.sub_ids_new[other_s],other_s) )
                        self.subject_select[new_other_id]=self.subject_select[other_id]
                        gv.sub_dates[new_other_id]=gv.sub_dates[other_id]
                        gv.sub_nums[new_other_id]=gv.sub_nums[other_id]
                        gv.sub_trials[new_other_id]=gv.sub_trials[other_id]
                        gv.sub_order[new_other_id]=gv.sub_order[other_id]
                        sub_origlines[new_other_id]=sub_origlines[other_id]
                        self.subject_select.pop(other_id)
                        gv.sub_dates.pop(other_id)
                        gv.sub_nums.pop(other_id)
                        gv.sub_trials.pop(other_id)
                        gv.sub_order.pop(other_id)
                        sub_origlines.pop(other_id)

            gv.sub_ids[subjno]=s_id
            new_s_id=gv.sub_ids_new[subjno]
            logmsg( "Subject "+`subjno`+", ID="+new_s_id)

            self.subject_select[new_s_id]=1

            # check date, time, computer
            datetimestr=date_.split()
            if (len(datetimestr)<4 or datetimestr[2]!="on"):
                logmsg( "Date/time/PC parsing error, will use ****" )
                date_ = "**/**/** **:**:** on ******"
            #
            gv.sub_dates[new_s_id]=date_
            gv.sub_nums[new_s_id]=subjno

            RTheaders = string.split(azklines[line+2][:-1])
            if (len(RTheaders)>2): # Recognize a 3rd column of COT data (clock-on-trial)
                if (RTheaders[2] == "COT"):
                    if (_COT_ == 0):
                        logmsg( "COT header detected")
                    _COT_ = 1
                else:
                    logmsg( "Unknown header identifier at line "+`line+2`)
                    self.subject_select[new_s_id]=0  # do not process subjects with not understood data
                    
            if (RTheaders[:2] != ["Item","RT"]):
                logmsg( "Item/RT identifier not found at line "+`line+2`)
                self.subject_select[new_s_id]=0  # definitely kill the subject with unparseable data
                # This is a problem for the subject's data but perhaps the rest of the file is OK

            line += 3
            startline = line
            exlines = 0
            # look for the first empty line; allow an abrupt break into next subject
            while (line<Nlines and len(azklines[line])>2 and (azklines[line][:2] != "**")):
                if (azklines[line][0] == '!'):
                    exlines += 1
                    # catch a case of split error-report lines on non-ASCII strings ## thp 2006-10-13
                    if (line<(Nlines-1) and len(azklines[line+1])>0 and azklines[line+1][0]!='!' and
                        string.count(azklines[line],'"')>0 and string.count(azklines[line+1],'"')>0):
                        exlines += 1
                line += 1
            # recover from a premature break
            if ((line<Nlines) and (len(azklines[line])>2) and (azklines[line][:2]=="**")):
                    line -= 1
            endline = line

            if (self.ntrials > 0): # skip check for first subject
                if (endline-startline-exlines != self.ntrials):
                    logmsg( "Mismatching number of trials (expected %i, found %i) for subject %i (ID %s)" % (self.ntrials, endline-startline-exlines, subjno, new_s_id))
                    self.subject_select[new_s_id]=0  # definitely kill the subject with unparseable data
                    # do not die on poor data, perhaps more good subject data follow
            else:
                self.ntrials = endline-startline-exlines

            tmptrials=[]
            itrial=0
            previous_line=" "
            for trial in azklines[startline:endline]:
                # catch a case of split error-report lines on non-ASCII strings ## thp 2006-10-13
                if (not (previous_line[0]=='!' and trial[0] != '!' and 
                         string.count(previous_line,'"')>0 and string.count(trial,'"')>0)):
                    if (trial[0] != '!'):
                        if (_COT_ == 1): # parse correctly data even if a 3rd (COT) column is present
                            item_,rt_,cot_ = string.split(trial)[:3] # COT data will be ignored
                        else:
                            item_,rt_ = string.split(trial)[:2] # also ignore "ABORTED" or other following data
                        item=string.atoi(item_)
                        rt=string.atof(rt_)
                        itrial += 1 # increase trial counter to store order
                        tmptrials.append((item,rt,itrial))
                previous_line=trial # retain for comparison
            tmptrials.sort(key=operator.itemgetter(0)) # don't sort on RTs and order
            tmptimes=[(x[0],x[1]) for x in tmptrials]
            tmporder=[(x[0],x[2]) for x in tmptrials]
            # this is the list of trial data for the subject, by increasing item number
            gv.sub_trials[new_s_id]=tmptimes
            # this is the order of trial presentation
            gv.sub_order[new_s_id]=tmporder
            # also save the original data lines in case we need to save in .azk
            sub_origlines[new_s_id]=azklines[startline:endline]

        logmsg( "End of processing trials in %s, %i lines left" % (gv.azkfilename,Nlines-line))

        for cur_subj in gv.sub_trials.keys():
            sub_trial_ind=0
            for trial in gv.sub_trials[cur_subj]:
                item,rt=trial
                gv.listoftrials.append(trial)
                gv.listoftrialsub.append(cur_subj)
                gv.trial_ind[gv.totaltrials]=sub_trial_ind
                gv.subj_ind[gv.totaltrials]=cur_subj
                sub_trial_ind += 1
                gv.totaltrials += 1

        # no point in moving on if there are no valid subject data
        valid_subjects=0
        for subject in gv.sub_trials.keys():
            if (self.subject_select[subject]==1): valid_subjects += 1
        if (valid_subjects==0):
            exiterror( "No valid subject data to process!")            

        # save a copy of the unaltered list of trials to be able to revert to
        gv.original_listoftrials=gv.listoftrials[:]

        continue_old=0

        if (continue_old == 0): # starting a new session from scratch

            self.current_index=0
            self.N_done=0
            # make sure we are at a writable directory before proceeding

            subselect=SubjectSelect(root)
            subselect.focus_force()
            subselect.wait_window(subselect)

            for subject in gv.sub_trials.keys():
                self.subject_select[subject]=subselect.sub_buttons[subject].get()

        # Remove deselected subject entries from the sub_trials dictionary
        for subject in gv.sub_trials.keys():
            if (not self.subject_select[subject]):
                del(gv.sub_trials[subject])

        # End of status restoration and subject verification       
        if (len(gv.sub_trials.keys()) < 1):
            exiterror( "No subjects left to process!")
           
        # Make sure items are numbered identically and sorted correctly between subjects
        ref_subj=gv.sub_trials.keys()[0]
        for cur_subj in gv.sub_trials.keys()[1:]:
            for trial in range(self.ntrials):
                if (gv.sub_trials[ref_subj][trial][0] != gv.sub_trials[cur_subj][trial][0]):
                        exiterror( "Trial (item) number mismatch between subjects %i (%s, item %i) and %i (%s, item %i)" % (gv.sub_nums[ref_subj],ref_subj,gv.sub_trials[ref_subj][trial][0],gv.sub_nums[cur_subj],cur_subj,gv.sub_trials[cur_subj][trial][0]))

        # prepare counters
        gv.done=1 # was 0 in CheckVocal
        self.N_todo=len(gv.sub_trials.keys())* self.ntrials # how many responses there are to be checked in total

        ############################################################################
        
        if (gv.save_rows==-1): # if AZK format output is requested, name the file accordingly
            outfilename=gv.expname+"-a2t.azk"
        else:
            outfilename=gv.expname+".txt"
        outfileOK=0
        try:
            outfile=open(outfilename)
            outfile.close()
        except IOError:
            outfileOK=1
            outfile=open(outfilename,"w") # hopefully this will work
        while (outfileOK<1):
            if (outfileOK==0):
                askmessagetext="Output file %s exists, overwrite? (y/n)" % (outfilename)
                ans=tkMessageBox.askyesno("File overwrite",askmessagetext)
            else: # -1 means we are looping
                ans=False
            if (ans): # overwrite
                outfileOK=1
            else: # ans=False
                outfilename=tkFileDialog.asksaveasfilename(parent=root,initialdir=gv.expdir,filetypes=[('Text files','*.txt')] ,
                                                           title="Enter another output file name")
            try:
                outfile=open(outfilename,"w")
                outfileOK=1
            except IOError:
                outfileOK=-1

        #live_subjects=gv.sub_ids_new.keys() # copying corrections from CheckVocal8 1.4b (14/7/06)
        # make a new list from sub_ids.keys() instead of sub_trials.keys()
        # in order to save the data in the original subject order (as in azk)
        live_subjects=[] 
        for subjnum in gv.sub_ids_new.keys(): # gv.sub_trials.keys() # was sub_ids.keys() // thp 26-Nov-06
            if (self.subject_select[gv.sub_ids_new[subjnum]]): # was sub_ids[subjnum]     // thp 26-Nov-06
                live_subjects.append(subjnum)          
                
        if (gv.save_rows==-1): # save subject data in AZK file
            outfile.write("\nSubjects incorporated to date: %03d\n"%len(gv.sub_trials.keys()))
            outfile.write("Data file started on machine CheckVocal\n")
            subjno=0
            # Need to re-sort by sub_origlines[s_id] (which was set to azklines[startline:endline])
            # and add COT if present in the original 
            ##for cur_subj in gv.sub_trials.keys():
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids_new[subjnum]
                subjno+=1
                sub_newlines={}
                outfile.write("\n**********************************************************************\n")
                outfile.write("Subject %d,%s,%s, ID %s\n" % (subjno,gv.sub_dates[cur_subj],refresh_,cur_subj.encode(gv.char_encoding)))
                if (_COT_):
                    outfile.write("  Item       RT       COT\n")
                else:
                    outfile.write("  Item       RT\n")
                for trial in gv.sub_trials[ref_subj]:
                    sub_newlines[`trial[0]`]=("%6d  %8.2f\n" % (trial[0],gv.sub_trials[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                for trial_line in sub_origlines[cur_subj]:
                    if (trial_line[0]=="!"):
                        outfile.write(trial_line)
                    else:
                        trialitem=string.split(trial_line)[0]
                        if (_COT_):
                            cotstr=" "+("%9.2f"%(string.atof(string.split(trial_line)[2])))
                        else:
                            cotstr=""
                        # need to remove final \n from trial line so it can be re-appended after optional COT
                        outfile.write(sub_newlines[trialitem][:-1]+cotstr+"\n")
        elif (gv.save_rows==0): # save subject data in columns
            outfile.write("item")
            ##for cur_subj in gv.sub_trials.keys():
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids_new[subjnum]
                outfile.write(gv._SEP+cur_subj.encode(gv.char_encoding)) # in case of non-latin IDs
            outfile.write("\n")
##
            if gv.savedate.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids_new[subjnum]
                    timedatestr=gv.sub_dates[cur_subj]
                    curdate=timedatestr.split()[0]
                    outfile.write(gv._SEP+curdate.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.savetime.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids_new[subjnum]
                    timedatestr=gv.sub_dates[cur_subj]
                    curtime=timedatestr.split()[1]
                    outfile.write(gv._SEP+curtime.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.savecomputer.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids_new[subjnum]
                    timedatestr=gv.sub_dates[cur_subj]
                    curcomputer=timedatestr.split(None,3)[3] # sep=None, maxsplit=3; computer name may contain spaces!
                    outfile.write(gv._SEP+curcomputer.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.saverefresh.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids_new[subjnum]
                    currefresh=refresh_.split()[1]
                    outfile.write(gv._SEP+currefresh.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
##                
            for trial in gv.sub_trials[ref_subj]:
                outfile.write(`trial[0]`)
                for subjnum in live_subjects:         # was gv.sub_trials.keys():   // thp 26-Nov-06
                    cur_subj=gv.sub_ids_new[subjnum]  # added to get ID from number // thp 26-Nov-06
                    outfile.write(gv._SEP+"%.1f"%(gv.sub_trials[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                outfile.write("\n")
            ## save trial order after RT; added 12/2011
            if gv.savetrialorder.get()==1:
                for trial in gv.sub_trials[ref_subj]:
                    outfile.write("ord"+`trial[0]`)
                    for subjnum in live_subjects:
                        cur_subj=gv.sub_ids_new[subjnum]
                        outfile.write(gv._SEP+"%i"%(gv.sub_order[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                    outfile.write("\n")

        elif (gv.save_rows==2): # save data in long format for R
            outfile.write("subject")
            if gv.savedate.get()==1:
                outfile.write(gv._SEP+"date")
            if gv.savetime.get()==1:
                outfile.write(gv._SEP+"time")
            if gv.savecomputer.get()==1:
                outfile.write(gv._SEP+"computer")
            if gv.saverefresh.get()==1:
                outfile.write(gv._SEP+"refresh")
            if gv.savetrialorder.get()==1: # save trial order; added 12/2011
                outfile.write(gv._SEP+"order")
            outfile.write(gv._SEP+"item"+gv._SEP+"RT"+"\n")
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids_new[subjnum]
                curdate,curtime,on_,curcomputer=gv.sub_dates[cur_subj].split(None,3)  # sep=None, maxsplit=3; computer name may contain spaces!
                for trial in gv.sub_trials[cur_subj]:
                    outfile.write(cur_subj.encode(gv.char_encoding)) # in case of non-latin IDs
                    if gv.savedate.get()==1:
                        outfile.write(gv._SEP+curdate)
                    if gv.savetime.get()==1:
                        outfile.write(gv._SEP+curtime)
                    if gv.savecomputer.get()==1:
                        outfile.write(gv._SEP+curcomputer)
                    if gv.saverefresh.get()==1:
                        currefresh=refresh_.split()[1]
                        outfile.write(gv._SEP+currefresh)
                    if gv.savetrialorder.get()==1: # save trial order; added 12/2011
                        outfile.write(gv._SEP+"%i"%(gv.sub_order[cur_subj][gv.sub_trials[cur_subj].index(trial)][1]))
                    outfile.write((gv._SEP+"%i"+gv._SEP+"%.1f")%(trial[0],trial[1]))
                    outfile.write("\n")

        else: # normally save_rows=1, save data in rows; but set as default just in case
            outfile.write("subject")
            if gv.savedate.get()==1:
                outfile.write(gv._SEP+"date")
            if gv.savetime.get()==1:
                outfile.write(gv._SEP+"time")
            if gv.savecomputer.get()==1:
                outfile.write(gv._SEP+"computer")
            if gv.saverefresh.get()==1:
                outfile.write(gv._SEP+"refresh")
            for trial in gv.sub_trials[ref_subj]:
                outfile.write(gv._SEP+`trial[0]`)
            ## Optionally, save trial order after RT; added 12/2011
            if gv.savetrialorder.get()==1: 
                for trial in gv.sub_order[ref_subj]:
                    outfile.write(gv._SEP+"ord"+`trial[0]`)
            outfile.write("\n")
            ##for cur_subj in gv.sub_trials.keys():
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids_new[subjnum]
                outfile.write(cur_subj.encode(gv.char_encoding)) # in case of non-latin IDs
                #
                curdate,curtime,on_,curcomputer=gv.sub_dates[cur_subj].split(None,3)  # sep=None, maxsplit=3; computer name may contain spaces!
                if gv.savedate.get()==1:
                    outfile.write(gv._SEP+curdate)
                if gv.savetime.get()==1:
                    outfile.write(gv._SEP+curtime)
                if gv.savecomputer.get()==1:
                    outfile.write(gv._SEP+curcomputer)
                if gv.saverefresh.get()==1:
                    currefresh=refresh_.split()[1]
                    outfile.write(gv._SEP+currefresh)
                #
                for trial in gv.sub_trials[cur_subj]:
                    outfile.write(gv._SEP+"%.1f"%(trial[1]))
                if gv.savetrialorder.get()==1: # save trial order; added 12/2011
                    for trial in gv.sub_order[cur_subj]:
                        outfile.write(gv._SEP+"%i"%(trial[1]))
                outfile.write("\n")
        outfile.close()
        logmsg( "Successfully wrote "+outfilename)
        tkMessageBox.showinfo("azk2txt message","Successful completion")

        global_quit()
        ############################################################################
        ############################################################################


def global_quit():
    try: # close message file before quitting
        logfile.close()
    except: # no time to make a fuss
        pass
    root.destroy()
    root.quit()
    time.sleep(0.5)
    # need to raise explicit exit because cv_process is not 
    # a Tk widget and so it is not destroyed by killing root
    raise SystemExit


CurDir=os.getcwdu() # u for unicode; really important for Tkinter!
IconFile=os.path.join(CurDir,u"a2t.ico")
try:
    if (not os.path.exists(IconFile)): IconFile="None"
except:
    IconFile="None" # to catch any problems

root=Tk()
root.title(u"azk2txt main")
if (IconFile!="None"): root.wm_iconbitmap(IconFile)
root.withdraw()

gv=GlobVariables()
gv.reset()
cv_process=azkConvertClass()  

startw=SetupWindow(root)
#startw.startup()

root.mainloop()
