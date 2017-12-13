#
# CheckVocal.pyw v.2.3.1
# Athanassios Protopapas
# 11 August 2017 (parse new subject line produced by DMDX 5.1.5.2)
#
# This program will help with naming task data from DMDX
# It will present each recorded vocal response along with
# the correct response and will adjust RT to indicate correct,
# wrong, or timed-out response, as indicated by the user.
# (A text file with the correct response for each trial is needed.)
# The user can also check and fix improperly triggered RT measurements.
# The results are saved in a tab-separated file, one row per subject.
#
VERSION="2.3.1.0"
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
import codecs
import copy
if _WINDOWS_: import _winreg
elif _MAC_ : import plistlib
from Tkinter import *
from tkSnack import *
import tkFileDialog, tkMessageBox, tkSimpleDialog
import tkFont
import re

## FIXED PARAMETERS set in GlobVariables
##DEFAULT_SEP="\t" # separator for output data file, change to "," for CSV
DEFAULT_C_DIST=10 # pixels for vertical panel separation
DEFAULT_SUBJ_COL=25 # subjects per column in subject selection frame
DEFAULT_ONE_COL_MAX=20 # maximum number of subjects in a single column
DEFAULT_SUBJ_MAX=100 # above this number, only subject IDs are shown for selection
DEFAULT_DETRIGGER=0.5 # proportion of RT trigger RMS value signaling silence

## PARAMETERS ADJUSTED IN OPENING WINDOW
##
DEFAULT_ENCODING="Latin (ISO-8859-1)" # Encoding for answer string display
DEFAULT_RMS=45 # RMS threshold for VoX triggering, in dB
DEFAULT_WINDOWMS=10 # window for calculation of RMS energy, in ms
DEFAULT_TRIGGER=0 # use RT marks from DMDX VoX
DEFAULT_REVERSETRIGGER=False # trigger response onset (from the beginning of the file)
DEFAULT_TIMEOUT=10000 # default value for the manually set timeout GUI
DEFAULT_TIMEOUTSRC=0 # use timeout value from .rtf item file 
DEFAULT_BLINK=0 # no blink of answer string
DEFAULT_CANVASWIDTH=800 # pixels of wave & spectrogram width
DEFAULT_CANVASHEIGHT=150 # pixels of wave & spectrogram height
DEFAULT_SAVEFMT=1 # save output in one row per subject
DEFAULT_REMOVEDC=False # do not filter the sounds to remove DC offset
DEFAULT_REMOVEDUPLICATEITEMS=False # remove repeated items (adjacent trials only!) # ThP May 2014
DEFAULT_SEP=0 # separator for output data file defaults to Tab (1=space, 2=comma)
# added in v 1.7.3, optionally save date/time/PC info along with output data
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
DMDXMODE=False # True when processing azk files; False when processing abitrary audio file sets
UBPATH='~/Library/Application Support'
APPNAME="CheckVocal"
FILES_EXPNAME="CheckVocal_AudioFiles"
PLISTFNAME="param.plist"
DEFAULT_WAVEXT=u".WAV" # apparently, DMDX saves with uppercase extension and the Mac doesn't like that
FILES_WAVEXTS=[DEFAULT_WAVEXT,DEFAULT_WAVEXT.lower()]
CV_REGISTRY_KEY="SOFTWARE\\CheckVocal\\" # location of last folder key
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
                  "Chinese (GB2312)",
                  "Japanese (EUC-JP)",
                  "Korean (EUC-KR)"]
_NO_SOUND_FLAG_=u"*!*" # the "code" string in -ans to skip a response (lacking an audio file)

# Variables and settings that need to be available to various widgets and processes
#
class GlobVariables:

    def __init__(self):
        self.encoding=StringVar(root)
        self.w_canvas=StringVar(root)
        self.h_canvas=StringVar(root)
        self.rmstxt=StringVar(root)
        self.wdutxt=StringVar(root)
        self.trigchoice=IntVar(root)
        self.reversetriggerchoice=IntVar(root)
        self.removeDCchoice=IntVar(root)
        self.remove_duplicates=IntVar(root) # ThP May 2014
        self.savechoice=IntVar(root)
        self.timetxt=StringVar(root)
        self.timechoice=IntVar(root)
        self.blnkchoice=IntVar(root)
        self.sepchoice=IntVar(root)
        self.azkff=StringVar(root)
        #
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
        self.sub_origlines={}
        self.sub_ids={}
        self.sub_dates={}
        self.sub_refresh={}
        self.sub_nums={}
        self.dmdxtimeout=-1
        self.char_encoding=PRINT_ENCODING
        self._COT_=0
        #
        self.tmplistofanswers=[] # ThP Oct 2014
        self.listofanswers=[]    # content & usage changed in v.2.2.7; now fully aligned with listoftrials # ThP Oct 2014
        self.listoftrials=[]
        self.listoffiles=[]
        self.listoftrialsub=[]
        self.listofproblems=[]
        self.listofloaded=[]
        self.original_listoftrials=[]
        #
        ## Fixed parameters
        self._SEP=DEFAULT_SEP
        self._C_DIST=DEFAULT_C_DIST
        self._SUBJ_COL=DEFAULT_SUBJ_COL
        self._ONE_COL_MAX=DEFAULT_ONE_COL_MAX
        self._SUBJ_MAX=DEFAULT_SUBJ_MAX
        self._DETRIGGER=DEFAULT_DETRIGGER
        #
        ## Fonts
        self.fontfamily=DEFAULT_FONTFAMILY
        self.fontsize=DEFAULT_FONTSIZE
        self.largefontsize=DEFAULT_FONTSIZE*10/9
        self.verylargefontsize=DEFAULT_FONTSIZE*14/9
        self.mainfont=tkFont.Font(family=self.fontfamily,size=self.fontsize,weight=tkFont.NORMAL,slant=tkFont.ROMAN)
        self.mainboldfont=tkFont.Font(family=self.fontfamily,size=self.fontsize,weight=tkFont.BOLD,slant=tkFont.ROMAN)
        self.largefont=tkFont.Font(family=self.fontfamily,size=self.largefontsize,weight=tkFont.NORMAL,slant=tkFont.ROMAN)
        self.largeboldfont=tkFont.Font(family=self.fontfamily,size=self.largefontsize,weight=tkFont.BOLD,slant=tkFont.ROMAN)
        self.verylargefont=tkFont.Font(family=self.fontfamily,size=self.verylargefontsize,weight=tkFont.NORMAL,slant=tkFont.ROMAN)
        self.verylargeboldfont=tkFont.Font(family=self.fontfamily,size=self.verylargefontsize,weight=tkFont.BOLD,slant=tkFont.ROMAN)
        #

    def scale(self,w):
        return w*DEFAULT_FONTSIZE/DEFAULT_SCALE
    
    def update(self):
        if (self.timechoice.get()==0): self.timeout=self.dmdxtimeout # get from DMDX
        else: self.timeout=string.atoi(self.timetxt.get())
        tmp_enc=self.encoding.get().split()[-1]
        self.char_encoding = tmp_enc.strip('()')
##        
        self._BLINK=self.blnkchoice.get()
        self._RETRIGGER=self.trigchoice.get()
        self._REVERSETRIGGER=self.reversetriggerchoice.get()
        self._REMOVEDC=self.removeDCchoice.get()
        self._REMOVEDUPLICATES=self.remove_duplicates.get() # ThP May 2014
        self._RMSDUR=string.atof(self.wdutxt.get())
        self._RMSLIM=string.atof(self.rmstxt.get())
        self._C_WIDTH=string.atoi(self.w_canvas.get())
        self._C_HEIGHT=string.atoi(self.h_canvas.get())
        self.save_rows=self.savechoice.get()
        if (self.sepchoice.get()==0):
            self._SEP="\t"
        elif (self.sepchoice.get()==1):
            self._SEP=" "
        elif (self.sepchoice.get()==2):
            self._SEP=","

    def reset(self):
        self.encoding.set(DEFAULT_ENCODING)
        self.w_canvas.set(`DEFAULT_CANVASWIDTH`)
        self.h_canvas.set(`DEFAULT_CANVASHEIGHT`)
        self.rmstxt.set(`DEFAULT_RMS`)
        self.wdutxt.set(`DEFAULT_WINDOWMS`)
        self.trigchoice.set(DEFAULT_TRIGGER)
        self.reversetriggerchoice.set(DEFAULT_REVERSETRIGGER)
        self.removeDCchoice.set(DEFAULT_REMOVEDC)
        self.remove_duplicates.set(DEFAULT_REMOVEDUPLICATEITEMS) # ThP May 2014
        self.savechoice.set(DEFAULT_SAVEFMT)
        self.sepchoice.set(DEFAULT_SEP)
        self.timetxt.set(`DEFAULT_TIMEOUT`)
        self.timeout=DEFAULT_TIMEOUT
        self.timechoice.set(DEFAULT_TIMEOUTSRC)
        self.blnkchoice.set(DEFAULT_BLINK)
        #
        self.savedate.set(DEFAULT_SAVEDATE)
        self.savetime.set(DEFAULT_SAVETIME)
        self.savecomputer.set(DEFAULT_SAVECOMPUTER)
        self.saverefresh.set(DEFAULT_SAVEREFRESH)
        self.savetrialorder.set(DEFAULT_SAVETRIALORDER)
        #
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
            rkey=_winreg.CreateKey(_winreg.HKEY_CURRENT_USER,CV_REGISTRY_KEY) # change from _LOCAL_MACHINE as suggested by Jonathan Forster, ThP 16Jun15
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


# General message display, both on screen ("console" window) and log file
#
def logmsg(logtext):
    # added encoding to deal with printing out non-ascii path components
    try:
        gv.logfile.write(logtext.encode(gv.char_encoding, 'replace')+"\n")
        gv.logfile.flush()
    except:
        pass # in case logfile isn't working yet
    msgwindow.display(logtext+"\n")

# Fatal program failure, logging and displaying message before exiting
#
def exiterror(errtext):
    logmsg("**ERROR: "+errtext)
    tkMessageBox.showerror("CheckVocal error",errtext)
    global_quit()

# Deal with non-terminated text files by adding a final newline
#
def myreadlines(f):
    rlines=f.readlines()
    if (len(rlines)<1):
        exiterror("Empty file %s"%(f.name))
    if (rlines[-1][-1] != "\n"):
        last=rlines[-1]+"\n"
        rlines=rlines[:-1]
        rlines.append(last)
    return rlines


# Subclassing of tkSimpleDialog for the continuation prompt window
# Based on code from Chapter 10 of 'An Introduction to Tkinter' by F. Lundh
# http://www.pythonware.com/library/tkinter/introduction/dialog-windows.htm
#
class ContDialog(tkSimpleDialog.Dialog):

    def __init__(self, parent, title = None, p_time = "UNKNOWN"):

        Toplevel.__init__(self, parent)
        self.transient(parent) # OK to be transient, parent is msgwindow
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)
        if title:
            self.title(title)
        self.parent = parent
        self.result = None
        body = Frame(self)
        self.initial_focus = self.body(body, p_time)
        body.pack(padx=gv.scale(5), pady=gv.scale(5))
        self.buttonbox()
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.geometry("+%d+%d" % (gv.scale(200),gv.scale(200)))
        self.initial_focus.focus_set()
        self.wait_window(self)

    def contin(self, event=None):
        self.result=1
        if not self.validate():
            self.initial_focus.focus_set() # put focus back
            return
        self.withdraw()
        self.update_idletasks()
        self.cancel()

    def stnew(self):
        self.result=0
        if not self.validate():
            self.initial_focus.focus_set() # put focus back
            return
        self.withdraw()
        self.update_idletasks()
        self.cancel()

    def quitall(self):
        self.result=-1
        self.parent.focus_set()
        self.destroy()

    def body(self, master, p_time):

        Label(master, text="An unfinished session was found", font=gv.mainfont).pack(side="top")
        Label(master, text="last modified on "+p_time, font=gv.mainfont).pack(side="top")

    def buttonbox(self):

        buttonframe=Frame(self)
        w = Button(buttonframe,text="Continue", font=gv.mainfont,width=10,command=self.contin,default=ACTIVE).pack(side=LEFT, padx=gv.scale(5), pady=gv.scale(5))
        w = Button(buttonframe,text="Start new", font=gv.mainfont,width=10,command=self.stnew).pack(side=LEFT, padx=gv.scale(5), pady=gv.scale(5))
        w = Button(buttonframe,text="Quit", font=gv.mainfont,width=10,command=self.quitall).pack(side=LEFT, padx=gv.scale(5), pady=gv.scale(5))
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        buttonframe.pack(pady=gv.scale(10))

    def apply(self):
        pass


# The message 'console' window, with a display method to add text lines
#
class TextWindow(Toplevel):

    def __init__(self,parent):

        Toplevel.__init__(self,parent)

        self.geometry("+%d+%d" % (gv.scale(100),gv.scale(480)))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)

    def body(self):

        self.textscrollbar = Scrollbar(self)
        self.textscrollbar.pack(side=RIGHT, fill=Y)

        self.title("CheckVocal messages")
        self.textbox=Text(self)
        self.textbox.config(width=80,height=15)
        self.textbox.config(background="White",foreground="Black",font=("Courier",gv.scale(10)))
        self.textbox.config(wrap=CHAR,state=DISABLED,yscrollcommand=self.textscrollbar.set)
        self.textbox.pack(side="top")

        self.textscrollbar.config(command=self.textbox.yview)

        return self.textbox

    def display(self,text):

        self.textbox.config(state=NORMAL)
        self.textbox.insert(END,text)
        self.textbox.see(END)
        self.textbox.config(state=DISABLED)
        self.update()


# The parameter setup/configure window
#
class SetupWindow(Toplevel):

    def __init__(self,parent):
        
        Toplevel.__init__(self,parent)
        #self.transient(parent) # cannot be transient, parent is root (withdrawn)
        self.parent=parent
        self.filename_prompt="DMDX results file:"
        self.setup_window(u"CheckVocal setup")

    def setup_window(self,windowtitle):
        self.title(windowtitle)
        self.geometry("+%d+%d" % (gv.scale(100),gv.scale(150)))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)
        self.initial_focus = self
        self.initial_focus.focus_set()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.mquit)
        self.bind("<Return>", self.proceed)
        self.bind("<Escape>", self.mquit)

        self.status=-1
        self.fileOK=BooleanVar()
        self.fileOK.set(False)
        self.fileOK.trace("w",self.mayproceed)
        self.timeoutOK=BooleanVar()
        self.timeoutOK.set(True)
        self.timeoutOK.trace("w",self.mayproceed)

        self.pack_actionbuttons_row()
        self.pack_encoding_row()
        self.pack_wavedisplay_row()
        self.pack_voxtrigger_row()
        self.pack_separator_row()
        self.pack_datetime_row()
        self.pack_dataformat_row()
        self.pack_timeout_row()
        self.pack_blink_row()
        self.pack_filename_row()

    # widgets

    def pack_actionbuttons_row(self):
        self.contF0=Frame(self)
        self.cbutton0=Button(self.contF0,text="Proceed",command=self.proceed)
        self.cbutton0.config(width=12,font=gv.mainboldfont,anchor="s")
        self.cbutton0.pack(side="right",padx=gv.scale(24))
        self.cbutton01=Button(self.contF0,text="Reset",command=self.reset)
        self.cbutton01.config(width=8,font=gv.mainfont,anchor="s")
        self.cbutton01.pack(side="right",padx=gv.scale(5))
        self.cbutton02=Button(self.contF0,text="Cancel",command=self.mquit)
        self.cbutton02.config(width=8,font=gv.mainfont,anchor="s")
        self.cbutton02.pack(side="right",padx=gv.scale(5))
        self.contF0.pack(side="bottom",pady=gv.scale(10))

    def pack_encoding_row(self):
        self.encF2=Frame(self)
        self.rlabel2=Label(self.encF2,text="Character encoding:")
        self.rlabel2.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel2.pack(side="left")
        self.encodingmenu2=apply(OptionMenu,(self.encF2,gv.encoding)+tuple(ENCODING_OPTIONS)) # from http://effbot.org/tkinterbook/optionmenu.htm
        self.encodingmenu2.config(width=15,font=gv.mainfont)
        self.encodingmenu2.pack(side="left")
        self.remdup=Checkbutton(self.encF2,text="Remove consecutive repeated items",variable=gv.remove_duplicates) #
        self.remdup.config(font=gv.mainfont)
        self.remdup.pack(side="left",anchor="s",padx=35)
        self.encF2.pack(side="bottom",anchor="w")

    def pack_wavedisplay_row(self):
        self.canvasF7=Frame(self)
        self.rlabel71=Label(self.canvasF7,text="Wave display:")
        self.rlabel71.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel71.pack(side="left",anchor="e")
        self.sdim71=Spinbox(self.canvasF7,values=(200,300,400,500,600,700,800,900,1000,1100,1200))
        self.sdim71.config(width=4,font=gv.mainfont,textvariable=gv.w_canvas)
        self.sdim71.pack(side="left")
        self.rlabel72=Label(self.canvasF7,text="pixels wide by")
        self.rlabel72.config(width=12,font=gv.mainfont,anchor="w")
        self.rlabel72.pack(side="left")
        self.sdim72=Spinbox(self.canvasF7,values=(100,110,120,130,140,150,160,170,180,190,200))
        self.sdim72.config(width=4,font=gv.mainfont,textvariable=gv.h_canvas)
        self.sdim72.pack(side="left")
        self.rlabel73=Label(self.canvasF7,text="pixels tall")
        self.rlabel73.config(width=10,font=gv.mainfont,anchor="w")
        self.rlabel73.pack(side="left")
        self.canvasF7.pack(side="bottom",anchor="w",pady=gv.scale(5))

    def pack_voxtrigger_row(self):
        self.rmsF3=Frame(self)
        self.rlabel30=Label(self.rmsF3,text=" ")
        self.rlabel30.config(width=32,font=gv.mainfont,anchor="e")
        self.rlabel30.pack(side="left",anchor="e")
        self.rlabel31=Label(self.rmsF3,text="RMS threshold:")
        self.rlabel31.config(width=12,font=gv.mainfont,anchor="e")
        self.rlabel31.pack(side="left",anchor="e")
        self.rmslim31=Spinbox(self.rmsF3,from_=1,to=90)
        self.rmslim31.config(width=3,font=gv.mainfont,textvariable=gv.rmstxt)
        self.rmslim31.pack(side="left")
        self.rlabel32=Label(self.rmsF3,text="dB")
        self.rlabel32.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel32.pack(side="left")
        self.rlabel33=Label(self.rmsF3,text="Window length: ")
        self.rlabel33.config(width=14,font=gv.mainfont,anchor="e")
        self.rlabel33.pack(side="left")
        self.rmslim32=Spinbox(self.rmsF3,from_=1,to=30)
        self.rmslim32.config(width=3,font=gv.mainfont,textvariable=gv.wdutxt)
        self.rmslim32.pack(side="left")
        self.rlabel34=Label(self.rmsF3,text="ms")
        self.rlabel34.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel34.pack(side="left")
        self.rmsF3.pack(side="bottom",anchor="w")

        self.trigF4=Frame(self)
        self.rlabel4=Label(self.trigF4,text="RT marks from:")
        self.rlabel4.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel4.pack(side="left")
        self.trigchoice1=Radiobutton(self.trigF4,text="DMDX Vox",variable=gv.trigchoice,value=0,command=self.trigdmdx)
        self.trigchoice1.config(font=gv.mainfont)
        self.trigchoice1.pack(side="left",anchor="s")
        self.trigchoice2=Radiobutton(self.trigF4,text="CheckVocal",variable=gv.trigchoice,value=1,command=self.trignew)
        self.trigchoice2.config(width=9,font=gv.mainfont)
        self.trigchoice2.pack(side="left",anchor="s")
        self.filtdcC=Checkbutton(self.trigF4,text="Remove DC",variable=gv.removeDCchoice)
        self.filtdcC.config(font=gv.mainfont)
        self.filtdcC.pack(side="left",anchor="s",padx=2)
        self.revtrigC=Checkbutton(self.trigF4,text="Reverse (mark end)",variable=gv.reversetriggerchoice)
        self.revtrigC.config(font=gv.mainfont)
        self.revtrigC.pack(side="left",anchor="s",padx=2)
        self.trigF4.pack(side="bottom",anchor="w",pady=1)

    def pack_separator_row(self):
        self.sepF9=Frame(self)
        self.rlabel9=Label(self.sepF9,text="Value separator:")
        self.rlabel9.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel9.pack(side="left")
        self.sepchoice1=Radiobutton(self.sepF9,text="Tab",variable=gv.sepchoice,value=0)
        self.sepchoice1.config(font=gv.mainfont)
        self.sepchoice1.pack(side="left",anchor="s")
        self.sepchoice2=Radiobutton(self.sepF9,text="Space",variable=gv.sepchoice,value=1)
        self.sepchoice2.config(font=gv.mainfont)
        self.sepchoice2.pack(side="left",anchor="s")
        self.sepchoice3=Radiobutton(self.sepF9,text="Comma",variable=gv.sepchoice,value=2)
        self.sepchoice3.config(font=gv.mainfont)
        self.sepchoice3.pack(side="left",anchor="s")
        self.sepF9.pack(side="bottom",anchor="w",pady=1)

    def pack_datetime_row(self):
        self.saveF99=Frame(self)
        self.rlabel99=Label(self.saveF99,text="Include:")
        self.rlabel99.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel99.pack(side="left")
        self.savebox91=Checkbutton(self.saveF99,text="date",variable=gv.savedate)
        self.savebox91.config(font=gv.mainfont)
        self.savebox91.pack(side="left",anchor="s")
        self.savebox92=Checkbutton(self.saveF99,text="time",variable=gv.savetime)
        self.savebox92.config(font=gv.mainfont)
        self.savebox92.pack(side="left",anchor="s")
        self.savebox93=Checkbutton(self.saveF99,text="computer",variable=gv.savecomputer)
        self.savebox93.config(font=gv.mainfont)
        self.savebox93.pack(side="left",anchor="s")
        self.savebox94=Checkbutton(self.saveF99,text="refresh",variable=gv.saverefresh)
        self.savebox94.config(font=gv.mainfont)
        self.savebox94.pack(side="left",anchor="s")
        self.savebox95=Checkbutton(self.saveF99,text="trial order",variable=gv.savetrialorder)
        self.savebox95.config(font=gv.mainfont)
        self.savebox95.pack(side="left",anchor="s")
        self.saveF99.pack(side="bottom",anchor="w",pady=1)

    def pack_dataformat_row(self):
        self.saveF8=Frame(self)
        self.rlabel8=Label(self.saveF8,text="Save subject data in:")
        self.rlabel8.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel8.pack(side="left")
        self.savechoice1=Radiobutton(self.saveF8,text="Rows",variable=gv.savechoice,value=1,command=self.saveother)
        self.savechoice1.config(font=gv.mainfont)
        self.savechoice1.pack(side="left",anchor="s")
        self.savechoice2=Radiobutton(self.saveF8,text="Columns",variable=gv.savechoice,value=0,command=self.saveother)
        self.savechoice2.config(font=gv.mainfont)
        self.savechoice2.pack(side="left",anchor="s")
        self.savechoice3=Radiobutton(self.saveF8,text="AZK file",variable=gv.savechoice,value=-1,command=self.saveazk)
        self.savechoice3.config(font=gv.mainfont)
        self.savechoice3.pack(side="left",anchor="s")
        self.savechoice4=Radiobutton(self.saveF8,text="long format",variable=gv.savechoice,value=2,command=self.saveother)
        self.savechoice4.config(font=gv.mainfont)
        self.savechoice4.pack(side="left",anchor="s")
        self.saveF8.pack(side="bottom",anchor="w")

    def pack_timeout_row(self):
        self.timeF5=Frame(self)
        self.rlabel51=Label(self.timeF5,text="Timeout value:")
        self.rlabel51.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel51.pack(side="left")
        self.timechoice1=Radiobutton(self.timeF5,text="From DMDX item file",variable=gv.timechoice,value=0,command=self.timedmdx)
        self.timechoice1.config(font=gv.mainfont)
        self.timechoice1.pack(side="left",anchor="s")
        self.timechoice2=Radiobutton(self.timeF5,text="Set to",variable=gv.timechoice,value=1,command=self.timenew)
        self.timechoice2.config(font=gv.mainfont)
        self.timechoice2.pack(side="left",anchor="s")
        self.tentry5=Entry(self.timeF5,textvariable=gv.timetxt)#,validate="key",validatecommand=self.valtime)
        gv.timetxt.trace("w", self.timeout_edit)
        self.ptimetxt=gv.timetxt.get()
        self.tentry5.config(width=6,font=gv.mainfont)
        self.tentry5.pack(side="left")
        self.rlabel52=Label(self.timeF5,text="ms")
        self.rlabel52.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel52.pack(side="left")
        self.timeF5.pack(side="bottom",anchor="w")

    def pack_blink_row(self):
        self.blnkF6=Frame(self)
        self.rlabel6=Label(self.blnkF6,text="Blink correct response:")
        self.rlabel6.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel6.pack(side="left")
        self.blnkchoice1=Radiobutton(self.blnkF6,text="Never",variable=gv.blnkchoice,value=0)
        self.blnkchoice1.config(font=gv.mainfont)
        self.blnkchoice1.pack(side="left",anchor="s")
        self.blnkchoice2=Radiobutton(self.blnkF6,text="Always",variable=gv.blnkchoice,value=1)
        self.blnkchoice2.pack(side="left",anchor="s")
        self.blnkchoice2.config(font=gv.mainfont)
        self.blnkchoice3=Radiobutton(self.blnkF6,text="On change",variable=gv.blnkchoice,value=-1)
        self.blnkchoice3.pack(side="left",anchor="s")
        self.blnkchoice3.config(font=gv.mainfont)
        self.blnkF6.pack(side="bottom",anchor="w")

    def pack_filename_row(self):
        self.azkF1=Frame(self)
        self.rlabel1=Label(self.azkF1,text=self.filename_prompt)
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

    ## end of widget setup

    def startup(self):
        self.status=0 # startup
        self.trigdmdx()
        self.timedmdx()
        self.focus_force()
        self.get_filename()

    def inrun(self):
        self.status=1 # inrun
        # azk filename cannot be changed after startup
        self.rmessage1.unbind("<Button-1>")
        self.rlabel1.config(state=DISABLED)
        self.rlabel10.config(state=DISABLED)
        self.rmessage1.config(state=DISABLED)
        # timeout should also not be changed once started, to prevent inconsistencies
        self.rlabel51.config(state=DISABLED)
        self.rlabel52.config(state=DISABLED)
        self.timechoice1.config(state=DISABLED)
        self.timechoice2.config(state=DISABLED)
        self.tentry5.config(state=DISABLED)
        #
        self.cbutton01.config(state=DISABLED) # no reset
        self.update()

    def proceed(self,c_event=None):
        gv.update()
        if (self.status==1): # inrun
            self.parent.focus_set()
            self.destroy()
        else: # self.status==0, startup
            self.destroy()
            cv_process.run()
        #self.mquit()

    def mquit(self,c_event=None):
        if (self.status==1): # inrun
            self.parent.focus_set()
            self.destroy()
        else: # self.status==0, startup
            global_quit()
    
    def reset(self):
        gv.reset() # also resets the filename!
        self.fileOK.set(False)
        self.timedmdx()
        self.trigdmdx()

    def saveazk(self):
        self.savebox95.config(state="disabled")

    def saveother(self):
        self.savebox95.config(state="normal")

    def trigdmdx(self):
        self.revtrigC.config(state="disabled")
        ## leave other choices active to adjust on-demand triggering

    def trignew(self):
        self.rlabel31.config(state="normal")
        self.rmslim31.config(state="normal")
        self.rlabel32.config(state="normal")
        self.rlabel33.config(state="normal")
        self.rmslim32.config(state="normal")
        self.rlabel34.config(state="normal")
        self.revtrigC.config(state="normal")

    def timedmdx(self):
        self.tentry5.config(state="disabled")
        self.rlabel52.config(state="disabled")
        self.timeoutOK.set(True)

    def timenew(self):
        self.tentry5.config(state="normal")
        self.rlabel52.config(state="normal")
        self.timeout_edit()

    def timeout_edit(self, *dummy):  # based on http://effbot.org/zone/tkinter-entry-validate.htm
        strvalue = gv.timetxt.get()
        newvalue = self.timeout_validate(strvalue)
        #print "Checking timeout, got %s"%(newvalue)
        if newvalue is None:
            gv.timetxt.set(self.ptimetxt)
            if (gv.timeout>0):
                self.timeoutOK.set(True)
        elif (newvalue != strvalue):
            gv.timetxt.set(newvalue)
            self.ptimetxt=newvalue
            if newvalue:
                gv.timeout=int(newvalue)
                if (gv.timeout>0):
                    self.timeoutOK.set(True)
                else:
                    self.timeoutOK.set(False)
            else:
                gv.timeout=0
                self.timeoutOK.set(False)
        else:
            self.ptimetxt=newvalue
            if newvalue:
                gv.timeout=int(newvalue)
                if (gv.timeout>0):
                    self.timeoutOK.set(True)
                else:
                    self.timeoutOK.set(False)
            else:
                gv.timeout=0
                self.timeoutOK.set(False)
        #print "gv.timetxt='%s' gv.timeout=%i self.ptimetxt='%s'"%(gv.timetxt.get(),gv.timeout,self.ptimetxt)

    def timeout_validate(self, value):
        if len(value)>6: return None
        try:
            if value:
                v = int(value)
            return value
        except ValueError:
            return None

    def mayproceed(self, *dummy):
        if (self.fileOK.get() and self.timeoutOK.get()):
            self.cbutton0.config(state="normal")
        else:
            self.cbutton0.config(state="disabled")

    def save_expdir(self):
        if _WINDOWS_:
            try: # save selected expdir as last folder and close registry key
                rkey=_winreg.OpenKey(_winreg.HKEY_CURRENT_USER,CV_REGISTRY_KEY,0,_winreg.KEY_SET_VALUE)
                _winreg.SetValueEx(rkey,"LastFolder",0,_winreg.REG_SZ,gv.expdir)
                _winreg.CloseKey(rkey)
            except: # failed to update registry
                msgwindow.display("Failed to update registry with selected folder\n")
                # logfile not open yet, so cannot use logmsg
        elif _MAC_:
            basepath = os.path.expanduser(UBPATH)
            cvpath = os.path.join(basepath,APPNAME)
            if not os.path.exists(cvpath):
                os.mkdir(cvpath)
                msgwindow.display("System application folder created: %s\n"%(cvpath))
            pl=dict(lastfolder=gv.expdir)
            try:
                plistlib.writePlist(pl,os.path.join(cvpath,PLISTFNAME))
            except: # failed to update
                msgwindow.display("Failed to update registry with selected folder\n")
                # logfile not open yet, so cannot use logmsg
        
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
            self.fileOK.set(True)
            self.save_expdir()
        elif (len(gv.azkff.get())==0):
            self.fileOK.set(False)


class SetupWindow_Files(SetupWindow):

    def __init__(self,parent):
        
        Toplevel.__init__(self,parent)
        #self.transient(parent) # cannot be transient, parent is root (withdrawn)
        self.parent=parent
        self.filename_prompt="Audio files location:"
        self.setup_window(u"CheckVocal files setup")

    def get_filename(self,c_event=None): 
        filename=tkFileDialog.askdirectory(parent=self.parent,initialdir=gv.lastfolder,title="Choose the folder with your audio files")
        if (len(filename)>0):
            gv.expdir=filename+"/"
            gv.lastfolder=gv.expdir
            try:
                os.chdir(gv.expdir)
            except OSError:
                exiterror("Unable to work in this folder -- check security permissions")
            gv.expname=FILES_EXPNAME
            gv.azkff.set(filename)
            self.fileOK.set(True)
            self.save_expdir()
        elif (len(gv.azkff.get())==0):
            self.fileOK.set(False)

    def pack_dataformat_row(self):
        pass

    def pack_datetime_row(self):
        pass

    def pack_timeout_row(self):
        self.timeF5=Frame(self)
        self.rlabel51=Label(self.timeF5,text="Timeout value:")
        self.rlabel51.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel51.pack(side="left")
        self.tentry5=Entry(self.timeF5,textvariable=gv.timetxt)#,validate="key",validatecommand=self.valtime)
        gv.timetxt.trace("w", self.timeout_edit)
        self.ptimetxt=gv.timetxt.get()
        self.tentry5.config(width=6,font=gv.mainfont)
        self.tentry5.pack(side="left")
        self.rlabel52=Label(self.timeF5,text="ms")
        self.rlabel52.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel52.pack(side="left")
        self.timeF5.pack(side="bottom",anchor="w")

    def pack_voxtrigger_row(self):
        self.rmsF3=Frame(self)
        self.rlabel30=Label(self.rmsF3,text=" ")
        self.rlabel30.config(width=20,font=gv.mainfont,anchor="e")
        self.rlabel30.pack(side="left",anchor="e")
        self.rlabel31=Label(self.rmsF3,text="RMS threshold:")
        self.rlabel31.config(width=12,font=gv.mainfont,anchor="e")
        self.rlabel31.pack(side="left",anchor="e")
        self.rmslim31=Spinbox(self.rmsF3,from_=1,to=90)
        self.rmslim31.config(width=3,font=gv.mainfont,textvariable=gv.rmstxt)
        self.rmslim31.pack(side="left")
        self.rlabel32=Label(self.rmsF3,text="dB")
        self.rlabel32.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel32.pack(side="left")
        self.rlabel33=Label(self.rmsF3,text="Window length: ")
        self.rlabel33.config(width=14,font=gv.mainfont,anchor="e")
        self.rlabel33.pack(side="left")
        self.rmslim32=Spinbox(self.rmsF3,from_=1,to=30)
        self.rmslim32.config(width=3,font=gv.mainfont,textvariable=gv.wdutxt)
        self.rmslim32.pack(side="left")
        self.rlabel34=Label(self.rmsF3,text="ms")
        self.rlabel34.config(width=2,font=gv.mainfont,anchor="w")
        self.rlabel34.pack(side="left")
        self.rmsF3.pack(side="bottom",anchor="w")

        self.trigF4=Frame(self)
        self.rlabel4=Label(self.trigF4,text="RT marks from:")
        self.rlabel4.config(width=20,anchor="e",font=gv.mainfont)
        self.rlabel4.pack(side="left")
##        self.trigchoice1=Radiobutton(self.trigF4,text="DMDX Vox",variable=gv.trigchoice,value=0,command=self.trigdmdx)
##        self.trigchoice1.config(font=gv.mainfont)
##        self.trigchoice1.pack(side="left",anchor="s")
        self.trigchoice2=Radiobutton(self.trigF4,text="CheckVocal",variable=gv.trigchoice,value=1,command=self.trignew)
        self.trigchoice2.config(width=9,font=gv.mainfont)
        self.trigchoice2.pack(side="left",anchor="s")
        self.trigchoice2.config(state="disabled") # this should not appear changeable
        self.filtdcC=Checkbutton(self.trigF4,text="Remove DC",variable=gv.removeDCchoice)
        self.filtdcC.config(font=gv.mainfont)
        self.filtdcC.pack(side="left",anchor="s",padx=2)
        self.revtrigC=Checkbutton(self.trigF4,text="Reverse (mark end)",variable=gv.reversetriggerchoice)
        self.revtrigC.config(font=gv.mainfont)
        self.revtrigC.pack(side="left",anchor="s",padx=2)
        self.trigF4.pack(side="bottom",anchor="w",pady=1)

    def startup(self):
        self.status=0 # startup
        gv.savechoice.set(0)
        gv.trigchoice.set(1)
        gv.timechoice.set(1)
        self.trignew()
        self.timenew()
        gv.savedate.set(0)
        gv.savetime.set(0)
        self.focus_force()
        self.get_filename()
        ## need to allow selecting or disabling the -ans file somewhere

    def inrun(self):
        self.status=1 # inrun
        # azk filename cannot be changed after startup
        self.rmessage1.unbind("<Button-1>")
        self.rlabel1.config(state=DISABLED)
        self.rlabel10.config(state=DISABLED)
        self.rmessage1.config(state=DISABLED)
        # timeout should also not be changed once started, to prevent inconsistencies
        self.rlabel51.config(state=DISABLED)
        self.rlabel52.config(state=DISABLED)
        #self.timechoice1.config(state=DISABLED)
        #self.timechoice2.config(state=DISABLED)
        self.tentry5.config(state=DISABLED)
        #
        self.cbutton01.config(state=DISABLED) # no reset
        self.update()

    def reset(self):
        gv.reset() # also resets the filename!
        self.fileOK.set(False)
        self.timenew()
        self.trignew()


############################################################################
# Check audio files in a Tkinter interface using Snack sound functionality #
############################################################################
#
class CheckWaves(Toplevel):

    ## definitions of Tkinter/Snack callback functions

    def interrupt(self,event=None): # exit without saving data and without cleaning up
        self.s.stop()
        self.s.flush()
        cv_process.statusfile1.close()
        cv_process.statusfile2.close()
        self.quit()
        self.destroy()
        time.sleep(0.5) # stupid, but Tkinter may hang if parent process dies before everything is cleaned up
        
    def playleft(self,event=None): # play sound file up to the RT mark
        self.s.stop()
        self.s.play(end=int(gv.SRATE*abs(self.rt)))

    def playright(self,event=None): # play sound file from the RT mark on
        self.s.stop()
        self.s.play(start=int(gv.SRATE*abs(self.rt)))

    def goback(self,event=None): # go to the previous file
        self.s.stop()
        save_index = cv_process.current_index
        cv_process.current_index -= 1
        while ((cv_process.current_index >= 0) and
                ((gv.subj_ind[cv_process.current_index] not in gv.sub_trials.keys()) # skip removed subject trials
                 or (gv.listofanswers[cv_process.current_index]==_NO_SOUND_FLAG_)) # skip trials with no sound
               ):
            cv_process.current_index -= 1
        if (cv_process.current_index >= 0): # not reached the beginning of the list
            cv_process.current_index -= 1 # set to one before the target because advance_file() will increase by one
            cv_process.N_done -= 2
            self.advance_file()
        else:
            cv_process.current_index = save_index # do not go back if there are no previous trials
            self.gobackb.config(state=DISABLED)

    def snd_ok(self,event=None): # mark response as correct and move on to the next file
        self.s.stop()
        self.item,self.rt=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]
        if (self.rt < 0): # RT was negative to indicate wrong answer, should be changed
            gv.listofproblems.append(gv.listoffiles[cv_process.current_index])
            # Set RT to positive value to indicate correct answer
            newrt=abs(self.rt)
            gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,newrt)
            gv.listoftrials[cv_process.current_index]=(self.item,newrt)
            cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(newrt)+"\n")
            cv_process.statusfile2.flush()
            logmsg( "%s is OK, changing RT to %.1f"%(gv.listoffiles[cv_process.current_index],newrt))
        self.advance_file()

    def not_ok(self,event=None): # mark this file as an incorrect response and move on
        self.s.stop()
        gv.listofproblems.append(gv.listoffiles[cv_process.current_index])
        self.item,self.rt=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]
        # Set RT to negative value to indicate wrong answer
        newrt=-abs(self.rt) # in case the RT was already negative, e.g., from interrupted session
        gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,newrt)
        gv.listoftrials[cv_process.current_index]=(self.item,newrt)
        cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(newrt)+"\n")
        cv_process.statusfile2.flush()
        logmsg( "%s is NOT OK, changing RT to %.1f"%(gv.listoffiles[cv_process.current_index],newrt))
        self.advance_file()

    def timedout(self,event=None): # mark this file as not containing a response (timed out)
        self.s.stop()
        gv.listofproblems.append(gv.listoffiles[cv_process.current_index])
        self.item,self.rt=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]
        # Set RT to negative timeout value to indicate no response
        if (self.newlabel=="---"): # correct to not respond
            newrt=gv.timeout
        else: # failed to utter a  response when one was appropriate
            newrt=-gv.timeout
        gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,newrt)
        gv.listoftrials[cv_process.current_index]=(self.item,newrt)
        cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(newrt)+"\n")
        cv_process.statusfile2.flush()
        logmsg( "%s contains no response, changing RT to %.1f" % (gv.listoffiles[cv_process.current_index],newrt))
        self.advance_file()

    def advance_file(self,event=None): # prepare to load next file and update status
        self.hide_lines()
        self.delete_wavespec()
        cv_process.current_index += 1
        cv_process.N_done += 1
        cv_process.statusfile1.seek(0)
        cv_process.statusfile1.truncate()
        cv_process.statusfile1.write(`cv_process.current_index`+" "+`cv_process.N_done`+"\n")
        cv_process.statusfile1.flush()
        while ((cv_process.current_index < gv.totaltrials)
               and ((gv.subj_ind[cv_process.current_index] not in gv.sub_trials.keys()) # skip removed subject trials
                    or (gv.listofanswers[cv_process.current_index]==_NO_SOUND_FLAG_)) # skip trials with no sound
               ):
            cv_process.current_index += 1
        if (cv_process.current_index > 0):
            self.gobackb.config(state=NORMAL) # beyond the beginning, it is possible to go back
        else:
            self.gobackb.config(state=DISABLED) # this will happen when going back to the first item
        if (cv_process.current_index < gv.totaltrials):
            self.load_new_file()
        else:
            gv.done=1
            self.s.stop()
            self.s.flush()
            self.destroy()
            time.sleep(0.5)

    def delete_wavespec(self):
        self.c.delete(self.wave)
        self.c.delete(self.spec)
        self.c.update()

    def draw_wavespec(self):
        absrt=abs(self.rt)
        absrtis=self.ms2is(absrt)
        zlen=self.nsamples/self.zoom
        zside=zlen/2
        if (absrtis<zside):
            self.zstart=0
            self.zend=zlen
        elif (self.nsamples-absrtis < zside):
            self.zstart=self.nsamples-zlen
            self.zend=self.nsamples
        else:
            self.zstart=absrtis-zside
            self.zend=absrtis+zside
##        self.c.delete(self.wave)
##        self.c.delete(self.spec)
        self.wave=self.c.create_waveform(0, 0, sound=self.s, start=self.zstart, end=self.zend, width=gv._C_WIDTH, height=gv._C_HEIGHT, zerolevel=1)
        self.spec=self.c.create_spectrogram(0, gv._C_HEIGHT+gv._C_DIST, sound=self.s, start=self.zstart, end=self.zend, width=gv._C_WIDTH, height=gv._C_HEIGHT)
        self.zdisp.config(text=("%1d"%self.zoom))
        self.c.update()

    def hide_lines(self):
        self.c.itemconfig(self.redline1,state=HIDDEN)
        self.c.itemconfig(self.redline2,state=HIDDEN)
        self.c.itemconfig(self.grayline1,state=HIDDEN)
        self.c.itemconfig(self.grayline2,state=HIDDEN)

    def draw_lines(self):
        zlen=self.nsamples/self.zoom
        absrtis=self.ms2is(abs(self.rt))
        absorigrtis=self.ms2is(abs(self.remember_rt))
        plotrt=absrtis-self.zstart
        plotorigrt=absorigrtis-self.zstart
        self.c.coords(self.redline1,plotrt*gv._C_WIDTH/zlen,0,plotrt*gv._C_WIDTH/zlen,gv._C_HEIGHT)
        self.c.coords(self.redline2,plotrt*gv._C_WIDTH/zlen,gv._C_HEIGHT+gv._C_DIST,plotrt*gv._C_WIDTH/zlen,gv._C_DIST+2*gv._C_HEIGHT)
        self.c.coords(self.grayline1,plotorigrt*gv._C_WIDTH/zlen,0,plotorigrt*gv._C_WIDTH/zlen,gv._C_HEIGHT)
        self.c.coords(self.grayline2,plotorigrt*gv._C_WIDTH/zlen,gv._C_HEIGHT+gv._C_DIST,plotorigrt*gv._C_WIDTH/zlen,gv._C_DIST+2*gv._C_HEIGHT)
        self.c.pack()
        if (self.rt>0):
            self.c.itemconfig(self.redline1,dash=())
            self.c.itemconfig(self.redline2,dash=())
        else:
            self.c.itemconfig(self.redline1,dash=(3,1))
            self.c.itemconfig(self.redline2,dash=(3,1))
        self.c.itemconfig(self.redline1,state=NORMAL)
        self.c.itemconfig(self.redline2,state=NORMAL)
        self.c.itemconfig(self.grayline1,state=NORMAL)
        self.c.itemconfig(self.grayline2,state=NORMAL)
        self.c.lift(self.grayline1)
        self.c.lift(self.grayline2)
        self.c.lift(self.redline1)
        self.c.lift(self.redline2)
        self.c.update()

    def load_new_file(self): # load new sound file and update display
        self.s=Sound(load=gv.expdir+gv.listoffiles[cv_process.current_index])
        if (self.s["channels"]>1): # added for CheckFiles because existing audio files might contain more channels
            logmsg("Converting %s from stereo (%i channels to 1)"%(gv.listoffiles[cv_process.current_index],self.s["channels"]))
            self.s.convert(channels=1)
        if ( gv._REMOVEDC ): self.s.filter(self.DCfilt, continuedrain=0)
        self.zoom=1
        gv.SRATE=self.s.info()[1]/1000.0
        self.nsamples=self.s.length()
        self.item,self.rt=gv.listoftrials[cv_process.current_index]
        dummy,self.remember_rt=gv.original_listoftrials[cv_process.current_index] # to allow reverting
        oldlabel=self.newlabel
        self.update_answer()
        self.progress.configure(text="Item %i (%i of %i)  " %(gv.listoftrials[cv_process.current_index][0],cv_process.N_done+1,cv_process.N_todo))
        self.progress.update()
#        self.draw_wavespec() # no lines yet, may retrigger... ## actually this does not work, seems to display the old wave?
        if ( gv._RETRIGGER == 1 and gv.listoffiles[cv_process.current_index] not in gv.listofloaded):
            if (self.rt != 0): signrt=self.rt/abs(self.rt)
            else: signrt=1.0
            absrt=self.trigger(0,gv._REVERSETRIGGER)
            newrt=signrt*absrt
            gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,newrt)
            gv.listoftrials[cv_process.current_index]=(self.item,newrt)
            logmsg( "Automatically retriggering %s to %.1f" % (gv.listoffiles[cv_process.current_index],newrt))
            cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(newrt)+"\n")
            cv_process.statusfile2.flush()
            self.rt=newrt
        else:
            absrt=abs(self.rt)
        gv.listofloaded.append(gv.listoffiles[cv_process.current_index])
        self.zoomall() # use zoom-all instead of drawing wave and lines separately
                                   # this also takes care of updating the display fully
        if gv._REVERSETRIGGER: self.playleft()
        else: self.playright()
        if ( (gv._BLINK == 1) or ((gv._BLINK == -1) and (oldlabel!=self.newlabel)) ):
            self.anslabel.configure(text=self.newlabel, font=gv.verylargeboldfont, fg='yellow')
            self.anslabel.update()
            time.sleep(0.15)
            self.anslabel.configure(text=self.newlabel, font=gv.verylargeboldfont, fg='red')
            self.anslabel.update()
            time.sleep(0.15)
            self.anslabel.configure(text=self.newlabel, font=gv.verylargeboldfont, fg='blue')
            self.anslabel.update()        

    def retrigger(self,event=None,start=0):
        self.s.stop()
        self.hide_lines()
        self.c.update()
        if (self.rt != 0): signrt=self.rt/abs(self.rt)
        else: signrt=1.0
        absrt=self.trigger(start,gv._REVERSETRIGGER)
        newrt=signrt*absrt
        self.rt=newrt
        self.draw_lines()
        if gv._REVERSETRIGGER: self.playleft()
        else: self.playright()
        self.item=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]][0]
        gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,self.rt)
        gv.listoftrials[cv_process.current_index]=(self.item,self.rt)
        logmsg( "Retriggering RT for %s to %.1f" % (gv.listoffiles[cv_process.current_index],self.rt))
        cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(self.rt)+"\n")
        cv_process.statusfile2.flush()

    def next_onset(self,event=None):
        self.retrigger(start=self.rt)

    def revertRT(self,event=None):
        self.s.stop()
        if (self.rt != 0): signrt=self.rt/abs(self.rt) # retain the sign regardless of RT value
        else: signrt=1.0
        self.rt=signrt*abs(self.remember_rt)
        self.draw_lines()
        if gv._REVERSETRIGGER: self.playleft()
        else: self.playright()
        self.item=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]][0]
        gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,self.rt)
        gv.listoftrials[cv_process.current_index]=(self.item,self.rt)
        logmsg( "Reverting RT for %s to %.1f" % (gv.listoffiles[cv_process.current_index],self.rt))
        cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(self.rt)+"\n")
        cv_process.statusfile2.flush()

    def zoomin(self,event=None):
        if (self.zoom<8): # so there will be 2x, 4x, 8x
            self.zoom=2*self.zoom
            self.zoomoutb.config(state=NORMAL)
            self.zoomallb.config(state=NORMAL)
            self.hide_lines()
            self.delete_wavespec()
            self.draw_wavespec()
            self.draw_lines()
        if (self.zoom>=8):
            self.zoominb.config(state=DISABLED)
        else:
            self.zoominb.config(state=NORMAL)           

    def zoomout(self,event=None):
        if (self.zoom>1):
            self.zoom=self.zoom/2
            self.zoominb.config(state=NORMAL)
            self.hide_lines()
            self.delete_wavespec()
            self.draw_wavespec()
            self.draw_lines()
        if (self.zoom<=1):
            self.zoomoutb.config(state=DISABLED)
            self.zoomallb.config(state=DISABLED)
        else:
            self.zoomallb.config(state=NORMAL)
            self.zoomoutb.config(state=NORMAL)

    def zoomall(self,event=None,delete=True):
        self.zoom=1
        self.zoominb.config(state=NORMAL)
        self.zoomoutb.config(state=DISABLED)
        self.zoomallb.config(state=DISABLED)
        self.hide_lines()
        if delete: self.delete_wavespec()
        self.draw_wavespec()
        self.draw_lines()

    def mouseclick(self,event): # update RT with position of mouse click
        self.s.stop()
        x=event.x
        y=event.y
        if ( 0<x<gv._C_WIDTH and 0<y<(2*gv._C_HEIGHT+gv._C_DIST) ):
            if (self.rt != 0): signrt=self.rt/abs(self.rt)
            else: signrt=1.0
            absrt=round(((float(x)/gv._C_WIDTH)*(float(self.nsamples)/self.zoom)+self.zstart)/gv.SRATE,1)
            newrt=signrt*absrt
            self.rt=newrt
            self.draw_lines()
            if gv._REVERSETRIGGER: self.playleft()
            else: self.playright()
            self.item=gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]][0]
            gv.sub_trials[gv.subj_ind[cv_process.current_index]][gv.trial_ind[cv_process.current_index]]=(self.item,newrt)
            gv.listoftrials[cv_process.current_index]=(self.item,newrt)
            logmsg( "%s is mis-triggered, changing RT to %.1f" % (gv.listoffiles[cv_process.current_index],newrt))
            cv_process.statusfile2.write(`cv_process.current_index`+"\t"+`self.item`+"\t"+gv.listoffiles[cv_process.current_index].encode(gv.char_encoding,'replace')+"\t"+"%.1f"%(newrt)+"\n")
            cv_process.statusfile2.flush()

    ## end of callback definitions

    ## definitions of sound data calculation functions

    def ms2is(self,ms): # milliseconds to samples
        return (int(float(ms)*gv.SRATE))

    def is2ms(self,ns): # samples to milliseconds
        return (float(ns)/gv.SRATE)

    def trigger(self,start,reverse=False):

        # This should not be necessary but you never know... better safe than crashed
        if (self.s.length() > self.ssqr.length()):
            self.ssqr.length(self.s.length())
            self.srms.length(self.s.length())
        #
        inidur=self.ms2is(gv._RMSDUR)   # interval for RMS calculation, in samples
        tval=math.pow(10.0,gv._RMSLIM/10.0) # temporary variable
        current_trigger=tval*inidur         # sum of squares threshold
        minval=math.sqrt(tval)              # min amplitude threshold
        slen=self.s.length()
        #
        start=abs(start) # to make sure incorrect responses are properly retriggered
        nstart=self.ms2is(start)
        if reverse:        # backward trigger (offset time)
            direction=-1
            refpoint=slen
            endpoint=0
            if (nstart==0): nstart=slen-1
        else:              # regular trigger (onset time)
            direction=1
            refpoint=0
            endpoint=slen
        #
        # If not starting at the edge then we need to skip the current event
        startcheck=nstart
        if (nstart != refpoint):
            stinidur=startcheck+inidur*direction
            if ((stinidur < 0) or (stinidur > slen)):
                return(start) # not enought points to check further
            # Stuff the sliding accumulator of sums of squares with enough samples
            self.srms.sample(stinidur,0)
            for ns in range(startcheck,stinidur+direction,direction):
                ss=self.s.sample(ns)
                sss=ss*ss
                self.ssqr.sample(ns,sss)
                nval=self.srms.sample(stinidur)+sss
                self.srms.sample(stinidur,nval)
            # Calculate each new value by adding one, subtracting one, sum of squares
            # Don't calculate average, or dB (logarithm) for each sample, it's faster
            # to have the threshold converted to a sum of squares value and test that
            inv_thres=current_trigger*gv._DETRIGGER
            drop_offset=direction*(inidur+1)
            for ns in range(stinidur+direction,endpoint,direction):
                ss=self.s.sample(ns)
                sss=ss*ss
                self.ssqr.sample(ns,sss)
                nval=self.srms.sample(ns-direction)-self.ssqr.sample(ns-drop_offset)+sss
                self.srms.sample(ns,nval)
                if (nval < inv_thres): 
                    startcheck=ns
                    break
            # if we go through the whole file and fail to fall below threshold
            # then just return the original value
            if (startcheck==nstart):
                logmsg("Failed to detect silent interval past current RT for %s"%gv.listoffiles[cv_process.current_index])
                return(start)
            # else we can proceed from this new point on
            # END of setting up when not starting at the edge
        # Advance quickly to the minimum amplitude that can support threshold RMS
        # This speeds up computation greatly because it does not compute lots of
        # squares and sums over the initial low-amplitude (RT) interval
        for ns in range(startcheck,endpoint,direction):
            if (self.s.sample(ns) > minval):
                startcheck=ns
                break
        stinidur=startcheck+inidur*direction
        if ((stinidur < 0) or (stinidur > slen)):
            return(start) # not enought points to check further
        # Stuff the sliding accumulator of sums of squares with enough samples
        self.srms.sample(stinidur,0)
        for ns in range(startcheck,stinidur+direction,direction):
            ss=self.s.sample(ns)
            sss=ss*ss
            self.ssqr.sample(ns,sss)
            nval=self.srms.sample(stinidur)+sss
            self.srms.sample(stinidur,nval)
        # Calculate each new value by adding one, subtracting one, sum of squares
        # Don't calculate average, or dB (logarithm) for each sample, it's faster
        # to have the threshold converted to a sum of squares value and test that
        if (nval > current_trigger and (direction*(startcheck-refpoint) > inidur)):
            drop_offset=direction*inidur
            for ns in range(stinidur-direction,startcheck,-direction):
                ss=self.s.sample(ns-drop_offset)
                sss=ss*ss
                self.ssqr.sample(ns-drop_offset,sss)
                nval=self.srms.sample(ns+direction)-self.ssqr.sample(ns+direction)+sss
                if (nval < current_trigger):
                    return(self.is2ms(ns+direction))
                else:
                    self.srms.sample(ns,nval)
            logmsg("Failed to reverse trigger RT for %s"%gv.listoffiles[cv_process.current_index])
            return(self.is2ms(startcheck))
        else:
            drop_offset=direction*(inidur+1)
            for ns in range(stinidur+direction,endpoint,direction):
                ss=self.s.sample(ns)
                sss=ss*ss
                self.ssqr.sample(ns,sss)
                nval=self.srms.sample(ns-direction)-self.ssqr.sample(ns-drop_offset)+sss
                if (nval > current_trigger):
                    return(self.is2ms(ns))
                else:
                    self.srms.sample(ns,nval)
            logmsg("Failed to trigger RT for %s" % (gv.listoffiles[cv_process.current_index]))
        # if anything has failed in the previous sequence, just return the original point
        return(start) # except for a "next onset" request, this would be zero
            
    ## end of sound data calculation functions

    def redraw_canvas(self):
        self.c.config(height=2*gv._C_HEIGHT+gv._C_DIST, width=gv._C_WIDTH)
        self.c.coords(self.wave,0,0)
        self.c.coords(self.spec,0,gv._C_HEIGHT+gv._C_DIST)
        self.c.itemconfig(self.wave,width=gv._C_WIDTH, height=gv._C_HEIGHT)
        self.c.itemconfig(self.spec,width=gv._C_WIDTH, height=gv._C_HEIGHT)
        self.c.coords(self.grayline1,0,0,0,gv._C_HEIGHT)
        self.c.coords(self.grayline2,0,gv._C_HEIGHT+gv._C_DIST,0,gv._C_DIST+2*gv._C_HEIGHT)
        self.c.coords(self.redline1,0,0,0,gv._C_HEIGHT)
        self.c.coords(self.redline2,0,gv._C_HEIGHT+gv._C_DIST,0,gv._C_DIST+2*gv._C_HEIGHT)
        self.c.update()        

    def update_answer(self):
##        self.newlabel=gv.listofanswers[cv_process.current_index]
##        self.newlabel=unicode(gv.listofanswers[cv_process.current_index],gv.char_encoding)
        self.newlabel=gv.listofanswers[cv_process.current_index] # already in Unicode
##
        self.anslabel.configure(text=self.newlabel, font=gv.verylargeboldfont, fg='blue')
        self.anslabel.update()

    def setup(self):
        if DMDXMODE:
            setupw=SetupWindow(self)
        else:
            setupw=SetupWindow_Files(self)
        setupw.inrun()
        setupw.focus_force()
        self.wait_window(setupw)
        self.redraw_canvas()
        #self.update_answer()
        self.load_new_file() # to update sound display with possible changes (answer, DC)

    ## set up Tkinter/Snack display and load first file to create the necessary objects (widgets)

    def __init__(self,parent):

        Toplevel.__init__(self,parent)
        self.geometry("+%d+%d" % (gv.scale(100),gv.scale(50)))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)
        initializeSnack(self)
        self.title(u"Check DMDX vocal responses: Experiment %s" % (gv.expname))

        self.frame=Frame(self)
        self.frame.pack(pady=5)
        ##
        if (cv_process.current_index >= len(gv.subj_ind)):
            # this should never happen normally, only in case of a crash after all files were checked
            logmsg("No files left!\nProbably crashed at the end of previous session before saving data.")
            gv.done=1
            self.destroy()
            time.sleep(0.5)
            return None
        ##
        while ((gv.subj_ind[cv_process.current_index] not in gv.sub_trials.keys()) # skip removed subject trials
               or (gv.listofanswers[cv_process.current_index]==_NO_SOUND_FLAG_) # skip trials with no sound
               ):
            cv_process.current_index += 1
        self.s=Sound()
        self.DCfilt = Filter('iir', "-numerator", "0.99 -0.99", "-denominator", "1 -0.99")
        self.zoom=1 # scale=1 means no zoom ; zoom=2 is a magnification x2
        self.zstart=0
        ####if (_RETRIGGER_): # define the variables anyway to allow on-demand retriggering
        self.ssqr=Sound()
        self.srms=Sound()
        self.ssqr.configure(encoding="Float")
        self.srms.configure(encoding="Float")
        self.ssqr.length(2*self.ms2is(gv.timeout))
        self.srms.length(2*self.ms2is(gv.timeout))
        ####
        self.ansrow=Frame(self)
        self.anslabel=Label(self.ansrow, fg='blue', font=gv.verylargeboldfont)
        self.anslabel.pack(fill=X,expand=1,side=LEFT)
        self.ansrow.pack(fill=X,expand=1)
        self.newlabel=""
        #### keyboard shortcuts
        self.bind("c",self.snd_ok)
        self.bind("v",self.not_ok)
        self.bind("b",self.timedout)
        self.bind("<Left>",self.goback)
        self.bind("<Right>",self.advance_file)
        self.bind("z",self.zoomin)
        self.bind("x",self.zoomout)
        self.bind("a",self.zoomall)
        self.bind("<space>",self.playright)
        self.bind("g",self.playleft)
        self.bind("r",self.revertRT)
        self.bind("n",self.next_onset)
        self.bind("t",self.retrigger)
        self.bind("<Escape>",self.interrupt)
        #### Set up canvas elements here, dimensions adjusted in redraw_canvas (to allow runtime changes)
        self.c=SnackCanvas(self)
        self.c.bind("<ButtonPress>", self.mouseclick)
        self.wave=self.c.create_waveform(0, 0, sound=self.s, zerolevel=1)
        self.spec=self.c.create_spectrogram(0, 0, sound=self.s)
        self.grayline1=self.c.create_line(0,0,0,0,fill='#A89888')
        self.grayline2=self.c.create_line(0,0,0,0,fill='#C8A898')
        self.redline1=self.c.create_line(0,0,0,0,fill='red')
        self.redline2=self.c.create_line(0,0,0,0,fill='red')
        ####
        self.redraw_canvas()
        self.c.pack()
        ####
        self.buttonrow1=Frame(self)
        self.gobackb=Button(self.buttonrow1, text=' << Previous ', font=gv.largefont, fg='#000000', command=self.goback)
        if (cv_process.current_index > 0):
            self.gobackb.config(state=NORMAL)
        else:
            self.gobackb.config(state=DISABLED)
        self.gobackb.pack(side='left',padx=1)
        Button(self.buttonrow1,text=' Play left ',    font=gv.largefont, fg='#000000', command=self.playleft).pack(side='left',padx=1)
        Button(self.buttonrow1,text=' Play right ',   font=gv.largefont, fg='#000000', command=self.playright).pack(side='left',padx=1)
        Button(self.buttonrow1,text=' [[ CORRECT ]] ',font=gv.largeboldfont, fg='#118811', command=self.snd_ok).pack(side='left',padx=1)
        Button(self.buttonrow1,text=' [[* WRONG *]] ',font=gv.largeboldfont, fg='#BB1111', command=self.not_ok).pack(side='left',padx=1)
        Button(self.buttonrow1,text=' [NO RESPONSE] ',font=gv.largefont, fg='#1111AA', command=self.timedout).pack(side='left',padx=1)
        Button(self.buttonrow1,text='As is', font=gv.largefont, fg='#555555', command=self.advance_file).pack(side='left',padx=1)
        self.progress=Label(self.buttonrow1,text="Item %i (%i of %i)" %(0,0,0), font=gv.largefont)
        self.progress.pack(side='right')
        self.buttonrow1.pack(fill=X,expand=1)
        self.buttonrow2=Frame(self)
        Label(self.buttonrow2,text="Zoom:",font=gv.mainfont).pack(side=LEFT)
        self.zdisp=Label(self.buttonrow2,width=1,text=("%1d"%self.zoom),font=gv.mainfont)
        self.zdisp.pack(side=LEFT)
        self.zoominb=Button(self.buttonrow2,text="IN",width=3,font=gv.mainfont,command=self.zoomin)
        self.zoomoutb=Button(self.buttonrow2,text="OUT",width=3,font=gv.mainfont,command=self.zoomout,state=DISABLED)
        self.zoomallb=Button(self.buttonrow2,text="ALL",width=3,font=gv.mainfont,command=self.zoomall,state=DISABLED)
        self.zoominb.pack(side=LEFT,padx=1)
        self.zoomoutb.pack(side=LEFT,padx=1)
        self.zoomallb.pack(side=LEFT,padx=1)
        Label(self.buttonrow2, text=" ", font=gv.mainfont,width=10).pack(side=LEFT)
        Button(self.buttonrow2, text='Interrupt', font=gv.mainfont, width=13, fg='#000000', command=self.interrupt).pack(side=RIGHT)
        Button(self.buttonrow2, text='Setup', font=gv.mainfont, width=7, fg='#000000', command=self.setup).pack(side=RIGHT,padx=2)
        Label(self.buttonrow2, text=" ", font=gv.mainfont).pack(side=RIGHT,padx=10,expand=1)
        Button(self.buttonrow2,text="Revert RT",font=gv.mainfont,width=9,command=self.revertRT).pack(side=RIGHT,padx=1)
        Button(self.buttonrow2,text="Next onset",font=gv.mainfont,width=9,command=self.next_onset).pack(side=RIGHT,padx=1)
        Button(self.buttonrow2,text="Retrigger",font=gv.mainfont,width=9,command=self.retrigger).pack(side=RIGHT,padx=1)
        self.buttonrow2.pack(pady=4,fill=X,expand=1)
        self.load_new_file()


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
        self.geometry("+%d+%d" % (gv.scale(100),gv.scale(175)))
        if (IconFile!="None"): self.wm_iconbitmap(IconFile)

        self.sublabel=Label(self, text="Select subjects to process data", font=gv.mainboldfont)
        self.sublabel.pack(padx=gv.scale(30),pady=gv.scale(5))
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
        ##                                       ##
        ##for subject in gv.sub_trials.keys():   ## corrected to present subjects
        live_subjects=gv.sub_ids.keys()          ## in order of appearance in azk
        for subjnum in live_subjects:            ##
            subject=gv.sub_ids[subjnum]          ##
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
##        self.subbutton=Button(self,text="OK",width=20,font=gv.mainboldfont,command=self.subj_select_ok)
##        self.subbutton.pack(pady=10,side='top')


# This is where all of the preparation and track-keeping is done
#
class CheckVocalClass:

    def __init__(self):
        # we need the instance created so that it is globally available
        pass

    def get_rtf_timeout(self):
        self.timeout = -1
        rtfitemfilename=gv.expname+".rtf"
        try:
            rtfitemfile=open(gv.expdir+rtfitemfilename)
        except:
            exiterror("Could not open item file %s to determine timeout" % rtfitemfilename)
        rtfitems=rtfitemfile.read()
        for wschar in "\n\t\r": # turn all white space into regular space characters
            rtfitems=rtfitems.replace(wschar,' ')
        while (rtfitems.find('\\')>=0): # get rid of rtf tags
            backslashpos=rtfitems.find('\\')
            nextspacepos=rtfitems[backslashpos:].find(' ') # removing up to a space gets rid of successive tags
            if (nextspacepos<1):
                rtfitems=rtfitems[:backslashpos]
            else:
                rtfitems=rtfitems[:backslashpos]+rtfitems[backslashpos+nextspacepos+1:]
        rtfitemfile.close()
        linepieces=string.split(rtfitems,'<')
        for parstr in linepieces:
            lstr=string.lower(parstr) # to deal with arbitrary capitalization
            if ((len(lstr)>2 and lstr[:2]=="t ") or (len(lstr)>8 and lstr[:8]=="timeout ")):
                try:
                    self.timeout=string.atof(lstr[lstr.find(' '):lstr.find('>')])
                    logmsg("Setting timeout to %.1f" % (self.timeout))
                    break
                except:
                    logmsg("Failed to parse parameters properly: %s" % (lstr))
                    exiterror("Error in determining timeout from file %s" % (rtfitemfilename))

        if (self.timeout<0): # did not manage to figure out timeout
            exiterror("Could not determine timeout from file %s" % (rtfitemfilename))
        else:
            gv.dmdxtimeout=self.timeout # save it for future setup updates
            gv.timeout=gv.dmdxtimeout

    def process_azk(self):

        gv.azkfilename=gv.expname+".azk"
        try:
            azkfile=open(gv.expdir+gv.azkfilename)
        except:
            exiterror("Could not open %s to read experiment data" % (gv.azkfilename))

        logmsg("Working folder is %s" % gv.expdir)
        azklines=myreadlines(azkfile)
        logmsg("Opened %s" % gv.azkfilename)

        Nlines=len(azklines)
        subjno=0
        self.ntrials=0
        line=0

        self.subject_select={}

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
            # Do not proceed if there are no lines left
            if (line >= len(azklines)):
                logmsg("File ended at line %i before processing %i subjects" % (line,gv.Nsubj))
                #sys.exit(-1) # do not die on premature file end
                # perhaps data have been manually removed
                subj=gv.Nsubj
                break
            #
            while (len(azklines[line])<11 or azklines[line][:10] != "**********"):
                line += 1
                if (line >= len(azklines)):
                    logmsg("Unexpected end of file at line %i before processing %i subjects" % (line,gv.Nsubj))
                    #sys.exit(-1) # do not die on premature file end
                    # perhaps data have been manually removed
                    subj=gv.Nsubj
                    break


            slineparts = string.split(azklines[line+1].decode(gv.char_encoding).rstrip(),',') # discard newline
            # ThP 20130107 added .decode() to allow for possible non-latin data -- not just subject ID but also PC ID
            try:
                # ThP 20170811 recoded parsing of this line to read the first two and last field, due to added intervening fields by DMDX 5.1.5.2
                subj_,date_,refresh_ = slineparts[:3] # ignore DMDX/Windows versions
                ids_ = slineparts[-1]
            except ValueError:
                logmsg("Insufficient fields for subject %i (Subject line messed up)"%(subjno+1))
                # useless data without ID, try to skip this subject
                # assume there will always be subject, date, and refresh fields
                subj_,date_,refresh_ = string.split(azklines[line+1].rstrip(),',') # discard newline
                ids_=u"xxx xxx" # so that a dummy ID will be made up below
            s_ = string.atoi(string.split(subj_)[1])
            if (subj_[:7]!="Subject" or subjno+1 != s_):
                logmsg("Unexpected subject number "+`s_`+" (expected "+`subjno+1`+") at line "+`line`)
                #sys.exit(-1) # do not die on unexpected number,
                # perhaps a subject's data have been manually removed from the .azk
            subjno += 1

            #ids_=unicode(ids_,gv.char_encoding) # possible non-latin ID # ThP commented out 20130107 v2.2.4 b/c entire line now unicoded above
            splitOK=False
            try:
                splitid=string.split(ids_)[1:]
                splitOK=True
            except IndexError: # could not split into two?
                if (len(ids_)>3 and ids_[1:3]==u"ID"):
                    s_id = ids_[3:].strip()
                else:
                    s_id = u""
            if splitOK:
                if (len(splitid)>1): # ID with space in it
                    s_id = u" ".join(splitid) # put the space back in
                else: # normal case
                    s_id = splitid[0] # [0] because a list was returned by split

            if (ids_[1:3]!="ID" or len(s_id)<1):
                s_id = u"SID%04i" % s_
                logmsg("Could not determine ID for subject %i at line %i, will use %s" % (s_,line+1,s_id))
                # sys.exit(-1) # do not die on undefined ID, create one from the subject number
                self.subject_select[s_id]=0 # deactivate subject because wav filenames need a real ID 
            else:
                self.subject_select[s_id]=1 

            for s_temp in gv.sub_ids.keys():
                if (s_id == gv.sub_ids[s_temp]):
                    logmsg( "Duplicate subject ID %s (subjects %i and %i)" % (s_id,gv.sub_nums[s_id],s_))
                    # Deal with duplicate IDs, first in v. 1.6.0
                    dup_id=s_id+'*'
                    dup_num=gv.sub_nums[s_id]
                    gv.sub_dates[dup_id]=gv.sub_dates[s_id]
                    gv.sub_refresh[dup_id]=gv.sub_refresh[s_id]
                    gv.sub_nums[dup_id]=dup_num
                    gv.sub_ids[dup_num]=dup_id
                    self.subject_select[dup_id]=0
                    #logmsg( "CERTAIN ID-BASED CHECKS MAY FAIL; BE SURE TO CHECK OUTPUT!" )
                    #sys.exit(-1) # Do not die on duplicate IDs
                    #they probably mean a repeated run (and overwritten wavs)
                    #dictionary values will be replaced (as they should), hopefully without side-effects

            logmsg( "Subject "+`subjno`+", ID="+s_id)
            gv.sub_ids[subjno]=s_id
            gv.sub_dates[s_id]=date_
            gv.sub_refresh[s_id]=refresh_
            gv.sub_nums[s_id]=subjno

            RTheaders = string.split(azklines[line+2].rstrip())
            if (len(RTheaders)>2): # Recognize a 3rd column of COT data (clock-on-trial)
                if (RTheaders[2] == "COT"):
                    if (gv._COT_ == 0):
                        logmsg( "COT header detected")
                    gv._COT_ = 1
                else:
                    logmsg( "Unknown header identifier at line "+`line+2`)
                    self.subject_select[s_id]=0  # do not process subjects with not understood data
                    
            if (RTheaders[:2] != ["Item","RT"]):
                logmsg( "Item/RT identifier not found at line "+`line+2`)
                self.subject_select[s_id]=0  # definitely kill the subject with unparseable data
                #sys.exit(-1) # Do not die on problematic item/RT data;
                #This is a problem for the subject's data but perhaps the rest of the file is OK

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

##            if (self.ntrials > 0): # skip check for first subject
##                if (endline-startline-exlines != self.ntrials):
##                    logmsg( "Mismatching number of trials (expected %i, found %i) for subject %i (ID %s)" % (self.ntrials, endline-startline-exlines, subjno, s_id))
##                    self.subject_select[s_id]=0  # definitely kill the subject with unparseable data
##                    #sys.exit(-1) # do not die on poor data, perhaps more good subject data follow
##            else:
##                self.ntrials = endline-startline-exlines

            tmptrials=[]
            itrial=0
            nirem=0 # number of removed repeated items # ThP May 2014
            previous_line=" "
            previous_item= -1
            for trial in azklines[startline:endline]:
                # catch a case of split error-report lines on non-ASCII strings ## thp 2006-10-13
                if (not (previous_line[0]=='!' and trial[0] != '!' and 
                         string.count(previous_line,'"')>0 and string.count(trial,'"')>0)):
                    if (trial[0] != '!'):
                        if (gv._COT_ == 1): # parse correctly data even if a 3rd (COT) column is present
                            item_,rt_,cot_ = string.split(trial)[:3] # COT data will be ignored
                        else:
                            item_,rt_ = string.split(trial)[:2] # need to discard possible ABORT tags
                        item=string.atoi(item_)
                        rt=string.atof(rt_)
                        ## optionally remove repeated items ## ThP May 2014
                        if (item==previous_item and gv._REMOVEDUPLICATES):
                            junk_item = tmptrials.pop()
                            logmsg("Removing repeated item %i"%(item))
                            nirem += 1
                            exlines += 1
                        else:
                            itrial += 1 # increase trial counter to store order
                        previous_item = item
                        ## end of item cancellation
                        tmptrials.append((item,rt,itrial))
                previous_line=trial # retain for comparison
            tmptrials.sort(key=operator.itemgetter(0)) # don't sort on RTs and order
            tmptimes=[(x[0],x[1]) for x in tmptrials]
            tmporder=[(x[0],x[2]) for x in tmptrials]
            # this is the list of trial data for the subject, by increasing item number
            gv.sub_trials[s_id]=tmptimes
            # this is the order of trial presentation
            gv.sub_order[s_id]=tmporder
            # also save the original data lines in case we need to save in .azk
            gv.sub_origlines[s_id]=azklines[startline:endline]
            if ( gv._REMOVEDUPLICATES and nirem>0 ):                    # ThP May 2014
                logmsg( "Removed %i repetitions in total" % (nirem) )   # ThP May 2014

            ### Moved from above in 2.2.5, to check number of lines after removal of repeated items
            if (self.ntrials > 0): # skip check for first subject
                if (endline-startline-exlines != self.ntrials):
                    logmsg( "Mismatching number of trials (expected %i, found %i) for subject %i (ID %s)" % (self.ntrials, endline-startline-exlines, subjno, s_id))
                    self.subject_select[s_id]=0  # definitely kill the subject with unparseable data
                    #sys.exit(-1) # do not die on poor data, perhaps more good subject data follow
            else:
                self.ntrials = endline-startline-exlines

        logmsg( "End of processing trials in %s, %i lines left" % (gv.azkfilename,Nlines-line))

    def process_ans(self):
        ansfilename=gv.expdir+gv.expname+"-ans.txt"
        try: # default location for answers same as rtf/azk/wav
##            ansfile=open(ansfilename)
            ansfile=codecs.open(ansfilename,"r",gv.char_encoding) # open straight into Unicode
##            
        except:
            ansfilename=gv.expname+"-ans.txt"
            try: # see if they are in the current directory instead
##                ansfile=open(ansfilename)
                ansfile=codecs.open(ansfilename,"r",gv.char_encoding) # open straight into Unicode
            except:
                exiterror( "Could not open %s to read correct answers!" % (ansfilename))
        #
        # We need to make sure the answers match precisely the experimental items
        #
        answers=myreadlines(ansfile)
        if ( len(answers) != self.ntrials):
            exiterror( "Number of lines (%i) in %s does not match expected number of trials (%i)" % (len(answers),ansfilename,self.ntrials))
        i=0
        for line in answers:
            i += 1
            if ( len(line)<2 ):
                exiterror ( "Empty line (%i) in answers file!" % (i) )
            try:
                item,answer=string.split(line.rstrip(),maxsplit=1)
            except:
                exiterror( "Error reading answers file\n(problem line %i, contents %s)" % (i,line.rstrip()) )
        i=0
        try:
            answers.sort(key=lambda(s):int(s.split()[0]))
        except:
            exiterror ("Error sorting answer list!\nPerhaps an item number is not a number?")
        ref_subj=gv.sub_trials.keys()[0] # For reporting ans-inconsistencies only; might cause problems 
                                      # if it refers to a subject not intended to be processed.
                                      # Perhaps should be be moved past subject selection?
        for line in answers:
            item,answer=string.split(line.rstrip(),maxsplit=1)
            subjitem=gv.sub_trials[gv.sub_trials.keys()[0]][i][0]
            if (string.atoi(item) != subjitem):
                exiterror( "Item number %i (position %i) in asnwers file\n%s\ndoes not match item number %i in reference subject %i (ID %s)" % (string.atoi(item),i,ansfilename,subjitem,gv.sub_nums[ref_subj],ref_subj))
            # tmplistofanswers introduced in v.2.2.7 to enforce listing of all answers for all subjects
            # because assuming ntrials in the listoftrials by all subjects failed when subjects had skipped
            # trials before missing files in which case the skipping desynchronized trials and answers
            gv.tmplistofanswers.append(answer) # take care of encoding during display, to allow runtime changes; # ThP Oct 2014
            i += 1
            # Try to allow trials without sound/spoken responses, use "*!*" ans as a no-answer flag
            if (answer==_NO_SOUND_FLAG_):
                logmsg( "No spoken response for item number %i (position %i)"%(int(item),i) )
            #
        ansfile.close()

    def verify_audiofiles(self):
        logmsg( "Verifying audio files...")

        for cur_subj in gv.sub_trials.keys():
            sub_trial_ind=0
            for trial in gv.sub_trials[cur_subj]:
                item,rt=trial
                answer=gv.tmplistofanswers[sub_trial_ind]
                audiofilename=gv.expname+cur_subj+`item`+DEFAULT_WAVEXT
                if (answer!=_NO_SOUND_FLAG_): # do not check trials flagged as lacking spoken responses
                    #####################################################
                    # Make sure all audio files for this subject exist
                    try:
                        testfile=open(gv.expdir+audiofilename)
                        testfile.close()
                    except IOError:
                        logmsg( "Could not open file %s, removing subject %s" % (audiofilename,cur_subj))
                        #sys.exit (-2) # don't die, just remove the subject
                        self.subject_select[cur_subj]=0
                        break # no need to continue checking this subject's files
                    ######################################################
                gv.listoftrials.append(trial)
                gv.listoffiles.append(audiofilename)
                gv.listoftrialsub.append(cur_subj)
                gv.listofanswers.append(answer)
                gv.trial_ind[gv.totaltrials]=sub_trial_ind
                gv.subj_ind[gv.totaltrials]=cur_subj
                sub_trial_ind += 1
                gv.totaltrials += 1

    def check_saved_status(self):
        self.statusfilename1="##"+gv.expname+"##"
        self.statusfilename2=gv.expname+"-chg.txt"
        self.statusfilename3=gv.expname+"-sub.txt"
        continue_old=0

        # Check that current directory is writeable and prepare status saving
        try:
            self.statusfile1=open(self.statusfilename1,"r")
            status=myreadlines(self.statusfile1)
            saved_session=1
        except IOError: # non-existent or unreadable saved session
            saved_session=0

        if (saved_session==1):
            previoustime=time.ctime(os.path.getmtime(self.statusfilename1))
            # ans is transient on msgwindow, not root, because root is hidden (withdrawn)
            ans=ContDialog(msgwindow,"Continue previous session?",previoustime).result

            if (ans==-1): # QUIT
                global_quit()
                # if we don't return, cv_process will continue to run past this if,
                # because it is not a widget and so it is not killed when root is destroyed!
                return -1

            elif (ans==1): # to continue session, we must parse the status
                continue_old=1
                self.current_index, self.N_done = map(lambda x: string.atoi(x),string.split(status[0]))
                logmsg( "Restarting at index %i" % (self.current_index))
                for filenum in range(self.current_index):              # recreate list of loaded to avoid re-triggering
                        gv.listofloaded.append(gv.listoffiles[filenum]) # not checked with multiple/deselected subjects!!
                self.statusfile1.close()
                try:
                    self.statusfile2=open(self.statusfilename2,"r")
                except:
                    exiterror ("Could not open status file %s to read previous changes" % (self.statusfilename2))
                status=myreadlines(self.statusfile2)
                for line in status: # recreate stated changes
                    splitline=string.split(line.rstrip())
                    savedindex=string.atoi(splitline[0])
                    saveditem=string.atoi(splitline[1])
                    savedfile=' '.join(splitline[2:-1]).decode(gv.char_encoding) # was [2] but filename may contain spaces!
                    savedrt=string.atof(splitline[-1]) # was 3 ; deals with spaces in filenames ### March 2013
                    if (gv.listoffiles[savedindex] != savedfile):
                        logmsg( "Saved filename %s at saved index %i not matching filename %s" % (savedfile, savedindex, gv.listoffiles[savedindex]))
                        exiterror( "Mismatching status information")
                    else:
                        gv.sub_trials[gv.subj_ind[savedindex]][gv.trial_ind[savedindex]]=(saveditem,savedrt)
                        gv.listoftrials[savedindex]=(saveditem,savedrt)
                        logmsg( "Repeating change of index %i (trial %i) RT to %.1f" % (savedindex,saveditem,savedrt))
                self.statusfile2.close()
                try: # you never know when someone will try to process a folder on a CD...
                    self.statusfile1=open(self.statusfilename1,"w")
                    self.statusfile1.write(`self.current_index`+" "+`self.N_done`+"\n")
                    self.statusfile1.flush()
                    self.statusfile2=open(self.statusfilename2,"a")
                except IOError:
                    logmsg( "Unable to open status files %s and %s for writing" % (self.statusfilename1, self.statusfilename2))
                    exiterror( "Could not write status files" )
                try:
                    self.statusfile3=open(self.statusfilename3,"r")
                except:
                    exiterror ("Could not open status file %s to read subject selection" % (self.statusfilename3))                
                statuslines3=myreadlines(self.statusfile3)
                for statusline in statuslines3:
                    try:
                        action, subject = string.split(statusline.decode(gv.char_encoding))
                    except:
                        exiterror ("Error parsing status file %s\nOffending line: %s" % (self.statusfilename3,statusline)) 
                    if (action=="REMOVE"):
                        actioncode=0
                    elif (action=="RETAIN"):
                        actioncode=1
                    else:
                        exiterror( "Unrecognized subject action in %s" % (self.statusfilename3))
                    try:
                        currentcode=self.subject_select[subject]
                    except:
                        exiterror( "Failed to determine current code for subject %s" % (subject))
                    if (actioncode==0): # previously deselected
                        self.subject_select[subject]=0 # just deselect
                    else: # actioncode=1
                        if (currentcode==0):
                            exiterror( "Subject %s should be removed but saved status is RETAIN" % (subject))
                        else:
                            self.subject_select[subject]=-1 # a sign that subject is checked and retained
                self.statusfile3.close()
                for subject in gv.sub_trials.keys():
                    if (self.subject_select[subject] == 1): # subject to retain not saved in status file
                        exiterror( "Did not find retain status for subject %s in %s" % (subject,self.statusfilename3))
                    elif (self.subject_select[subject] == -1):
                        self.subject_select[subject]=1
                    else: # this better be zero
                        if (self.subject_select[subject] != 0):
                            exiterror( "Error in subject retain/remove selection")
                            
            else: # ans=0, start over
                continue_old=0
        return (continue_old)

    def select_subjects(self):
        subselect=SubjectSelect(root)
        subselect.focus_force()
        subselect.wait_window(subselect)
        for subject in gv.sub_trials.keys():
            self.subject_select[subject]=subselect.sub_buttons[subject].get()

    def start_new_session(self):
        self.current_index=0
        self.N_done=0
        try: # make sure we are at a writable directory before proceeding
            self.statusfile1=open(self.statusfilename1,"w")
            self.statusfile1.write(`self.current_index`+" "+`self.N_done`+"\n")
            self.statusfile1.flush()
            self.statusfile2=open(self.statusfilename2,"w")
            self.statusfile3=open(self.statusfilename3,"w")
        except IOError: # problem!
            logmsg( "Unable to write status files." )
            logmsg( "Make sure the folder you want to work in is not read-only!" )
            exiterror( "Could not write status file")

        self.select_subjects()

        for subject in gv.sub_trials.keys():
            if (not self.subject_select[subject]):
                self.statusfile3.write("REMOVE "+subject.encode(gv.char_encoding,'replace')+"\n")
            else:
                self.statusfile3.write("RETAIN "+subject.encode(gv.char_encoding,'replace')+"\n")

        self.statusfile3.close()

    def verify_output_filename(self,outfilename):
        outfileOK=0
        try: # file should not exist, so open for read should fail
            outfile=open(outfilename)
            outfile.close()
        except IOError:
            outfileOK=1
            outfile=open(outfilename,"w") # hopefully this will work
        while (outfileOK<1): # keep requesting an output file name until valid
            if (outfileOK==0):
                ans=tkMessageBox.askyesno("File overwrite",
                                          ("Output file %s exists, overwrite? (y/n)" % (outfilename)))
            else: # -1 means we are looping (probably after a click on "cancel")
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
        return(outfile)

    def save_output(self):
        #
        if (gv.save_rows==-1): # if AZK format output is requested, name the file accordingly
            outfilename=gv.expname+"-datalist.azk"
        else:
            outfilename=gv.expname+"-datalist.txt"
        outfile=self.verify_output_filename(outfilename)

        # make a new list from sub_ids.keys() instead of sub_trials.keys()
        # in order to save the data in the original subject order (as in azk)
        live_subjects=[]
        for subjnum in gv.sub_ids.keys(): # gv.sub_trials.keys():
            if (self.subject_select[gv.sub_ids[subjnum]]):
                live_subjects.append(subjnum)
        ref_subj=gv.sub_trials.keys()[0]
        
        if (gv.save_rows==-1): # save subject data in AZK file
            outfile.write("\nSubjects incorporated to date: %03d\n"%len(gv.sub_trials.keys()))
            outfile.write("Data file started on machine CheckVocal\n")
            subjno=0
            # Need to re-sort by sub_origlines[s_id] (which was set to azklines[startline:endline])
            # and add COT if present in the original 
            # for cur_subj in gv.sub_trials.keys():
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids[subjnum]
                subjno+=1
                sub_newlines={}
                outfile.write("\n**********************************************************************\n")
                outfile.write("Subject %d,%s,%s, ID %s\n" % (subjno,gv.sub_dates[cur_subj],gv.sub_refresh[cur_subj],cur_subj.encode(gv.char_encoding,'replace')))
                if (gv._COT_):
                    outfile.write("  Item       RT       COT\n")
                else:
                    outfile.write("  Item       RT\n")
                for trial in gv.sub_trials[ref_subj]:
                    sub_newlines[`trial[0]`]=("%6d  %8.2f\n" % (trial[0],gv.sub_trials[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                for trial_line in gv.sub_origlines[cur_subj]:
                    if (trial_line[0]=="!"):
                        outfile.write(trial_line)
                    else:
                        trialitem=string.split(trial_line)[0]
                        if (gv._COT_):
                            cotstr=" "+("%9.2f"%(string.atof(string.split(trial_line)[2])))
                        else:
                            cotstr=""
                        # need to remove final \n from trial line so it can be re-appended after optional COT
                        outfile.write(sub_newlines[trialitem][:-1]+cotstr+"\n")
        ##
        elif (gv.save_rows==0): # save subject data in columns
            outfile.write("item")
            #for cur_subj in gv.sub_trials.keys():
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids[subjnum]
                outfile.write(gv._SEP+cur_subj.encode(gv.char_encoding,'replace'))
            outfile.write("\n")
            # Optionally, save date/time/PC info 
            if gv.savedate.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum] # not sub_ids_new! (from azk2txt)
                    timedatestr=gv.sub_dates[cur_subj]
                    curdate=timedatestr.split()[0]
                    outfile.write(gv._SEP+curdate.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.savetime.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum]
                    timedatestr=gv.sub_dates[cur_subj]
                    curtime=timedatestr.split()[1]
                    outfile.write(gv._SEP+curtime.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.savecomputer.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum]
                    timedatestr=gv.sub_dates[cur_subj]
                    curcomputer=timedatestr.split(None,3)[3] # sep=None, maxsplit=3; computer name may contain spaces!
                    outfile.write(gv._SEP+curcomputer.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            if gv.saverefresh.get()==1:
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum]
                    refreshstr=gv.sub_refresh[cur_subj]
                    currefresh=refreshstr.split()[1]
                    if (len(currefresh)>2 and currefresh[-2:]=="ms"): currefresh=currefresh[:-2]
                    outfile.write(gv._SEP+currefresh.encode(gv.char_encoding)) # in case of non-latin IDs
                outfile.write("\n")
            ##
            for trial in gv.sub_trials[ref_subj]:
                outfile.write(`trial[0]`)
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum]
                    outfile.write(gv._SEP+"%.1f"%(gv.sub_trials[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                outfile.write("\n")
            ## save trial order after RT; added 12/2011
            if gv.savetrialorder.get()==1:
                for trial in gv.sub_trials[ref_subj]:
                    outfile.write("ord"+`trial[0]`)
                    for subjnum in live_subjects:
                        cur_subj=gv.sub_ids[subjnum]
                        outfile.write(gv._SEP+"%i"%(gv.sub_order[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                    outfile.write("\n")
                
        ##
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
                cur_subj=gv.sub_ids[subjnum]
                curdate,curtime,on_,curcomputer=gv.sub_dates[cur_subj].split(None,3)  # sep=None, maxsplit=3; computer name may contain spaces!
                refreshstr=gv.sub_refresh[cur_subj]
                for trial in gv.sub_trials[cur_subj]:
                    outfile.write(cur_subj.encode(gv.char_encoding)) # in case of non-latin IDs
                    if gv.savedate.get()==1:
                        outfile.write(gv._SEP+curdate)
                    if gv.savetime.get()==1:
                        outfile.write(gv._SEP+curtime)
                    if gv.savecomputer.get()==1:
                        outfile.write(gv._SEP+curcomputer)
                    if gv.saverefresh.get()==1:
                        currefresh=refreshstr.split()[1]
                        if (len(currefresh)>2 and currefresh[-2:]=="ms"): currefresh=currefresh[:-2]
                        outfile.write(gv._SEP+currefresh)
                    if gv.savetrialorder.get()==1: # save trial order; added 12/2011
                        outfile.write(gv._SEP+"%i"%(gv.sub_order[cur_subj][gv.sub_trials[cur_subj].index(trial)][1]))
                    outfile.write((gv._SEP+"%i"+gv._SEP+"%.1f")%(trial[0],trial[1]))
                    outfile.write("\n")
        ##
        else: # normally save_rows=1, save data in rows; but set as default just in case
            outfile.write("subject")
            # Optionally, save date/time/PC info 
            if gv.savedate.get()==1:
                outfile.write(gv._SEP+"date")
            if gv.savetime.get()==1:
                outfile.write(gv._SEP+"time")
            if gv.savecomputer.get()==1:
                outfile.write(gv._SEP+"computer")
            if gv.saverefresh.get()==1:
                outfile.write(gv._SEP+"refresh")
            ##
            for trial in gv.sub_trials[ref_subj]:
                outfile.write(gv._SEP+`trial[0]`)
            ## Optionally, save trial order after RT; added 12/2011
            if gv.savetrialorder.get()==1: 
                for trial in gv.sub_order[ref_subj]:
                    outfile.write(gv._SEP+"ord"+`trial[0]`)
            outfile.write("\n")
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids[subjnum]
                outfile.write(cur_subj.encode(gv.char_encoding,'replace'))
                #
                curdate,curtime,on_,curcomputer=gv.sub_dates[cur_subj].split(None,3)  # sep=None, maxsplit=3; computer name may contain spaces!
                refreshstr=gv.sub_refresh[cur_subj]
                if gv.savedate.get()==1:
                    outfile.write(gv._SEP+curdate)
                if gv.savetime.get()==1:
                    outfile.write(gv._SEP+curtime)
                if gv.savecomputer.get()==1:
                    outfile.write(gv._SEP+curcomputer)
                if gv.saverefresh.get()==1:
                    currefresh=refreshstr.split()[1]
                    if (len(currefresh)>2 and currefresh[-2:]=="ms"): currefresh=currefresh[:-2]
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

        logmsg(("CheckVocal version %s (%s)"%(VERSION,mtime)))
        logmsg("Session started : "+str(datetime.datetime.now()))
        logmsg("Processing experiment "+gv.expname)

        # Processing of .rtf item file to determine timeout value
        # We need this value for "no-response" items (in DMDX fashion, RT=-timeout)
        #
        if (gv.timeout<0): # no timeout command-line option, must determine from item file parameter
            self.get_rtf_timeout()

        self.process_azk() # Read in trial and RT information from DMDX output file
        self.process_ans() # Read in correct answers from the special "-ans" file     
        self.verify_audiofiles() # Prepare lists of items and verify audio files

        # no point in moving on if there are no valid subject data
        valid_subjects=0
        for subject in gv.sub_trials.keys():
            if (self.subject_select[subject]==1): valid_subjects += 1
        if (valid_subjects==0):
            exiterror( "No valid subject data to process!")            

        # save a copy of the unaltered list of trials to be able to revert to
        gv.original_listoftrials=gv.listoftrials[:]
        logmsg( "Audio files OK, verifying status...")

        continue_old=self.check_saved_status() # Restore previous status, if detected

        # no point in moving on if there are no valid subject data
        valid_subjects=0
        for subject in gv.sub_trials.keys():
            if (self.subject_select[subject]==1): valid_subjects += 1
        if (valid_subjects==0):
            exiterror( "No valid subject data to process!")            

        if (continue_old == 0): # starting a new session from scratch
            self.start_new_session() # set up status files and select subjects

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

        logmsg( "Check audio in GUI")
        # prepare counters
        gv.done=0
        gv.SRATE=0
        self.N_todo=len(gv.sub_trials.keys())* self.ntrials # how many responses there are to be checked in total

        cvwave=CheckWaves(root)
        if not gv.done: # for rare cases of end-session crashes
            cvwave.focus_force()
            cvwave.wait_window(cvwave)

        ############################################################################
        # After wave+sound checking is exited, produce tab-separated list of RTs
        if gv.done: # do not save in case Tkinter loop is broken before list end
            self.save_output() 

            # delete status files to indicate successful completion
            self.statusfile1.close()
            os.remove(self.statusfilename1)
            self.statusfile2.write("\n")
            for param in sys.argv:
                self.statusfile2.write(param+"\n")
            self.statusfile2.write("Completed at "+str(datetime.datetime.now())+"\n")
            self.statusfile2.close()
            logfilename2=gv.expname+"-log.txt"
            try:
                os.rename(self.statusfilename2,logfilename2)
                tkMessageBox.showinfo("Done","CheckVocal terminated successfully")
            except:
                logmsg("Failed to rename log file %s to %s" % (self.statusfilename2,logfilename2))
                tkMessageBox.showwarning("Log file naming failure",
                ("Did not succeed in renaming %s to %s\nTake care of log file manually!" % (self.statusfilename2,logfilename2)))

        else: # not done
            logmsg("Exiting, not done")
            tkMessageBox.showinfo("Interrupt","CheckVocal was interrupted\nYou may continue at a later time")

        global_quit()
        #
        ############################################################################


class CheckVocalClass_Files(CheckVocalClass):

    def get_rtf_timeout(self):
        exiterror("Error in determining timeout")

    def process_azk(self):
        # there is no azk file to process here, need to make up info for audio file processing
        logmsg("Working folder is %s" % gv.expdir)
        dirlist=os.listdir(gv.expdir)
        #wavlist=[x for x in dirlist if x[-len(DEFAULT_WAVEXT):]==DEFAULT_WAVEXT] # original - misses ".WAV"
        #wavlist=[x for x in dirlist if x[-len(DEFAULT_WAVEXT):].upper()==DEFAULT_WAVEXT] # edited - assumes case insensitivity
        #wavlist=[x for x in dirlist if x[-len(DEFAULT_WAVEXT):] in FILES_WAVEXTS] # dangerous: not all extensions might have the same length!
        wavlist=[x for x in dirlist for y in FILES_WAVEXTS if x[-len(y):]==y] 
        self.ntrials=len(wavlist) # assume all audio files are workable "trials"
        if (self.ntrials<1): exiterror("No audio files found")
        logmsg("%i audio files found"%(self.ntrials))
        self.subject_select={}
        gv.Nsubj = 1 # assign all audio files to one dummy "subject"
        subjno = 1
        s_id = u"DummySubj"
        date_= "00/00/0000 00:00:00 on XXXXXX" # we could get a date from the audio files but it is not saved anyway
        refresh_= "refresh 00.00ms"
        logmsg( "Subject "+`subjno`+", ID="+s_id)
        gv.sub_ids[subjno]=s_id
        gv.sub_dates[s_id]=date_
        gv.sub_refresh[s_id]=refresh_
        gv.sub_nums[s_id]=subjno
        self.subject_select[s_id]=1
        for trial in range(self.ntrials):
            gv.listoftrials.append((trial+1,0.0)) # dummy item number, dummy RT // no sorting necessary
            gv.listoftrialsub.append(s_id)
            gv.trial_ind[trial]=trial
            gv.subj_ind[trial]=s_id
        gv.sub_trials[s_id]=copy.deepcopy(gv.listoftrials)
        gv.listoffiles=wavlist
        gv.totaltrials=self.ntrials      
        # sub_origlines[s_id]=azklines[startline:endline] # no original lines; not possible to save as azk
        logmsg( "End of setting up dummy trials for %s" % (gv.expdir))

    def process_ans(self):
        ansfilename=gv.expdir+"CheckFiles-ans.txt"
        try: # default location for answers same as rtf/azk/wav
            ansfile=codecs.open(ansfilename,"r",gv.char_encoding) # open straight into Unicode
        except: # if no -ans file found, default to no answer display
            logmsg( "Could not open %s to read correct answers!" % (ansfilename))
            answer=u"--"
            for trial in range(self.ntrials):
                gv.listofanswers.append(answer)
            return
        #
        # We need to make sure the answers match the audio file names
        #
        answers=myreadlines(ansfile)
        ansfile.close()
        i=0
        relist = []
        anlist = []
        for line in answers:
            i += 1
            if ( len(line)<2 ):
                exiterror ( "Empty line (%i) in answers file!" % (i) )
            try:
                item,answer=string.split(line.rstrip(),maxsplit=1)
            except:
                exiterror( "Error reading answers file\n(problem line %i, contents %s)" % (i,line.rstrip()) )
            rei = re.compile(item)
            nf  = len(filter(lambda f: re.search(rei,f),gv.listoffiles))
            if (nf < 1):
                logmsg( "Answer line %i (%s) not matching any files" % (i,item))
            relist.append(item)
            anlist.append(answer)

        gv.listofanswers = []
        for f in gv.listoffiles:
            fl = map(lambda a: re.search(a,f),relist)
            na = len(filter(lambda x: x,fl))
            if (na < 1):
                logmsg( "No spoken response for file %s"%(f))
                gv.listofanswers.append(u"--")
            elif (na > 1):
                exiterror( "Multiple answers matching file %s"%(f))
            else:
                ai = [i for i,a in enumerate(fl) if a!=None]
                if (len(ai)<>1):
                    exiterror( "Something weird happened while processing answers file") # this should not happen
                gv.listofanswers.append(anlist[ai[0]])

    
    def select_subjects(self):
        pass # no subjects to select

    def verify_audiofiles(self):
        pass # nothing to verify any more

    def save_output(self):
        outfilename=gv.expname+"-datalist.txt"
        outfile=self.verify_output_filename(outfilename)

        live_subjects=gv.sub_ids.keys() # single dummy subject, necessarily "selected"
        ref_subj=gv.sub_trials.keys()[0]
        
        if (gv.save_rows==0): # save subject data in columns, default
##            outfile.write("item")
##            #for cur_subj in gv.sub_trials.keys():
##            for subjnum in live_subjects:
##                cur_subj=gv.sub_ids[subjnum]
##                outfile.write(gv._SEP+cur_subj.encode(gv.char_encoding,'replace'))
##            outfile.write("\n")
            for trial in gv.sub_trials[ref_subj]:
                #outfile.write(`trial[0]`)
                outfile.write(gv.listoffiles[trial[0]-1])
                for subjnum in live_subjects:
                    cur_subj=gv.sub_ids[subjnum]
                    outfile.write(gv._SEP+"%.1f"%(gv.sub_trials[cur_subj][gv.sub_trials[ref_subj].index(trial)][1]))
                outfile.write("\n")
        else: # save data in rows
            outfile.write("subject")
            for trial in gv.sub_trials[ref_subj]:
                #outfile.write(gv._SEP+`trial[0]`)
                outfile.write(gv._SEP+gv.listoffiles[trial[0]-1])
            outfile.write("\n")
            for subjnum in live_subjects:
                cur_subj=gv.sub_ids[subjnum]
                outfile.write(cur_subj.encode(gv.char_encoding,'replace'))
                for trial in gv.sub_trials[cur_subj]:
                    outfile.write(gv._SEP+"%.1f"%(trial[1]))
                outfile.write("\n")
        outfile.close()
        logmsg( "Successfully wrote "+outfilename)


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

myfile=sys.argv[0]
myname=os.path.splitext(os.path.basename(myfile))[0]
finfo=os.path.getmtime(myfile)
mtime=datetime.date.isoformat(datetime.date.fromtimestamp(finfo))
if (myname=="CheckVocal"): DMDXMODE=True
elif (myname=="CheckFiles"): DMDXMODE=False
# otherwise, it is a development version, in which case the default (defined at the top) applies

CurDir=os.getcwdu() # u for unicode; really important for Tkinter!
if DMDXMODE: IconFile=os.path.join(CurDir,u"icons",u"cv.ico")
else: IconFile=os.path.join(CurDir,u"icons",u"cf.ico")
try:
    if (not os.path.exists(IconFile)): IconFile="None"
except:
    IconFile="None" # to catch any problems

root=Tk()
root.title(u"CheckVocal main")
if (IconFile!="None"): root.wm_iconbitmap(IconFile)
root.withdraw()

gv=GlobVariables()
gv.reset()

msgwindow=TextWindow(root)
msgwindow.body()

if DMDXMODE:
    cv_process=CheckVocalClass()  
    startw=SetupWindow(root)
else:
    cv_process=CheckVocalClass_Files()  
    startw=SetupWindow_Files(root)

startw.startup()

root.mainloop()
