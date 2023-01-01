import tkinter as tk
from tkinter import messagebox as msg
from tkinter import filedialog as fd
from PIL import Image, ImageTk
import shutil
import os
import glob
from datetime import date
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from tkinter import filedialog as fd
from tkinter import simpledialog as sd
import pandas as pd
import numpy as np
import win32gui
import re
import win32com.client
pathwin32=win32com.__gen_path__
directory_contents = os.listdir(pathwin32)
for item in directory_contents:    
    if os.path.isdir(os.path.join(pathwin32, item)):        
        shutil.rmtree(os.path.join(pathwin32, item))
import os
import time
import xlwings as xw
import sys


#Initial parameters and paths
today=date.today()
DF={}
cnt=0
itemcnt=0
currpath=os.getcwd()
odpath=currpath.split("2019",1)
odpath=odpath[0]+"2019/"
xlspath= odpath+ "00FlightPlan Runway Analysis\\" + str(today) + ".xlsx"
PDFcur=odpath + '00FlightPlan Runway Analysis\Flightplan PDFs\Current'
PDFarc=odpath+ '00FlightPlan Runway Analysis\Flightplan PDFs\Archieve'
#Initial parameters and paths

#sub-definitions to be used in definitions

def convertp2t(i):
    
    
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp=open(FPs[i], 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()
    
    FPpaths[i]=FPpaths[i].replace("/","\\")
    
    replacepath=odpath+ r'00FlightPlan Runway Analysis\Flightplan PDFs\Current'
    FPpaths[i]=FPpaths[i].replace(replacepath,"")  
    
    
    splitted=text.split(" ")
    
    
    global db
    for ijk in range(len(splitted)):
        if splitted[0]=="PLAN":
            
            
            
    
            for j in range(len(splitted)):
                if splitted[j]=="EST":
                    MTOW=splitted[j+6]
                
                    LNDW=splitted[j+9]
                   

          
            for k in range(len(splitted)):
                if splitted[k]=="738W" or splitted[k]=="8MX7":
                    Arr=splitted[k-1]
                    Dep=splitted[k-3]
                    
                    

            db=[Dep, MTOW, Arr, LNDW, FPpaths[i]]
            
            break

        else:
            db=[]
            
            
    
    

    

def openfile():
    global FPs
    FPs = fd.askopenfilenames(filetypes = [("PDF Files",".pdf")])
    global FPpaths
    FPpaths=list(FPs)
    global Df
    Df=[]
    DFend=[]
    global readcnt
    global ureadcnt
    readcnt=0
    ureadcnt=0
    
    
    for i in range(len(FPs)):

        #subdefinition
        convertp2t(i)
        #subdefinition
        if db!=[]:
            DFend.append(db)
            readcnt +=1 

        else:
            ureadcnt +=1
        
        
        
        
        
    
    Df=pd.DataFrame(DFend)
    
#sub-definitions to be used in definitions

    
#Definitions to be used in button labels

def newsession():
    try:
        
        res = msg.askyesnocancel('Are you sure?','If you click on Yes, all the downloaded flight plans will be lost and plan database will be empty. If you click No, only plan database will be erased, downloaded flight plans will not be erased. Cancel stops execution.')
        if res==True:
            DF={}
            Dfout={}
            DFout=[]
            cnt=0
            db=[]
            DFend=[]
            Df=[]
            cntlabel.config(text=cnt)
            if os.path.exists(PDFcur):
                shutil.rmtree(PDFcur)
                msg.showinfo('Success!' , 'New session has been started!')
                os.makedirs(PDFcur)
            else:
                os.makedirs(PDFcur)
        elif res==False:
            DF={}
            Dfout={}
            db=[]
            cnt=0
            DFend=[]
            DFout=[]
            Df=[]
            cntlabel.config(text=cnt)
            msg.showinfo('Done!' , 'Flight Plans are still exist!')
        else:
            msg.showinfo('Interruption!' , 'Cancelled!')
            
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered;  '+str(e))

def copysession():
    try:
        itms=glob.glob(PDFcur+ '/*')
        for it in range(len(itms)):
            shutil.move(itms[it], PDFarc)
        msg.showinfo('Success!' , 'The session has been archived!')
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered in Archiving Session;  '+str(e))
        

def TOPet():
    try:
        TOfile = fd.askopenfilename(filetypes = [("Text files",".txt")])
        Old_PetTO=pd.read_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODB.csv', header=0, names=("Airport", "Runway", "MTOW"))                
        PetTO=open(TOfile, "r")
        TOline=PetTO.readline()
        cntTO=1
        DFTO=[]
        DFTO=pd.DataFrame(DFTO)
        DFTO=DFTO.T

        while TOline:
            tairpcnt=0
            if TOline[1:5]=="Elev":
               splitted=TOline.split(" ")
               Airport=splitted[-1].strip()
               for i in range(len(splitted)):
                   if splitted[i]=="Runway":
                       
                       RWY=splitted[i+1]
                       ARWY=Airport+"-"+RWY              
               cntrange=range(cntTO,cntTO+15)
               for cntrdd in cntrange:
                   TOline=PetTO.readline()
                   if TOline[2:4]=="30":
                       MTOW=TOline[12:17]                       
               TripleTO=(Airport, RWY, MTOW)
               try:

                   DFTOdum=pd.DataFrame(TripleTO)#, columns=columnsTO)
                   DFTOdum=DFTOdum.T
                   int(DFTOdum.iat[0,2])
                   
                   if DFTO.empty:
                       DFTO=DFTO.append(DFTOdum)
                   
                   elif DFTOdum.iat[0,0]==DFTO.iat[-1,0]:
                       if int(DFTOdum.iat[0,2])>=int(DFTO.iat[-1,2]):
                       
                           DFTO=DFTO.drop(DFTO.index[-1])
                           DFTO=DFTO.append(DFTOdum)
                                
                   else:
                        DFTO=DFTO.append(DFTOdum)
                    
                   DFTO= DFTO.reset_index(drop=True)
                   tairpcnt +=1
               except:
                   print('selam')               
            TOline=PetTO.readline()
            cntTO += 1

        DFTO.columns=["Airport", "Runway", "MTOW"]
        DFTO=Old_PetTO.append(DFTO).drop_duplicates(['Airport'],keep='last')
        DFTO.to_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODB.csv', header=("Airport", "Runway", "MTOW"),index=False)
        TODBlabel.config(text=TODBrowcnt)

        msg.showinfo('Success!' , 'Done!' + str(tairpcnt)+ " Airports have been updated!")
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during Performance DB Update;  '+str(e))


def TOPet24K():
    try:
        TOfile = fd.askopenfilename(filetypes = [("Text files",".txt")])
        Old_PetTO=pd.read_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODB24K.csv', header=0, names=("Airport", "Runway", "MTOW"))                
        PetTO=open(TOfile, "r")
        TOline=PetTO.readline()
        cntTO=1
        DFTO=[]
        DFTO=pd.DataFrame(DFTO)
        DFTO=DFTO.T

        while TOline:
            tairpcnt=0
            if TOline[1:5]=="Elev":
               splitted=TOline.split(" ")
               Airport=splitted[-1].strip()
               for i in range(len(splitted)):
                   if splitted[i]=="Runway":
                       
                       RWY=splitted[i+1]
                       ARWY=Airport+"-"+RWY              
               cntrange=range(cntTO,cntTO+15)
               for cntrdd in cntrange:
                   TOline=PetTO.readline()
                   if TOline[2:4]=="30":
                       MTOW=TOline[12:17]                       
               TripleTO=(Airport, RWY, MTOW)
               try:

                   DFTOdum=pd.DataFrame(TripleTO)#, columns=columnsTO)
                   DFTOdum=DFTOdum.T
                   int(DFTOdum.iat[0,2])
                   
                   if DFTO.empty:
                       DFTO=DFTO.append(DFTOdum)
                   
                   elif DFTOdum.iat[0,0]==DFTO.iat[-1,0]:
                       if int(DFTOdum.iat[0,2])>=int(DFTO.iat[-1,2]):
                       
                           DFTO=DFTO.drop(DFTO.index[-1])
                           DFTO=DFTO.append(DFTOdum)
                                
                   else:
                        DFTO=DFTO.append(DFTOdum)
                    
                   DFTO= DFTO.reset_index(drop=True)
                   tairpcnt +=1
               except:
                   print('selam')               
            TOline=PetTO.readline()
            cntTO += 1

        DFTO.columns=["Airport", "Runway", "MTOW"]
        DFTO=Old_PetTO.append(DFTO).drop_duplicates(['Airport'],keep='last')
        DFTO.to_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODB24K.csv', header=("Airport", "Runway", "MTOW"),index=False)
        TODBlabel.config(text=TODBrowcnt)

        msg.showinfo('Success!' , 'Done!' + str(tairpcnt)+ " Airports have been updated!")
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during Performance DB Update;  '+str(e))

def TOPetMAX():
    try:
        TOfile = fd.askopenfilename(filetypes = [("Text files",".txt")])
        Old_PetTO=pd.read_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODBMAX.csv', header=0, names=("Airport", "Runway", "MTOW"))                
        PetTO=open(TOfile, "r")
        TOline=PetTO.readline()
        cntTO=1
        DFTO=[]
        DFTO=pd.DataFrame(DFTO)
        DFTO=DFTO.T

        while TOline:
            tairpcnt=0
            if TOline[1:5]=="Elev":
               splitted=TOline.split(" ")
               Airport=splitted[-1].strip()
               for i in range(len(splitted)):
                   if splitted[i]=="Runway":
                       
                       RWY=splitted[i+1]
                       ARWY=Airport+"-"+RWY              
               cntrange=range(cntTO,cntTO+15)
               for cntrdd in cntrange:
                   TOline=PetTO.readline()
                   if TOline[2:4]=="30":
                       MTOW=TOline[12:17]                       
               TripleTO=(Airport, RWY, MTOW)
               try:

                   DFTOdum=pd.DataFrame(TripleTO)#, columns=columnsTO)
                   DFTOdum=DFTOdum.T
                   int(DFTOdum.iat[0,2])
                   
                   if DFTO.empty:
                       DFTO=DFTO.append(DFTOdum)
                   
                   elif DFTOdum.iat[0,0]==DFTO.iat[-1,0]:
                       if int(DFTOdum.iat[0,2])>=int(DFTO.iat[-1,2]):
                       
                           DFTO=DFTO.drop(DFTO.index[-1])
                           DFTO=DFTO.append(DFTOdum)
                                
                   else:
                        DFTO=DFTO.append(DFTOdum)
                    
                   DFTO= DFTO.reset_index(drop=True)
                   tairpcnt +=1
               except:
                   print('selam')               
            TOline=PetTO.readline()
            cntTO += 1

        DFTO.columns=["Airport", "Runway", "MTOW"]
        DFTO=Old_PetTO.append(DFTO).drop_duplicates(['Airport'],keep='last')
        DFTO.to_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODBMAX.csv', header=("Airport", "Runway", "MTOW"),index=False)
        TODBlabel.config(text=TODBrowcnt)

        msg.showinfo('Success!' , 'Done!' + str(tairpcnt)+ " Airports have been updated!")
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during Performance DB Update;  '+str(e))

def LDPet():
    try:
        LDfile = fd.askopenfilename(filetypes = [("Text files",".txt")])
        Old_PetLD=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDB.csv', header=0, names=("Airport", "Runway", "LNDW"))
        PetLD=open(LDfile, "r")

        LDline=PetLD.readline()
        cntLD=1
        DFLD=[]
        DFLD=pd.DataFrame(DFLD)
        DFLD=DFLD.T

        while LDline:
            lairpcnt=0
            if LDline[1:5]=="Elev":
               splitted=LDline.split(" ")
               Airport=splitted[-1].strip()
               for i in range(len(splitted)):
                   if splitted[i]=="Runway":
                       
                       RWY=splitted[i+1]      
              
               cntrange=range(cntLD,cntLD+15)
               for cntldd in cntrange:
                   LDline=PetLD.readline()
                   if LDline[2:4]=="30":
                       LNDW=LDline[12:17]
                       
               TripleLD=(Airport, RWY, LNDW)
               try:
                   DFLDdum=pd.DataFrame(TripleLD)#, columns=columnsLD)
                   DFLDdum=DFLDdum.T
                   int(DFLDdum.iat[0,2])
                               
                   if DFLD.empty:
                       DFLD=DFLD.append(DFLDdum)
                   
                   elif DFLDdum.iat[0,0]==DFLD.iat[-1,0]:
                       if int(DFLDdum.iat[0,2])>=int(DFLD.iat[-1,2]):
                       
                           DFLD=DFLD.drop(DFLD.index[-1])
                           DFLD=DFLD.append(DFLDdum)
                      
                   else:
                        DFLD=DFLD.append(DFLDdum)
                    
                   DFLD= DFLD.reset_index(drop=True)
                   lairpcnt +=1
               except:
                   print('selam2')          
              
            
            LDline=PetLD.readline()
            cntLD += 1

        DFLD.columns=["Airport", "Runway", "LNDW"]

        DFLD=Old_PetLD.append(DFLD).drop_duplicates(['Airport'],keep='last')

        DFLD.to_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDB.csv', header=("Airport", "Runway", "LNDW"),index=False)
        LDDBlabel.config(text=LDDBrowcnt)

        msg.showinfo('Success!' , 'Done!'+ str(lairpcnt)+ " Airports have been updated!")
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during Performance DB Update;  '+str(e))
        
def LDPetMAX():
    try:
        LDfile = fd.askopenfilename(filetypes = [("Text files",".txt")])
        Old_PetLD=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDBMAX.csv', header=0, names=("Airport", "Runway", "LNDW"))
        PetLD=open(LDfile, "r")

        LDline=PetLD.readline()
        cntLD=1
        DFLD=[]
        DFLD=pd.DataFrame(DFLD)
        DFLD=DFLD.T

        while LDline:
            lairpcnt=0
            if LDline[1:5]=="Elev":
               splitted=LDline.split(" ")
               Airport=splitted[-1].strip()
               for i in range(len(splitted)):
                   if splitted[i]=="Runway":
                       
                       RWY=splitted[i+1]      
              
               cntrange=range(cntLD,cntLD+15)
               for cntldd in cntrange:
                   LDline=PetLD.readline()
                   if LDline[2:4]=="30":
                       LNDW=LDline[12:17]
                       
               TripleLD=(Airport, RWY, LNDW)
               try:
                   DFLDdum=pd.DataFrame(TripleLD)#, columns=columnsLD)
                   DFLDdum=DFLDdum.T
                   int(DFLDdum.iat[0,2])
                               
                   if DFLD.empty:
                       DFLD=DFLD.append(DFLDdum)
                   
                   elif DFLDdum.iat[0,0]==DFLD.iat[-1,0]:
                       if int(DFLDdum.iat[0,2])>=int(DFLD.iat[-1,2]):
                       
                           DFLD=DFLD.drop(DFLD.index[-1])
                           DFLD=DFLD.append(DFLDdum)
                      
                   else:
                        DFLD=DFLD.append(DFLDdum)
                    
                   DFLD= DFLD.reset_index(drop=True)
                   lairpcnt +=1
               except:
                   print('selam2')          
              
            
            LDline=PetLD.readline()
            cntLD += 1

        DFLD.columns=["Airport", "Runway", "LNDW"]

        DFLD=Old_PetLD.append(DFLD).drop_duplicates(['Airport'],keep='last')

        DFLD.to_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDBMAX.csv', header=("Airport", "Runway", "LNDW"),index=False)
        LDDBlabel.config(text=LDDBrowcnt)

        msg.showinfo('Success!' , 'Done!'+ str(lairpcnt)+ " Airports have been updated!")
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during Performance DB Update;  '+str(e))

def PDFread():
    try:      
        
        #subdefinition
        openfile()
        #subdefinition
        global cnt
        
        DF[cnt]=Df

        
        cnt +=1
        cntlabel.config(text=cnt)
        msg.showinfo('Attention!' , 'Done! '+ str(readcnt)+" Flight Plans have been processed succesfully, " + str(ureadcnt) + " Flight Plans couldn't be read!")
        
    except Exception as e:
        msg.showinfo('Error!' , 'An error has been countered during FP Upload;  '+str(e))


def outlookitems():
    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
          

            
        sfolder = outlook.GetDefaultFolder(6).Folders.Item("aTodo-RWY analysis")
       
        messages = sfolder.Items
        print(messages)
        main_path=odpath+ "00FlightPlan Runway Analysis\Flightplan PDFs\Current"
        foldercnt=1
        for i in range(len(messages)):
            
            msggg=messages[i+1]
            print(msggg)

            date = msggg.SentOn.strftime("%d-%m-%y")
            subject = msggg.Subject
            print(subject)
            subject=subject.replace(".","")
            subject=subject.replace(":","")
            subject=subject.replace(" ","")
            subject=subject.replace("/","")
            
            subject=date+"-"+subject+"\\"

            if os.path.exists(main_path+"\\"+subject)==False:

                os.mkdir(main_path+"\\"+subject)
            else:
                subject=subject.replace("\\",str(foldercnt)+"\\")
                foldercnt=foldercnt+1
                os.mkdir(main_path+"\\"+subject)
                

            get_path=(main_path+"\\"+subject)




            
            attachments = messages[i+1].Attachments
            num_attach = len([cor for cor in attachments])
            for cor in range(1, num_attach+1):
                att = attachments.Item(cor)
                fname=att.FileName    
           
                
                if fname.lower().endswith(".pdf"):
                    att.SaveAsFile(get_path + fname)
                    
                            
                        
        
        analycnt=glob.glob(main_path+ '/*')
        global itemcnt
        itemcnt=len(analycnt)
        itemcntlabel.config(text=itemcnt)
        
        msgcnt=len(messages)
        uitems=msgcnt-itemcnt
        

        msg.showinfo('Success!' , 'Done! '+ str(itemcnt)+" Analyses have been downloaded succesfully, " + str(uitems)+ " messages couldn't be downloaded!")      
    except Exception as e:
        exc_type, exc_obj, tb = sys.exc_info()
        print ('Error on line {}'.format(sys.exc_info()[-1].tb_lineno))
        print ('Error on obj {}'.format(sys.exc_info()[-2]))
        print ('Error on type {}'.format(sys.exc_info()[-3]))
        msg.showinfo('Error!' , 'An error has been countered during downloading items from MS Outlook;  '+str(e))
    
       

def outexcel():
    
   
        TODBB=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceTODB.csv', header=0, names=("Airport", "Runway", "MTOW"))
        TODBB24=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceTODB24K.csv', header=0, names=("Airport", "Runway", "MTOW"))
        TODBBMAX=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceTODBMAX.csv', header=0, names=("Airport", "Runway", "MTOW"))

        
        LDDBB=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDB.csv', header=0, names=("Airport", "Runway", "LNDW"))
        LDDBBMAX=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDBMAX.csv', header=0, names=("Airport", "Runway", "LNDW"))
        Dfout={}
        writer = pd.ExcelWriter(odpath+ '00FlightPlan Runway Analysis\\' + str(today)+ "_Perf.xlsx", engine='xlsxwriter')
        colhead=['TO Airport', 'TO Runway', "Est TOW", "24K-Max TOW" , "26K-Max TOW", "26K-SFP-Max TOW", "27K-Max TOW","737-8 Max TOW", "LD Airport", "LD Runway", "Est LDW", "24K-Max LDW" , "26K-Max LDW", "26K-SFP-Max LDW", "27K-Max LDW","737-8 Max LDW","Comment", "Mail"] 
        global cnt
        for t in range(cnt):
            try:     
                DFout = pd.DataFrame(np.zeros((len(DF[t].index), 18)), dtype=str, columns=colhead)
                Dfin=len(DF[t].index)
                Mailplace=str(DF[t].iat[0,4])
                Mailplace=Mailplace[129:]
                
                DFout.iat[0,-1]=Mailplace
                
                for tt in range(len(DF[t].index)):
                    DFout.iat[tt,0]=DF[t].iat[tt, 0]
                    DFout.iat[tt,2]=DF[t].iat[tt, 1]
                    DFout.iat[tt,8]=DF[t].iat[tt, 2]
                    DFout.iat[tt,10]=DF[t].iat[tt, 3]

                    
                    to24df = TODBB24[TODBB24["Airport"] == DF[t].iat[tt, 0]]    
                    to26df = TODBB[TODBB["Airport"] == DF[t].iat[tt, 0]]
                    tomaxdf= TODBBMAX[TODBBMAX["Airport"] == DF[t].iat[tt, 0]]
                    try:
                        DFout.iat[tt,1]=to26df.iat[0,1]
                    except:
                        DFout.iat[tt,1]=0
                    try:
                        DFout.iat[tt,3]=to24df.iat[0,2]  
                    except:
                        DFout.iat[tt,3]=0
                    try:
                        DFout.iat[tt,4]=to26df.iat[0,2]
                    except:
                        DFout.iat[tt,4]=0
                    try:
                        DFout.iat[tt,7]=tomaxdf.iat[0,2]
                    except:
                        DFout.iat[tt,7]=0    
                 

                    
                    if int(DFout.iat[tt,4])>=int(DFout.iat[tt,2]):
                        DFout.iat[tt,5]="O.K."
                        DFout.iat[tt,6]="O.K."
                    else:
                        DFout.iat[tt,5]="Check!"
                        DFout.iat[tt,6]="Check!"

                    lddfdummy=LDDBB[LDDBB["Airport"] == DF[t].iat[tt, 2]]
                    lddfmax=LDDBBMAX[LDDBBMAX["Airport"] == DF[t].iat[tt, 2]]
                    try:
                        DFout.iat[tt,9]=lddfdummy.iat[0, 1]
                    except:
                        DFout.iat[tt,9]=0
                    try:
                        DFout.iat[tt,11]=lddfdummy.iat[0, 2]     
                    except:
                        DFout.iat[tt,11]=0
                    try:
                        DFout.iat[tt,12]=lddfdummy.iat[0, 2]
                    except:
                        DFout.iat[tt,12]=0
                    try:
                        DFout.iat[tt,15]=lddfmax.iat[0, 2]
                    except:
                        DFout.iat[tt,15]=0
                 
                       
                        
                    if int(DFout.iat[tt,12])>=int(DFout.iat[tt,10]):
                        DFout.iat[tt,13]="O.K."
                        DFout.iat[tt,14]="O.K."
                    else:
                        DFout.iat[tt,13]="Check!"
                        DFout.iat[tt,14]="Check!"
                        
                             

                    
                
                               
                
                Dfout[t]=DFout           

                
                sheetname="Sheet"+str(t+1)
                Dfout[t].to_excel(writer, header=colhead, sheet_name=sheetname)
            except Exception as e:
                print("ErrorOn: " + str(t)+" "+ str(e))
                exc_type, exc_obj, exc_tb = sys.exc_info()
                print(exc_tb.tb_lineno)
                
        writer.save()
        
        
        mwb= xw.Book(odpath+ '00FlightPlan Runway Analysis\\macro.xlsm')
        app=mwb.app
        mm = app.macro('styling')
    
        mm()

        msg.showinfo('Attention!' , 'Excel File has been created!')
        
    


    
    
#Definitions to be used in button labels     
        
        
#Tkinter
frame=tk.Tk()
frame.geometry('750x550')
frame.title("FOE-Performance Analysis Module")
frame.configure(background='black')

imgpath0=odpath+ 'phyton\db_python\Performance Module\corendon.png'
img= Image.open(imgpath0)
img=img.resize((160,55), Image.ANTIALIAS)
img = ImageTk.PhotoImage(img, master=frame)

panel = tk.Label(frame, image = img, borderwidth=2,  relief='solid', bg="black")
panel.place(relx=0.24, rely=0.13, anchor='se')

foelabel = tk.Label(frame, text="Flight Operations Engineering", font=("Arial Bold", 15), bg="black", fg="white")
foelabel.place(relx=.65, rely=0.12, anchor='n')
pamlabel = tk.Label(frame, text="Performance Analysis Module", font=("Arial Bold", 20), bg="black", fg="white")
pamlabel.place(relx=.65, rely=0.06, anchor='n')






line=tk.Canvas(frame, width=0, height=270, bg="white",highlightthickness=0 )
line.place(x=475, y=130)

line.create_line(500, 200, 500, 750,
              fill="gray", width=1)


rect1=tk.Canvas(frame, width=200, height=270, bg="black", highlightthickness=0)
rect1.place(x=40, y=120)
rect1.create_rectangle(10, 10, 190, 135,
              fill=None, width=1, outline="white")
rect1.create_rectangle(10, 140, 190, 265,
              fill=None, width=1, outline="white")
rect1.create_rectangle(5, 5, 195, 270,
              fill=None, width=2, outline="white")
rect1.create_text(100, 35,
              fill="white", text="Operations to be",font=("Arial Bold", 12), width=180)
rect1.create_text(100, 55,
              fill="white", text="Analyzed",font=("Arial Bold", 12), width=180)
rect1.create_text(100, 165,
              fill="white", text="Analyses to be",font=("Arial Bold", 12), width=180)
rect1.create_text(100, 185,
              fill="white", text="Exported In Excel",font=("Arial Bold", 12), width=180)

rect2=tk.Canvas(frame, width=350, height=80, bg="black", highlightthickness=2)
rect2.place(x=40, y=420)

rect2.create_rectangle(5, 5, 180, 80,
              fill=None, width=1, outline="white")

rect2.create_rectangle(183, 5, 349, 80,
              fill=None, width=1, outline="white")

rect2.create_text(87, 20,
              fill="white", text="#TO Airport",font=("Arial Bold", 12), width=175)

rect2.create_text(267, 20,
              fill="white", text="#LD Airport",font=("Arial Bold", 12), width=175)




itemcntlabel = tk.Label(rect1, text=itemcnt, font=("Arial Bold", 25), bg="black", fg="white")
itemcntlabel.place(x=100, y=75, anchor='n')

cntlabel = tk.Label(rect1, text=cnt, font=("Arial Bold", 25), bg="black", fg="white")
cntlabel.place(x=100, y=210, anchor='n')

Bs=tk.Button(frame, text="New Session", command = newsession, height = 3, width = 20)
Bs.place(x=300, y=230)

Cs=tk.Button(frame, text="Archive Session", command = copysession, height = 3, width = 20)
Cs.place(x=300, y=330)


def TOBTN():

    answerTO = sd.askstring("Input", "What is the configuration of TO file? Type '24K' or '26K' or 'MAX'")

    if answerTO=='24K':
        commandTO=TOPet24K()
    elif answerTO=='26K':
        commandTO=TOPet()
    elif answerTO=='MAX':
        commandTO=TOPetMAX()
    else:
        msg.showwarning("Warning","Wrong TO Configuration Input")
    return commandTO


def LDBTN():

    answerTO = sd.askstring("Input", "What is the configuration of LD file? Type '26K' or 'MAX'")

    
    if answerTO=='26K':
        commandLD=LDPet()
    elif answerTO=='MAX':
        commandLD=LDPetMAX()
    else:
        msg.showwarning("Warning","Wrong LD Configuration Input")
    return commandLD 

TObtn=tk.Button(frame, text="TOPetFile", command = TOBTN, height = 3, width = 20)
TObtn.place(x=500, y=130)    



LDbtn=tk.Button(frame, text="LDPetFile", command = LDBTN, height = 3, width = 20)
LDbtn.place(x=500, y=230)


PDFreadbtn=tk.Button(frame, text="Read Flight Plans", command = PDFread, height = 3, width = 20)
PDFreadbtn.place(x=500, y=330)

Exceloutbtn=tk.Button(frame, text="Produce Excel File", command = outexcel, height = 4, width = 30)
Exceloutbtn.place(x=430, y=430)

Dwnloutlook=tk.Button(frame, text="Download Outlook Items", command = outlookitems, height = 3, width = 20)
Dwnloutlook.place(x=300, y=130)

PetLDDB=pd.read_csv(odpath+ '00FlightPlan Runway Analysis\Output Files\PerformanceLDDB.csv', header=0, names=("Airport", "Runway", "LNDW"))
LDDBrowcnt=len(PetLDDB.index)+1
PetTODB=pd.read_csv(odpath+'00FlightPlan Runway Analysis\Output Files\PerformanceTODB.csv', header=0, names=("Airport", "Runway", "MTOW"))
TODBrowcnt=len(PetTODB.index)+1

TODBlabel = tk.Label(rect2, text=TODBrowcnt, font=("Arial Bold", 25), bg="black", fg="white")
TODBlabel.place(relx=0.25, rely=.650, anchor='c')

LDDBlabel = tk.Label(rect2, text=LDDBrowcnt, font=("Arial Bold", 25), bg="black", fg="white")
LDDBlabel.place(relx=0.75, rely=.650, anchor='c')


#Tkinter

tk.mainloop()
