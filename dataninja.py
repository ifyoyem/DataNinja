"""Copyright 2013 Ifeanyi Oyem

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License."""




#Imported libraries

#Imported to use arrays
import numpy
#Used to create arrays for the count matrices and vectors
from numpy import zeros, array
#Imported to use the sqrt function to decide what k should be
import math
#Used to manipulate files
import os
#Imported to use the cluster module
import nltk 
from nltk import cluster
#Used the GAAClusterer class to cluster documents, used cosine_distance to measure document similarity
from nltk.cluster import GAAClusterer, cosine_distance 
import pyPdf
#Used to read pdf files
from pyPdf import PdfFileReader
#Used to read docx files
import zipfile, re 
import xlwt
#Used to write and style .xls spreadsheets
from xlwt import Workbook,easyxf,Formula
#Used choice to randomly choose numbers from a list
from random import choice
#Used to get the date/time for outputted files
import datetime
#Used to get plurals for keywords entered for the "Sentence Search" feature
import inflect
#GUI package
import Tkinter 
from Tkinter import *
#Widgets to interact with the user, select file, entry box, message box
import tkFileDialog, tkSimpleDialog, tkMessageBox


#Make this variable accessible to save the current time as part of file names and inside the audit report draft management letter
global time
#Get the current time
time = datetime.datetime.now() 

#Make this variable accessible to parse files
global stopwords 
#Stopwords are words that will be removed/ignored from words extracted from files
stopwords = [
                 'a', 'able', 'and','edition','for','in','little','of','the','to', 'i', 'is', 'in', 'yes', 'or', 'if', 'therefore', 'however', 'but', 'of', 'no',
                 'yet', 'put', 'on', 'my', 'it', 'you', 'was', 'am', 'as', 'what', 'that', 'he', 'maybe', 'just', 'will', 'with', 'at', 'an',
                 'this', 'would', 'your', 'thing', 'are', 'la', 'ba', 'me', 'have', 'from', 'which', 'so', 'be', 'where', 'about', 'above',
                 'accordance', 'a', 'about', 'above', 'after', 'again', 'all', 'am', 'an', 'and', 'any', 'are', "aren't", 'as', 'at', 'be', 'because', 'been', 'before',
                 'being', 'below', 'between', 'both', 'but', 'by', "can't", 'cannot', 'could', "couldn't", 'did', "didn't", 'do', 'does', "doesn't", 'doing', "don't",
                 'down', 'during', 'each', 'few', 'for', 'from', 'further', 'had', "hadn't", 'has', "hasn't", 'have', "haven't", 'having', 'he', "he'd", "he'll", "he's",
                 'her', 'here', "here's", 'hers', 'herself', 'him', 'himself', 'his', 'how', "how's", 'i', "i'd", "i'll", "i'm", "i've", 'if', 'in', 'into', 'is', "isn't",
                 'it', "it's", 'its', 'itself', "let's", 'me', 'more', 'most', "mustn't", 'my', 'myself', 'no', 'nor', 'not', 'of', 'off', 'on', 'once', 'only', 'or', 'other',
                 'ought', 'our', 'ours', 'ourselves', 'out', 'over', 'own', 'same', "shan't", 'she', "she'd", "she'll", "she's", 'should', "shouldn't", 'so', 'some', 'such',
                 'than', 'that', "that's", 'the', 'their', 'theirs', 'them', 'themselves', 'then', 'there', "there's", 'these', 'they', "they'd", "they'll", "they're", "they've",
                 'this', 'those', 'through', 'though', 'to', 'too', 'under', 'until', 'up', 'very', 'was', "wasn't", 'we', "we'd", "we'll", "we're", "we've", 'were', "weren't",
                 'what', "what's", 'when', "when's", 'where', "where's", 'which', 'while', 'who', "who's", 'whom', 'why', "why's", 'with', "won't", 'would', "wouldn't", 'you',
                 "you'd", "you'll", "you're", "you've", 'your', 'yours', 'ur', 'use', 'yourself', 'yourselves', "ain't", 'all', 'us', 'u', 'thank', 'always', 'say', 'within', 'used',
                 'n', 'returned', 'page', 'name', 'ha', 'please'
            ] 

#Make this variable accessible to save processed files to the users desktop
global desktop
#Get the user's desktop path
desktop = os.path.join(os.path.expanduser("~"), "Desktop") 

#Make this variable accessible to save processed files to the users desktop in the correct format
global dpath
#Change the backward slashes in the path to forward slashes so that it can be processed by the xlwt module
dpath = r''+str(desktop).replace('\\', '/')  

#Function to change the background colour of the root window in the future if desired
def colour1(b4):
    #Make this variable accessible from the DNInter class
    global b2 
    b2 = b4

#Function to change the background colour of the buttons on the root window in the future if desired
def colour2(b5):
    #Make this variable accessible from the DNInter class
    global b3 
    b3 = b5

#Interface class
class DNInter(Frame):
    
    #Initialise itself and the parent object
    def __init__(self, parent):
        #Initialise frame objects 
        Frame.__init__(self, parent, bg=b2)
        #Parent refers to the parent widget which is the root window
        self.parent = parent 

        #Change the title next to the logo to "Data Ninja"
        self.parent.title("Data Ninja")
        #Place inside the window and fill the parent window
        self.pack(fill=BOTH, expand=1) 

        #Create button with text and format button
        self.btn = Button(self, text="Cluster Documents", bg=b3, fg="#0BB5FF", font="bold", cursor="circle", relief="flat", height="27", width="25", activeforeground="white", activebackground="#42C0FB", command=self.clusterdoc)        
        #Place the button in a specific position
        self.btn.place(x=208, y=95)
        #Place button inside the frame
        self.pack() 

        #Create button with text and format button
        self.btn = Button(self, text="Smart Search", bg=b3, fg="#0BB5FF", font="bold", cursor="circle", relief="flat", height="27", width="25", activeforeground="white", activebackground="#42C0FB", command=self.search1)
        #Place the button in a specific position
        self.btn.place(x=445, y=95)
        #Place button inside the frame
        self.pack() 

        #Create button with text and format button
        self.btn = Button(self, text="Sentence Search", bg=b3, fg="#0BB5FF", font="bold", cursor="circle",  relief="flat", height="27", width="25", activeforeground="white", activebackground="#42C0FB", command=self.search2)
        #Place the button in a specific position
        self.btn.place(x=682, y=95)
        #Place button inside the frame
        self.pack() 

        #Create button with text and format button
        self.btn = Button(self, text="Document Similarity", bg=b3, fg="#0BB5FF", font="bold", cursor="circle", relief="flat", height="27", width="25", activeforeground="white", activebackground="#42C0FB", command=self.docsim)
        #Place the button in a specific position
        self.btn.place(x=919, y=95)
        #Place button inside the frame
        self.pack() 

    #Variable to change the number of clusters
    global kchange
    #Placeholder to check whether the value of k (number of clusters) has been requested to change
    kchange = 'the' 

    #Make function accessible from the main function
    global clusterno
    #Function to allow the user to change the number of clusters
    def clusterno():
        #Make variable accessible from the DNC class to check it's value
        global kchange 
        #Dialog box to allow user to type in the number of clusters they desire
        kchange = tkSimpleDialog.askstring("DN Choose the number of clusters", "Enter the number of clusters you want the 'Cluster Documents' feature to produce\n\nThis will be the maximum number of clusters that can be generated")

    #Make function accessible from the main function
    global defclustno
    #Function to allow the user to change the number of clusters back to the default calculation
    def defclustno():
        #Make variable accessible from the DNC class to check it's value
        global kchange
        #If user requests the number of clusters to be the default calculation, set the variable to the 'the', the placeholder
        kchange = 'the' 

    #Interact with the user to get the directory to cluster then call the clusterdn and runcluster functions
    def clusterdoc(self):
        #Make variable have a global scope
        global dname1 
        #Dialog box to allow the user to select a directory and store as the variable "dname1"
        dname1 = tkFileDialog.askdirectory(parent=root,title='Select the directory you want to cluster')
        #Check that there is a non-empty directory name stored in dname1
        if len(dname1) > 0:
            #Error handling
            try:
                #Call the clusterdn function with the directory as an argument
                clusterdn(dname1)
                #Call the runcluster function
                runcluster() 
                #Message box to inform the user that the clustering process is complete and a file is saved on the desktop
                tkMessageBox.showinfo("DN Clustered Documents", "Done! \n\nCheck your desktop for a .xls file named 'Clustered Documents DN...'")
            #Error handling
            except Exception: 
                #Below is for a message box to inform the user that the clustering process was not successful
                tkMessageBox.showwarning("DN Cluster Documents", "Unable to cluster directory\n\nPlease ensure the directory contains at least 1 .pdf, .docx, .txt or .csv file with text included")
            #Call the probe function with the directory name as an argument
            probe(dname1) 

    #Interact with the user to get the search term and directory to search in then call the smart1 and smart2 functions
    def search1(self):
        #Dialog box to allow the user to enter a search term and store as the variable "keyword1"
        keyword1 = tkSimpleDialog.askstring("DN Smart Search", "Enter the keyword you want to search")
        #If the user typed something into the fialog box then bring up a file dialog box for the user to select a directory and store as "dname2"
        if keyword1 != '': 
            dname2 = tkFileDialog.askdirectory(parent=root,initialdir="/",title='Select the directory you would like to search in')
            #Check that there is a non-empty directory name stored in dname2
            if len(dname2) > 0:
                #Error handling
                try:
                    #Call the smart1 function with the search term and directory as arguments
                    smart1(keyword1, dname2)
                    #Call the smart2 function with the search term and directory as arguments
                    smart2(keyword1, dname2) 
                    #Message box to inform the user that the smart search process is complete and a file is saved on the desktop
                    tkMessageBox.showinfo("DN Smart Search", "Done!\n\nCheck your desktop for a .xls file named 'Smart Search DN...'")
                #Error handling
                except Exception: 
                    #Message box to inform the user that the smart search process was not successful
                    tkMessageBox.showwarning("DN Smart Search", "Unable to process directory\n\nCheck the files and try again")
                #Call the probe function with the directory name as an argument
                probe(dname2) 

    #Interact with the user to get the search term and directory to search in then call the sensearch function          
    def search2(self): 
        #Dialog box to allow the user to enter a search term and store as the variable "keyword2"
        keyword2 = tkSimpleDialog.askstring("DN Sentence Search", "Enter the keyword you want to search for sentence by sentence")
        #If the user typed something into the fialog box then bring up a file dialog box for the user to select a directory and store as "dname3"
        if keyword2 != '': 
            dname3 = tkFileDialog.askdirectory(parent=root,initialdir="/",title='Select the directory you would like to search in') 
            #Check that there is a non-empty directory name stored in dname3
            if len(dname3) > 0:
                #Error handling
                try:
                    #Call the sensearch function with the search term and directory as arguments
                    sensearch(keyword2, dname3) 
                    #Message box to inform the user that the sentence search process is complete and a file is saved on the desktop
                    tkMessageBox.showinfo("DN Sentence Search", "Done!\n\nCheck your desktop for a .xls file named 'Sentence Search DN...'")
                #Error handling
                except Exception: 
                    #Message box to inform the user that the sentence search process was not successful
                    tkMessageBox.showwarning("DN Smart Search", "Unable to process directory\n\nCheck the files and try again")
                #Call the probe function with the directory name as an argument
                probe(dname3) 

    #Interact with the user to get 2 files and then call the sim1 and sim2 functions
    def docsim(self): 
        #Dialog box to allow the user to select a pdf, docx, txt or csv file
        filename1 = tkFileDialog.askopenfilename(parent=root,title='Select File 1 (.pdf, .docx, .txt or .csv)')
        #Check that there is a non-empty file name, non-empty file and the file is a pdf, docx, txt or csv file
        if len(filename1) > 0 and filename1.endswith('pdf') or filename1.endswith('docx') or filename1.endswith('txt') or filename1.endswith('csv') and not filename1==[]:
            #Dialog box to allow the user to select the 2nd pdf, docx, txt or csv file
            filename2 = tkFileDialog.askopenfilename(parent=root,title='Select File 2 (.pdf, .docx, .txt or .csv)')
            #Check that there is a non-empty file name, non-empty file and the file is a pdf, docx, txt or csv file
            if len(filename2) > 0 and filename2.endswith('pdf') or filename2.endswith('docx') or filename2.endswith('txt') or filename1.endswith('csv') and not filename2==[]:
                #Error handling
                try:
                    #Call the sim1 function with the 2 files as arguments
                    sim1(filename1, filename2)
                    #Call the sim2 function with the 2 files as arguments
                    sim2(filename1, filename2) 
                #Error handling
                except Exception: 
                    #Message box to inform the user that the document similarity process was not successful
                    tkMessageBox.showwarning("DN Document Similarity", "Unable to process the similarity between these files\n\nCheck the files and try again")
            else:
                #Message box to inform the user that the file they selected was empty or not a pdf, docx, txt or csv file
                tkMessageBox.showerror("DN Open File", "Unable to open file\n\nPlease select a file with a .pdf, .docx, .txt or .csv extension containing text")
        else:
            #Message box to inform the user that the file they selected was empty or not a pdf, docx, txt or csv file
            tkMessageBox.showerror("DN Open File", "Unable to open file\n\nPlease select a file with a .pdf, .docx, .txt or .csv extension containing text")


    #Make function accessible from the main function
    global askquit
    #Function to quit the application
    def askquit():
        #Dialog box to ask the user if they are sure they want to quit
        if tkMessageBox.askokcancel("Quit", "Are you sure you want to quit?"):
            #If user selects OK, quit the application - destroy the root window
            root.destroy() 

    #Make function accessible from the main function
    global about
    #Function to bring up an "About DN" window with info
    def about(): 
        #Info list will be shown in the new window giving the user information about the app
        info = [
              'Hi, Welcome to Data Ninja.',
              'This document will give you a brief overview of the purpose and use of this software.  Data Ninja is a document clustering software with document comparison, report generation and search capabilities.',
              ' ', 'Section 1', 'The menu is displayed once the application is opened.  This consists of the following buttons "Cluster Documents", "Smart Search", "Sentence Search" and "Document Similarity".',
              ' ', '1.1 Cluster Documents', 'This feature takes documents as input and outputs an Excel spreadsheet displaying clusters of documents, with each cluster containing documents that are similar to each other.',
              ' ', 'How to use:', 'Click the "Cluster Documents" button and select the directory which contains the documents that you want to cluster, click OK.  The software will begin processing the documents, during this time, if you try ',
              'to move the mouse icon away from the buttons you should see a round loading icon.  Depending on the number of documents, length and types of documents, the speed of processing will vary.  Once ',
              'processing is complete, an information box will appear stating "Done!  Check your desktop...".  An Excel spreadsheet with the processed output will appear on your desktop.  The first sheet in this file displays ',
              'the clusters of documents as hyperlinks.  The 2nd sheet displays documents that are anomalies i.e. documents that are not particularly similar to the documents in any one cluster group.  This sheet may ',
              'also include documents that could not be processed by the software, however this is rare.', ' ', 'To customise this feature see section 2 below.', ' ',
              'Scope - This feature will only cluster .pdf, .docx, .txt and .csv files.  Pdf files take longer to process.  Directories containing none of the acceptable files will not be clustered.',
              ' ', '1.2 Smart Search', 'This feature takes a search term and documents as input and outputs an Excel spreadsheet displaying all documents containing the search term and similar documents to these.',
              ' ', 'How to use:', 'Click the "Smart Search" button and enter the search term you want to search for, click OK.  Next select the directory you want to search for this term in and click OK.  The software will begin processing the',
              'documents, during this time, if you try to move the mouse icon away from the buttons you should see a round loading icon.  Depending on the number of documents, length and types of documents, the speed',
              'of processing will vary.  Once processing is complete, an information box will appear stating "Done!  Check your desktop...".  An Excel spreadsheet with the processed output will appear on your desktop.',
              'The first sheet in this file displays hyperlinks to documents containing the search term.  Next to each of these hyperlinks is a list of the most similar documents to that document from the same directory.  The',
              'second sheet in this file displays a general overview of similar documents to the documents containing the search term shown on the first sheet.', ' ',
              'Scope - This feature will only search .pdf, .docx, .txt and .csv files.  Pdf files take longer to process.  This feature takes significantly longer to process than the "Cluster Documents" process.  This feature literally',
              'searches for the search term in the documents specified i.e. if you enter "business" as the search term it will not find documents containing the term "Business", it will only find those containing "business".', ' ',
              '1.3 Sentence Search', 'This feature takes a search term, documents and search term synonyms as input and outputs an Excel spreadsheet displaying all documents containing the search term and the sentence containing the',
              'search term or synonyms/plurals found.', ' ', 'How to use:',
              'Click the "Sentence Search" button and enter the search term you want to search for, click OK.  Next select the directory you want to search for this term in and click OK.  Then enter the synonyms of the',
              'search term that you want the software to also search for and click OK.  The software will begin processing the documents, during this time, if you try to move the mouse icon away from the buttons',
              'you should see a round loading icon.  Depending on the number of documents, length and types of documents, the speed of processing will vary.  Once processing is complete, an information box will',
              'appear stating "Done!  Check your desktop...".  An Excel spreadsheet with the processed output will appear on your desktop.  This file will display all documents containing the search term as hyperlinks',
              'and the sentence containing the search term or synonyms/plurals found.', ' ',
              'Scope - This feature will only search .pdf, .docx, .txt and .csv files.  Pdf files take longer to process.  This feature takes significantly longer to process than the "Cluster Documents" process.  This feature does',
              'not limit its search to the literal search term entered by you.  It searches for the search term specified, the plurals of this term, the synonyms entered by you and both the capitalised and lower versions of all',
              'of these.', ' ', '1.4 Document Similarity', 'This feature takes two documents as input and outputs a new window in the application stating the similarity value between these two documents.',
              ' ', 'How to use:', 'Click the "Document Similarity" button and select 1 of the files you want to compare and click OK.  Then select the 2nd file you want to compare and click OK.  The software will begin processing the documents,',
              'during this time, if you try to move the mouse icon away from the buttons you should see a round loading icon.  Depending on the length and types of documents, the speed of processing will vary.  Once',
              'processing is complete, a new window will appear in the application describing the similarity between the 2 documents.', ' ',
              'Scope - This feature will only process.pdf, .docx, .txt and .csv files.  Pdf files take longer to process.  This feature is generally very quick to process documents.  Files that do not have an acceptable extension will',
              'not be processed.', ' ', ' ', 'Section 2', 'On the toolbar are "File" and "Help" tabs.  The "File" tab consists of the "Report Generation", "Customise Number of Clusters", "Default Number of Clusters" and "Quit".  The "Help" tab consists of',
              'the "About" option which opens this screen.', ' ', '2.1 Report Generation', 'This feature takes an audit report as input and outputs an insight log and draft management letter with information pulled from the audit report.',
              ' ', 'How to use:', 'Click the "Report Generation" option under the "File" tab in the toolbar.  Select an audit report and click OK.  Enter the auditors company name and click OK, enter the client companys name for audit and click',
              'OK and enter the year of audit and click OK.  The software will begin processing the documents, during this time, if you try to move the mouse icon away from the buttons you should see a round loading icon.',
              'Depending on the length and types of documents, the speed of processing will vary.  Once processing is complete, an information box will appear stating "Done!  Check your desktop...".  An Excel spreadsheet',
              'named in the format "Auditors name Insight Log - Client name Audit Year.xls" and a Word file (.doc) named in the format "Auditors name Draft Management Letter - Client name Audit Year.doc" with the',
              'processed output will appear on your desktop.', ' ',
              'Scope - This feature will only accept .pdf and .docx  audit reports.  Pdf files take longer to process.  This feature is generally very quick to process documents.  In order to be able to process the draft',
              'management letter in one of the other features you need to re-save it as a .docx or .pdf file.  Reports that do not have an acceptable extension will not be processed.',
              ' ', 'For optimal accuracy of the insight log and draft management letter, the audit reports selected should only use full stops to end a sentence and at the end of each topic e.g. "Network Security.".  Also, insights',
              'in the report should be labelled in the following manner and order: "InsightDN - There is no firewall installed.  ContactDN - Sam Daniel (Network Administrator).  MitigationDN - There are plans for a',
              'firewall to be installed in the "Proposals Autumn 2013 document".  RaiseDN - Yes.  AreaDN - Networks."  DN refers to Data Ninja.  InsightDN refers to the finding, ContactDN refers to the contact',
              'this finding was discussed with,  MitigationDN refers to any mitigations of this finding, RaiseDN refers to whether this finding should be included in the draft management letter (this is a Yes or No answer) and',
              'AreaDN refers to the section of the report the finding belongs to.', ' ', '2.2 Customise Number of Clusters', 'This feature allows the user to specify an estimate of the number of clusters to be produced by the "Cluster Documents" feature.',
              ' ', 'How to use:', 'Click the "Customise Number of Clusters" option under the "File" tab in the toolbar.  Enter the number of clusters you want the "Cluster Documents" feature to produce and click OK.  Then proceed to cluster',
              'documents as described in section 1.1.  To change the number of clusters again just repeat the above steps.', ' ',
              'Scope - For each dataset there is a limit to which the number of clusters can be customised i.e. a dataset of 3 documents may not be able to cluster with the number of clusters set to 20.',
              'In cases like this, an error message box will appear explaining that the software was unable to cluster the dataset.', ' ', '2.3 Default Number of Clusters',
              'This feature returns the number of clusters settings specified for the "Cluster Documents" feature to its original setting.  This is a rule of thumb formula that calculates the estimate of the number of clusters',
              'based on the number of documents.', ' ', 'How to use:', 'Click the "Default Number of Clusters" option under the "File" tab in the toolbar. Then proceed to cluster documents as described in section 1.1.',
              ' ', '2.4 Quit', 'This feature quits the application.', ' ', 'How to use:', 'Click the "Quit" option under the "File" tab in the toolbar.  A prompt will appear asking you if you are sure you want to quit.  If you click "OK" the application will quit.',
              ' ', 'General Scope',
              'To uniquely identify the Excel spreadsheets produced, each spreadsheet has the hour and minute it was created included in its name.  However, this means if you process the exact same directory within the',
              'same minute, it is likely this spreadsheet will be overridden or if this spreadsheet it currently open, the directory will not be processed.', ' ',
              'If the application is not responsive while processing, leave it for a while as often times it becomes responsive again', ' ',
              'Whenever a directory is uploaded onto the software, if a Pdf or Docx company audit report is found within that directory, a probe will appear to generate an insight log and draft management letter from',
              'this report.  If you accept this probe, an insight log and draft management letter will be produced and appear on your desktop.', ' ', ' ', ' ', ' ', 'Thank you for using Data Ninja'
              ' ', ' ', ' ', ' ', ' ', ' ',
              
              ]


        #Create a new window with root as the parent
        dn_about_us=Toplevel(root)
        #Automatically maximise the window
        dn_about_us.state("zoomed")
        #Set the title next to the logo to "Data Ninja - About DN"
        dn_about_us.title("Data Ninja - About DN") 
        """Create a listbox to place info inside and a vertical scrollbar to allow the user to scroll up and down the listbox and place this
        inside the new window"""
        scrollbar = Scrollbar(dn_about_us, orient=VERTICAL, width=610) 
        listbox = Listbox(dn_about_us, fg="#0000EE", bg="gray96", relief="flat", width=200, height=35, yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=LEFT, fill=Y, expand=1)
        scrollbar.place(x=600, y=15)
        listbox.pack(side=LEFT, fill=X, expand=1)
        listbox.place(x=10, y=80)
        label = Label(dn_about_us, text="About DN", bg="gray96", fg="#00009C", font=("Times New Roman", 30))
        label.pack(side=LEFT, fill="both", expand=1)
        label.place(x=10, y=10)
        for i in info:
            #Add each item in info to the listbox
            listbox.insert(END, i) 

#Function to read and clean pdf files
def readPdf(path):  
    content = ""
    #Open pdf file in read mode
    pdf = pyPdf.PdfFileReader(file(path, "rb"))
    #Iterate pages
    for i in range(0, pdf.getNumPages()):
        #Get text from each page
        content += pdf.getPage(i).extractText() + "\n"
    #Remove whitespaces
    content = " ".join(content.replace(u"xa0", " ").strip().split())
    #Variable containing all the words in the pdf file
    return content

#Function to read and clean docx files
def readDocx(path):
    #Unzip the zipfile in the docx
    docx = zipfile.ZipFile(path)
    #Read the xml file from the unzipped file and store in the "content" variable
    content = docx.read('word/document.xml')
    #Replace html tags and other symbols shown with nothing, to clean the content
    cleaned = re.sub('<(.|\n)*?>','',content)
    #Variable containing all the words in the docx file
    return cleaned 



#-------------------------------------------------------------------------------------------------------Start of cluster process for the cluster button


#Function to open files in the directory and add all the words in each document to a list
def clusterdn(dname1):
    #All words in each file are appended to "documents" as a list item
    global documents
    documents = []
    #Holds all pdf, docx, txt and csv files
    global no_of_docs
    no_of_docs = []
    #Dictionary with file names as keys and the most common word in the file as the value
    global f_tags
    f_tags = {}

    #Get all files from the directory including files in sub directories
    for directory, subdirectories, files in os.walk(dname1):
        #If there are no files in a folder, go to the next folder
        if len(files)<1:
            continue
        else:
            for f in files:
                    global fileX
                    #Get file path
                    fileX = os.path.join(directory, f)
                    #Check file extension
                    if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
                        #Error handling
                        try:
                            fileX = str(fileX)
                        #Error handling
                        except Exception:
                            continue
                        #Add file name to the no_of_docs list
                        no_of_docs.append(fileX)
                        #If file is a pdf, call the readPdf function to read it
                        if fileX.endswith('pdf'):
                            #Error handling
                            try:
                                words1 = readPdf(fileX)
                            #Error handling
                            except Exception:
                                continue
                        #If file is a docx, call the readDocx function to read it
                        elif fileX.endswith('docx'):
                            words1 = readDocx(fileX)
                        else:
                            #If file is a txt or csv open the file and read all words into variable "words1"
                            file1 = open(fileX, 'r')
                            words1 = file1.read()
                            words1 = str(words1)
                        #If the file is empty go to the next file
                        if words1 == '':
                            continue
                        #If the file isn't empty, make all words be in the lowercase and discard stopwords and words that do not contain letters of the alphabet
                        else:
                            t_count = words1.split();
                            t_count = [x.lower() for x in t_count]
                            t_count = [x for x in t_count if x.isalpha() == True and x not in stopwords]
                            #If there are no words left, go to the next file
                            if t_count == []:
                                continue
                                """If there are word left, count how many times each word occurs, then select the word that occurs the most and
                                add this to the "f_tag" list as the value to the key which is the file name"""
                            else:
                                tags = nltk.FreqDist(t_count)
                                tag = tags.max()
                                tag = str(tag)
                                f_tags[fileX] = tag
                                #Add the words extracted from this file to the list "documents"
                                documents.append(words1)
                    #If the file is not a pdf, txt, docx or csv file, go to the next file     
                    else:
                        continue
    #doc_no stores the number of documents that are pdfs, csv, txt or docx from the directory
    doc_no = 0
    for i in no_of_docs:
        doc_no += 1

#Data Ninja Clustering class
class DNC(object):
    #initialise objects
    def __init__(self, stopwords):
        #dict is a dictionary with words as keys and the files and number of times this words has occurred e.g. {'hazardous': [0,0,1,1,1,2]...}
        self.dict = {}
        #dno is the document number i.e. the first file in is file 0, second file is file 1 etc.
        self.dno = 0

    def analyse(self, string):
        #Split document into words
        terms1 = string.split();
        #Make all words lower case and discard stopwords
        terms1 = [t.lower() for t in terms1 if t not in stopwords]
        #Only return words containing letters of the alphabet
        for t in terms1:
            t = ''.join(i for i in t if  i in 'qwertyuiopasdfghjklzxcvbnm ')
            #Add document number as a value to the word key in the dictionary if the word key is already in the dictionary 
            if t in self.dict:
                self.dict[t].append(self.dno)
            #Get rid of empty strings as a result from parsing
            elif t != '':
                #Add the word key and the document number value to the dictionary
                self.dict[t] = [self.dno]
        #Increment document number
        self.dno += 1
                    
        
    def matrix(self):
        #Get all word keys which have more than 1 file name value (i.e. words appearing in more than 1 document)
        self.keys = [k for k in self.dict.keys() if len(self.dict[k]) > 1]
        self.keys.sort()
        #Create an array with the number of keys and values in the dictionary as the dimensions for the array
        self.cmatrix = zeros([len(self.keys), self.dno])
        #Get the index(i) and the word key(x) for all word keys e.g. 0 hazardous, 1 bag, 2 tea etc.
        for i, x in enumerate(self.keys):
            #For every value(file number) for the word key x in the dictionary...
            for n in self.dict[x]:
                """index(i) refers to each row of the count matrix and each file number is a column in the count matrix, and for each time
                the file number occurs as a value for a word key, increment this number - which is how many times the word occurs in each file"""
                self.cmatrix[i,n] += 1

        #Get each column of the count matrix and store in vect1 as a vector for each file number
        vect1 = zip(*self.cmatrix)

        #Get the number of vectors in vect1 which corresponds to the number of files
        f_no = len(vect1)

        #If kchange is equal to 'the' this means the default calculation for k should be used
        if kchange == 'the':
            #The rule of thumb calculation of k i.e. square root of (number of objects /2)
            #Added +1 because sqrt automatically rounds down and I want figures to round up e.g. if 1.2 I want 2 clusters (the result is almost always a decimal before rounding)
            k = math.sqrt(f_no/2)+1
            #Get rid of decimal
            k = int(k)
        #Else the user's specified k should be used
        else:
            #Get rid of decimal for the user's requested k number
            k = int(kchange)

        #Turn every vector in vect into an array, as this is the format the cluster library from nltk takes as input     
        vect2 = [array(i) for i in vect1]

        #Use the GAAClusterer class from the cluster module in nltk with k as an argument and the normalise feature set to True, and store as clusterX
        clusterX = cluster.GAAClusterer(k, True)
        #Cluster the vectors stored as arrays in vect2, with the assign_clusters feature set to True, and stored as clusters
        clusters = clusterX.cluster(vect2, True)
        #Create a dictionary with the list of file names as the key and the list of cluster membership of each vector array
        membership = (dict(zip(no_of_docs, clusters)))

        """For each filename in membership, if the filename is in the f_tags dictionary, store tag1 as the values of the membership
        key and filename key and append to the d_tags list as an item e.g. [(2, 'business'),..]"""
        d_tags = []
        for key in membership:
            if key in f_tags:
                tag1 = membership[key], f_tags[key]
                d_tags.append(tag1)

        #Remove repetition in the d_tags dictionary e.g. turn {0: ['business, business, business, finance]..} to {0: ['business, finance']..}
        clust_tags={}
        for key, val in d_tags:
            #Search for the value in the list of values and if it's already there just continue
            if val in clust_tags.setdefault(key,[]): 
                continue
            #If it is not there, then add the value to the list of values
            else:
                clust_tags.setdefault(key, []).append(val) #

        #Create an xls workbook with the following text and formatting
        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Clustered Documents")
        sheet2 = book.add_sheet("Anomalous Documents")
        style = easyxf('font: underline single, color sky_blue;')
        style3 = easyxf('font: bold 1, italic 1, color sky_blue;') 
        style2 = easyxf('font: name Times New Roman size 05 bold on, color sky_blue, height 450;')
        style4 = easyxf('pattern: back_colour sky_blue, pattern thick_forward_diag, fore-colour sky_blue;')
        style5 = easyxf('font: italic 1, color 24;')
        sheet1.write_merge(0,1,0,10, "Clustered Documents", style2)
        sheet2.write_merge(0,1,0,20, "Anomalous Documents", style2)
        sheet1.write(3,0, "Categories:", style3)
        sheet2.write(3,0, "Documents:", style3)
        sheet1.write(6,0, "*Each doc is", style5)
        sheet1.write(7,0, "a hyperlink", style5)
        sheet2.write(6,0, "*This sheet may", style5)
        sheet2.write(7,0, "also contain", style5)
        sheet2.write(8,0, "docs that could", style5)
        sheet2.write(9,0, "not be processed", style5)
        sheet2.write(10,0, "but this is rare", style5)
        sheet1.write(2,0, "", style4)
        sheet1.write(2,1, "", style4)
        sheet2.write(2,0, "", style4)
        sheet2.write(2,1, "", style4)
        sheet1.write(4,0, "Documents", style)

        #Create vertical light blue separators
        format1 = range(3,200)
        for i in format1:
            sheet1.write(i,1, "", style4)
            sheet2.write(i,1, "", style4)
            
        #Create horizontal light blue separators
        format2 = range(2,200)
        for i in format2:
            sheet2.write(2,i, "", style4)
        
        #Write document tag terms to the xls file   
        y1=2
        for key in clust_tags:
            #Convert list to string with each item separated with a comma
            dn_tags = ', '.join(clust_tags[key])
            dn_tags = dn_tags.title()
            sheet1.write(3,y1,dn_tags, style3)
            sheet1.write(2,y1, "", style4)
            y1 += 1            

        #Try to write to each cell in a column and handle the error and repeat onto there is a free cell where it will write the hyperlink
        def check(x2, y2):   
            try:
                #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes
                doc_path = r''+str(doc1).replace('\\', '/')
                doc_path = str(doc_path)
                #Get the file name and extension name
                doc_name = os.path.basename(doc_path)
                doc_name = str(doc_name)
                doc_name = doc_name.title()
                docX = 'HYPERLINK('+'"'+doc_path+'"'+';"'+doc_name+'")'
                sheet1.write(x2,y2,Formula(docX),style)
            #Error handling
            except Exception:
                #Recheck the next cell in the column
                x2 += 1
                check(x2, y2)

        #Write file names as hyperlinks underneath the correct categories to sheet 1
        for key in membership:
            x2=3
            doc1 = str(key)
            #membership[key] is the cluster membership number, +2 is to position it under the correct category in the xls file
            y2 = membership[key]+2
            x2 += 1
            check(x2, y2)

        #This writes file names as hyperlinks to sheet 2
        x1 = 3
        for i in no_of_docs:
            if i not in membership:
                #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes
                a_path = r''+str(i).replace('\\', '/')
                a_path = str(i)
                #Get the file name and extension name
                a_name = os.path.basename(a_path)
                a_name = str(a_name)
                a_name = a_name.title()
                anomaly = 'HYPERLINK('+'"'+a_path+'"'+';"'+a_name+'")'
                sheet2.write(x1,2,Formula(anomaly), style)
                x1 += 1

        #Set column width 
        sheet1.col(0).width = 256*12
        sheet1.col(1).width = 256*03
        sheet2.col(0).width = 256*16
        sheet2.col(1).width = 256*03
        sheet2.col(2).width = 256*50
        count = 2
        #Adjust the column width for each category
        for i in clust_tags:
            sheet1.col(count).width = 256*80
            count += 1
        #Save file on the desktop with the current date time included in the xls file name
        dt = time.strftime("%Y-%m-%d %H.%M")
        folder = os.path.basename(dname1)
        folder = folder.title()
        file_name = "Clustered Documents DN - "+str(folder)+" Folder "+str(dt)+".xls"
        s_file = dpath+"/"+file_name
        book.save(s_file)

#Call the DNC class with stopwords as an argument and call the analyse function for each document and call the matrix function on all documents in one go
def runcluster():   
    dncluster = DNC(stopwords)
    for d in documents:
        dncluster.analyse(d)
    dncluster.matrix()


    
#-------------------------------------------------------------------------------------------------------End of cluster process for the cluster button
#-------------------------------------------------------------------------------------------------------Start of smart search process for the smart search button



#Function to check which files contain the keyword and write these files as hyperlinks an xls spreadsheet
def smart1(keyword1, dname2):
    #Make this variable accessible from the smart2 function, this variable stores the file names that contain the keyword
    global s_result
    s_result = []
    x1 = 4
    #globals are there to make these variables accessible from the smart2 function
    global book
    #Create an xls workbook with the following text and formatting
    book = xlwt.Workbook(encoding="utf-8")
    global sheet1
    sheet1 = book.add_sheet("S.Search Results")
    global sheet2
    sheet2 = book.add_sheet("Similar Results")
    global style
    f_name1 = os.path.basename(dname2)
    style = easyxf('font: underline single, color sky_blue;')
    style3 = easyxf('font: bold 1, color sky_blue;')
    style2 = easyxf('font: name Times New Roman size 05 bold on, color sky_blue, height 450;')
    style4 = easyxf('pattern: back_colour sky_blue, pattern thick_forward_diag, fore-colour sky_blue;')
    style5 = easyxf('font: name Arial size 05, color sky_blue;')
    global style6
    style6 = easyxf('font: italic 1, color 24;')
    sheet1.write_merge(0,1,0,60, "Search results for '"+str(keyword1)+"' from the "+str(dname2)+" directory", style2)
    sheet2.write_merge(0,1,0,30, "Generally similar documents to those produced in the 'S.Smart Search' sheet from the same directory", style2)
    sheet2.write(2,0, "", style4)
    sheet2.write(2,1, "", style4)
    sheet1.write(2,0, "", style4)
    sheet1.write(2,1, "", style4)
    sheet1.write(3,0, "*Each doc is", style6)
    sheet1.write(4,0, "a hyperlink", style6)
    sheet1.write(6,0, "*Similarity values are", style6)
    sheet1.write(7,0, "measured in terms of", style6)
    sheet1.write(8,0, "the cosine angle", style6)
    sheet1.write(9,0, "between 2 documents", style6)
    sheet1.write(10,0, "on a scale of -1 to 1:", style6)
    sheet1.write(11,0, "-1 meaning opposite", style6)
    sheet1.write(12,0, "0 meaning independent", style6)
    sheet1.write(13,0, "& 1 meaning duplicate", style6)
    sheet1.write(14,0, "i.e. the closer the", style6)
    sheet1.write(15,0, "value is to 1, the more", style6)
    sheet1.write(16,0, "similar the documents", style6)
    sheet1.write(3,2,"Documents containing the term "+str(keyword1), style3)
    sheet1.write_merge(3,3,4,20, "The most similar documents from the "+str(f_name1)+" directory", style3)


    #Create vertical light blue separators
    format1 = range(3,200)
    for i in format1:
        sheet1.write(i,1, "", style4)
        sheet1.write(i,3, "", style4)
        sheet2.write(i,1, "", style4)
        sheet2.write(i,3, "", style4)

    #Create horizontal light blue separators
    format2 = range(2,200)
    for i in format2:
        sheet1.write(2,i, "", style4)
        sheet2.write(2,i, "", style4)

    #Get all files from the directory including files in sub directories
    for directory, subdirectories, files in os.walk(dname2): 
            for f in files:
                #Get file path
                fileX = os.path.join(directory, f)
                #Check file extension
                if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
                    fileX = str(fileX)
                    #If file is a pdf, call the readPdf function to read it
                    if fileX.endswith('pdf'):
                        #Error handling
                        try:
                            words1 = readPdf(fileX)
                        #Error handling
                        except Exception:
                            continue
                    #If file is a docx, call the readDocx function to read it
                    elif fileX.endswith('docx'):
                        words1 = readDocx(fileX)
                    #If file is a txt or csv open the file and read all words into variable "words1"
                    else:
                        file2 = open(fileX, 'r')
                        words1 = file2.read()
                    #Error handling
                    try:
                        words1 = str(words1)
                    #Error handling
                    except Exception:
                            continue
                    #This writes file names containing the keyword, as hyperlinks to sheet 1
                    if keyword1 in words1:
                        #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes
                        fileX = r''+str(fileX).replace('\\', '/')
                        fileX = str(fileX)
                        #Get the file name and extension name
                        r_name = os.path.basename(fileX)
                        r_name = str(r_name)
                        r_name = r_name.title()
                        result = 'HYPERLINK('+'"'+fileX+'"'+';"'+r_name+'")'
                        sheet1.write(x1,2,Formula(result), style)
                        x1 += 1
                        s_result.append(fileX)
                    else:
                        continue
                #If the file is not a pdf, txt, docx or csv file, go to the next file 
                else:
                    continue


#Smart Search class for the smart2 function
class SmartS(object):
    #initialise objects
    def __init__(self, stopwords):
        #dict is a dictionary with words as keys and the files and number of times this words has occurred e.g. {'hazardous': [0,0,1,1,1,2]...}
        self.dict = {}
        #dno is the document number i.e. the first file in is file 0, second file is file 1 etc.
        self.dno = 0

    def analyse(self, string):
        #Split document into words
        terms1 = string.split();
        #Only return words containing letters of the alphabet
        for t in terms1:
            #Add document number as a value to the word key in the dictionary if the word key is already in the dictionary 
            if t in self.dict :
                self.dict[t].append(self.dno)
            #Get rid of empty strings as a result from parsing
            elif t != '':
                #Add the word key and the document number value to the dictionary
                self.dict[t] = [self.dno]
        #Increment document number
        self.dno += 1
                    
        
    def matrix(self):
        #Get all word keys which have more than 1 file name value (i.e. words appearing in more than 1 document)
        self.keys = [k for k in self.dict.keys() if len(self.dict[k]) > 1] 
        self.keys.sort()
        #Create an array with the number of keys and values in the dictionary as the dimensions for the array
        self.cmatrix = zeros([len(self.keys), self.dno])
        #Get the index(i) and the word key(x) for all word keys e.g. 0 hazardous, 1 bag, 2 tea etc.
        for i, x in enumerate(self.keys):
            #For every value(file number) for the word key x in the dictionary...
            for n in self.dict[x]:
                """index(i) refers to each row of the count matrix and each file number is a column in the count matrix, and for each time
                the file number occurs as a value for a word key, increment this number - which is how many times the word occurs in each file"""
                self.cmatrix[i,n] += 1

        #Get each column of the count matrix and store in vect1 as a vector for each file number
        global vect1
        vect1 = zip(*self.cmatrix)
        #Get the first two vectors from vect1 and store as u and v respectively
        u = vect1[0]
        v = vect1[-1]
        #Make cosine distance (cd) accessible from the smart2 function
        global cd
        #Calculate the cosine distance between the two vectors
        cd = (numpy.dot(u, v) / (math.sqrt(numpy.dot(u, u)) * math.sqrt(numpy.dot(v, v))))

#Function to write the file names of files that are similar to those produced by the smart1 function, to the xls file
def smart2(keyword1, dname2):
    x1 = 3
    #For each filename outputted by the smart1 function
    for i in s_result:
        y1 = 4
        x1 += 1
        fileY = str(i)
        folder = dname2
        #Get all files from the directory including files in sub directories
        for directory, subdirectories, files in os.walk(folder):
            #If there are no files in a folder, go to the next folder
            if len(files)<1:
                continue 
            else:
                """For every file in the directory, open and extract the words from the file in s_result and append to the documents list"""
                for f in files:
                    documents = []      
                    #Check file extension
                    if fileY.endswith('pdf') or fileY.endswith('docx') or fileY.endswith('txt') or fileY.endswith('csv'):
                        fileY = str(fileY)
                        #If file is a pdf, call the readPdf function to read it
                        if fileY.endswith('pdf'):
                            #Error handling
                            try:
                                words2 = readPdf(fileY)
                            #Error handling
                            except Exception:
                                continue
                        #If file is a docx, call the readDocx function to read it
                        elif fileY.endswith('docx'):
                            words1 = readDocx(fileY)
                        #If file is a txt or csv open the file and read all words into variable "words1"
                        else:
                            file2 = open(fileY, 'r')
                            words1 = file2.read()
                        #Error handling
                        try:
                            words1 = str(words1)
                        #Error handling
                        except Exception:
                            continue
                        #Add the words extracted from this file to the list "documents"
                        documents.append(words1)
                    #If the file is not a pdf, txt, docx or csv file, go to the next file 
                    else:
                        continue
                    """Then extract the words from the file in the directory and append to the documents list"""
                    #Get file path
                    fileX = os.path.join(directory, f)
                    #Check file extension
                    if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
                        fileX = str(fileX)
                        #If file is a pdf, call the readPdf function to read it
                        if fileX.endswith('pdf'):
                            #Error handling
                            try:
                                words1 = readPdf(fileX)
                            #Error handling
                            except Exception:
                                continue
                        #If file is a docx, call the readDocx function to read it
                        elif fileX.endswith('docx'):
                            words1 = readDocx(fileX)
                        #If file is a txt or csv open the file and read all words into variable "words1"
                        else:
                            file2 = open(fileX, 'r')
                            words1 = file2.read()
                        #Error handling
                        try:
                            words1 = str(words1)
                        #Error handling
                        except Exception:
                            continue
                        #Add the words extracted from this file to the list "documents"
                        documents.append(words1)
                    #If the file is not a pdf, txt, docx or csv file, go to the next file
                    else:
                        continue
                    #Check that there are 2 documents in the documents list
                    if len(documents) == 2:
                        #Call the SmartS class with stopwords as an argument and call the analyse function for each document and call the matrix function on both in one go
                        dnsmart = SmartS(stopwords)
                        for d in documents:
                            dnsmart.analyse(d)
                        dnsmart.matrix()
                        #This writes file names as hyperlinks to sheet 1 of all files that have a 0.5 or greater similarity from the directory for each file outputted by the smart1 function
                        if cd >0.5:
                            sim_result1 = str(f)+ " is " + str(cd)+ " similar to " + fileY    
                            #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes
                            f = r''+str(f).replace('\\', '/')
                            f = str(f)
                            #Get the file name and extension name
                            sim_docs = os.path.basename(f)
                            sim_docs = str(sim_docs)
                            sim_docs = sim_docs.title()
                            #Get the cosine distance to 2 decimal places
                            cd1 = "{0:.2f}".format(cd)
                            sim_docs = sim_docs+" has a similarity value of "+str(cd1)+" with this file"
                            sheet1.write(x1,y1, sim_docs, style6)
                            y1 += 1
                    else:
                       continue
                    
    #List of files that are generally similar to the files outputted by the smart1 function
    gen_sim = []
    #For each filename outputted by the smart1 function
    for fileZ in s_result:
        #List of file names and their cosine similarity distance to fileZ
        max1 = {}
        #Get all files from the directory including files in sub directories
        for directory, subdirectories, files in os.walk(dname2):
            """For each file in s_result open and extract the words from the file and append to the documents list"""
            for f in files:
                    documents = []
                    #Check file extension
                    if fileZ.endswith('pdf') or fileZ.endswith('docx') or fileZ.endswith('txt') or fileZ.endswith('csv'):
                            fileZ = str(fileZ)
                            #If file is a pdf, call the readPdf function to read it
                            if fileZ.endswith('pdf'):
                                #Error handling
                                try:
                                    words1 = readPdf(fileZ)
                                #Error handling
                                except Exception:
                                    continue
                            #If file is a docx, call the readDocx function to read it
                            elif fileZ.endswith('docx'):
                                words1 = readDocx(fileZ)
                            #If file is a txt or csv open the file and read all words into variable "words1"
                            else:
                                file2 = open(fileZ, 'r')
                                words1 = file2.read()
                            #Error handling
                            try:
                                words1 = str(words1)
                            #Error handling
                            except Exception:
                                continue
                            #Add the words extracted from this file to the list "documents"
                            documents.append(words1)
                    #If the file is not a pdf, txt, docx or csv file, go to the next file
                    else:
                        continue
                    """For every file in the directory, open and extract the words from the file (if the file is not in s_result) and append to the documents list"""
                    #Ensure files that were outputted by the smart1 function are not outputted again as part of gen_sim
                    if f not in s_result:
                        #Get file path
                        fileX = os.path.join(directory, f)
                        #Check file extension
                        if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
                            fileX = str(fileX)
                            #If file is a pdf, call the readPdf function to read it
                            if fileX.endswith('pdf'):
                                #Error handling
                                try:
                                    words1 = readPdf(fileX)
                                #Error handling
                                except Exception:
                                    continue
                            #If file is a docx, call the readDocx function to read it
                            elif fileX.endswith('docx'):
                                words1 = readDocx(fileX)
                            #If file is a txt or csv open the file and read all words into variable "words1"
                            else:
                                file2 = open(fileX, 'r')
                                words1 = file2.read()
                            #Error handling
                            try:
                                words1 = str(words1)
                            #Error handling
                            except Exception:
                                continue
                            #Add the words extracted from this file to the list "documents"
                            documents.append(words1)
                        #If the file is not a pdf, txt, docx or csv file, go to the next file
                        else:
                            continue
                    #If the file is in s_result then continue
                    else:
                        continue
                    #Check that there are 2 documents in the documents list
                    if len(documents) == 2:
                        #Call the SmartS class with stopwords as an argument and call the analyse function for each document and call the matrix function on both in one go
                        dnsmart = SmartS(stopwords)
                        for d in documents:
                            dnsmart.analyse(d)
                        dnsmart.matrix()
                        #Discard files that have a lower similarity than 0.3 to the output file from the smart1 function
                        if cd > 0.3:
                            sim_result1 = str(fileX)
                            #Update max1 with the file name and the cosine distance from fileZ
                            max1.update({sim_result1:cd})
                        else:
                            continue
                    else:
                        continue
   
        #Function to get the file with the highest similarity rate to fileZ
        def getSimDoc():
            #Get the key with the maximum value in max1
            maxD = max(max1, key=max1.get)
            #If the key is not already in gen_sim, add it to gen_sim #This is to ensure there isn't repetition of file names in gen_sim
            if maxD not in gen_sim:
                    gen_sim.append(maxD)
            #If the key is already in gen_sim, delete the key from max1 and get the new max from the list and repeat till a key that is not gen_sim is found
            else:
                del max1[maxD]
                getSimDoc()
        #Call the above function
        getSimDoc()

    #This writes the generally similar file names as hyperlinks to sheet 2
    x2 = 3
    for i in gen_sim:
        #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes 
        i = r''+str(i).replace('\\', '/')
        i = str(i)
        #Get the file name and extension name
        gen_sim_doc = os.path.basename(i)
        gen_sim_doc = str(gen_sim_doc)
        gen_sim_doc = gen_sim_doc.title()
        g_sim_r = 'HYPERLINK('+'"'+i+'"'+';"'+gen_sim_doc+'")'
        sheet2.write(x2,2,Formula(g_sim_r), style)
        x2 += 1

    #Set column width 
    sheet1.col(0).width = 256*20
    sheet1.col(1).width = 256*03
    sheet1.col(2).width = 256*50
    sheet1.col(3).width = 256*03
    sheet2.col(0).width = 256*05
    sheet2.col(1).width = 256*03
    sheet2.col(2).width = 256*170
    sheet2.col(3).width = 256*03
    count = range(4,20)
    for i in count:
        sheet1.col(i).width = 256*90
    #Save file on the desktop with the current date and time included in the xls file name
    dt = time.strftime("%Y-%m-%d %H.%M")
    keyword1 = keyword1.title()
    file_name1 = "Smart Search DN "+str(keyword1)+" "+str(dt)+".xls"
    ss_file = dpath+"/"+file_name1
    book.save(ss_file) 



#-------------------------------------------------------------------------------------------------------End of smart search process for the smart search button
#-------------------------------------------------------------------------------------------------------Start of sentence search process for the sentence search button


#Function to check which files contain the keyword, plurals and synonyms and write these files as hyperlinks and the sentences with these words to an xls spreadsheet
def sensearch(keyword2, dname3):
    keyword2 = str(keyword2)
    #List holding all terms to be searched for
    searchterms = []
    #Ensure both the capitalised and lower version of the keyword is searched for
    if keyword2[0].isupper() == True:
        kword = keyword2.lower()
    else:
        kword = keyword2.capitalize()
    searchterms.append(kword)
    #Dialog box to allow the user to enter in any synonyms to the keyword they want to also search for (separate each word with a comma)
    syns = tkSimpleDialog.askstring("DN Sentence Search Synonyms?", "Want to add synonyms?\n\nSeparate each synonym with a comma e.g. finance, accounts, banking")
    #Remove commas and split synonyms into phrases/words
    syns = syns.replace(",", "")
    syns = syns.split()
    
    #Get the plural of the keyword and ensure both the capitalised and lower version of the plural keyword is searched for
    p = inflect.engine()
    kp1 = p.plural(keyword2)
    if kp1[0].isupper() == True:
        kp2 = kp1.lower()
    else:
        kp2 = kp1.capitalize()
    searchterms.append(kp1)
    searchterms.append(kp2)

    #Ensure both the capitalised and lower version of the synonyms are searched for
    for i in syns:
        if i[0].isupper() == True:
            s_lower = i.lower()
        else:
            lo = i.capitalize()
        searchterms.append(i)
        searchterms.append(lo)

    #Get the plural of the synonyms and ensure both the capitalised and lower version of the plural synonyms are searched for
    for i in syns:
        p = inflect.engine()
        kp3 = p.plural(i)
        if kp3[0].isupper() == True:
            kp4 = kp3.lower()
        else:
            kp4 = kp3.capitalize()
        searchterms.append(kp3)
        searchterms.append(kp4)

    #Convert searchterms list to a string of phrases separated by commas
    syn_plu = ', '.join(searchterms)

    #Create an xls workbook with the following text and formatting
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Search Results")
    style = easyxf('font: underline single, color sky_blue;')
    style3 = easyxf('font: bold 1, color sky_blue;')
    style2 = easyxf('font: name Times New Roman size 05 bold on, color sky_blue, height 450;')
    style4 = easyxf('pattern: back_colour sky_blue, pattern thick_forward_diag, fore-colour sky_blue;')
    style5 = easyxf('font: name Arial size 05, color sky_blue;')
    style6 = easyxf('font: italic 1, color 24;')
    sheet1.write_merge(0,1,0,60, "Search results for the term '"+str(keyword2)+"' and the following synonyms and/or plurals - "+str(syn_plu)+ " from the "+str(dname3)+ " directory", style2)
    sheet1.write(2,0, "", style4)
    sheet1.write(2,1, "", style4)
    sheet1.write(3,2,"File Name", style3)
    sheet1.write(3,3,"Sentence found with the search term or synonym of the search term", style3)
    sheet1.write(3,4,"Search terms found in these sentences", style3)
    sheet1.write(6,0, "*Each doc is", style6)
    sheet1.write(7,0, "a hyperlink", style6)

    #Create vertical light blue separators
    format1 = range(3,200)
    for i in format1:
        sheet1.write(i,1, "", style4)

    #Create horizontal light blue separators
    format2 = range(2,200)
    for i in format2:
        sheet1.write(2,i, "", style4)

    
    x1 = 3
    #Add the keyword to searchterms, didn't add this before because I want to present syn_plu without the keyword (because those are synonyms and plurals
    searchterms.append(keyword2)

    #Get all files from the directory including files in sub directories
    for directory, subdirectories, files in os.walk(dname3):
        for f in files:
                #Get file path
                fileX = os.path.join(directory, f)
                #Check file extension
                if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
                    fileX = str(fileX)
                    #If file is a pdf, call the readPdf function to read it
                    if fileX.endswith('pdf'):
                        #Error handling
                        try:
                            words1 = readPdf(fileX)
                        #Error handling
                        except Exception:
                            continue
                    #If file is a docx, call the readDocx function to read it
                    elif fileX.endswith('docx'):
                        words1 = readDocx(fileX)
                    #If file is a txt or csv open the file and read all words into variable "words1"
                    else:
                        file1 = open(fileX, 'r')
                        words1 = file1.read()
                    #Error handling
                    try:
                        words1 = str(words1)
                    #Error handling
                    except Exception:
                        continue
                    #Split the document into sentences i.e. lines
                    lines = re.split(r'\s*[!?.]\s*', words1)
                    #For each line empty key_sent and for each search term if the search term is in the line, add the search term to the list key_sent
                    for i in lines:
                        key_sent=[]
                        for n in searchterms:
                            if n in i:
                                key_sent.append(n)
                            else:
                                continue
                        #If no search terms are found in the line, go to the next line
                        if key_sent == []:
                            continue
                        #This writes file names containing search terms as hyperlinks to sheet 1 and also the sentences with these search terms 
                        else:
                            x1 += 1
                            #xlwt hyperlinks accept paths with forward slashes, this replaces backward slashes with forward slashes
                            ks_path = r''+str(fileX).replace('\\', '/')
                            ks_path = str(fileX)
                            #Get the file name and extension name
                            ks_result = os.path.basename(ks_path)
                            ks_result = str(ks_result)
                            ks_result = ks_result.title()
                            k_sentence = 'HYPERLINK('+'"'+ks_path+'"'+';"'+ks_result+'")'
                            sheet1.write(x1,2,Formula(k_sentence), style)
                            #Search term found in the current sentence/line
                            key_words = ", ".join(key_sent)
                            #Error handling
                            try:
                                #Sentence containing search term
                                sheet1.write(x1,3,str(i), style5)
                                #Search terms found in this sentence
                                sheet1.write(x1,4,str(key_words), style5)
                            #Error handling
                            except Exception:
                                continue

    #Set column width      
    sheet1.col(0).width = 256*12
    sheet1.col(1).width = 256*03
    sheet1.col(2).width = 256*50
    sheet1.col(3).width = 256*110
    sheet1.col(4).width = 256*40
    #Save file on the desktop with the current date time included in the xls file name
    keyword2 = keyword2.title()
    dt = time.strftime("%Y-%m-%d %H.%M")
    file_name2 = "Sentence Search DN "+str(keyword2)+" "+str(dt)+".xls"
    ks_file = dpath+"/"+file_name2
    book.save(ks_file)



#-------------------------------------------------------------------------------------------------------End of sentence search process for the smart search button
#-------------------------------------------------------------------------------------------------------Start of doc similarity process for the doc similarity button


#Function to open both files selected and add all the words in each document to a list
def sim1(filename1, filename2):
    #Make this list accessible from the sim2 function
    global documents
    documents = []
    #Add both files to a list named files
    files = [filename1, filename2]
    for i in files:
        fileX = str(i)
        #Check file extension
        if fileX.endswith('pdf') or fileX.endswith('docx') or fileX.endswith('txt') or fileX.endswith('csv'):
            #If file is a pdf, call the readPdf function to read it
            if fileX.endswith('pdf'):
                words1 = readPdf(fileX)
            #If file is a docx, call the readDocx function to read it
            elif fileX.endswith('docx'):
                words1 = readDocx(fileX)
            #If file is a txt or csv open the file and read all words into variable "words1"
            else:
                doc = open(fileX, 'r')
                words1 = doc.read()
            words1 = str(words1)
            #Add the words extracted from this file to the list "documents"
            documents.append(words1)
        else:
            continue

#Document Similarity class to be accessed from the sim2 function
class SIM(object):
    #initialise objects
    def __init__(self, stopwords):
        #dict is a dictionary with words as keys and the files and number of times this words has occurred e.g. {'hazardous': [0,0,1,1,1,2]...}
        self.dict = {}
        #dno is the document number i.e. the first file in is file 0, second file is file 1 etc.
        self.dno = 0

    def analyse(self, string):
        #Split document into words
        terms1 = string.split();
        #Make all words lower case and discard stopwords
        terms1 = [t.lower() for t in terms1 if t not in stopwords]
        #Only return words containing letters of the alphabet
        for t in terms1:
            t = ''.join(i for i in t if  i in 'qwertyuiopasdfghjklzxcvbnm')
            #Add document number as a value to the word key in the dictionary if the word key is already in the dictionary
            if t in self.dict:
                self.dict[t].append(self.dno)
            #Get rid of empty strings as a result from parsing
            elif t != '':
                #Add the word key and the document number value to the dictionary
                self.dict[t] = [self.dno]
        #Increment document number
        self.dno += 1

    def matrix(self):
        #Get all word keys which have more than 1 file name value (i.e. words appearing in more than 1 document)
        self.keys = [x for x in self.dict.keys() if len(self.dict[x]) > 1] 
        self.keys.sort()
        #Create an array with the number of keys and values in the dictionary as the dimensions for the array
        self.cmatrix = zeros([len(self.keys), self.dno])
        #Get the index(i) and the word key(x) for all word keys e.g. 0 hazardous, 1 bag, 2 tea etc.
        for i, x in enumerate(self.keys):
            #For every value(file number) for the word key x in the dictionary...
            for n in self.dict[x]:
                """index(i) refers to each row of the count matrix and each file number is a column in the count matrix, and for each time
                the file number occurs as a value for a word key, increment this number - which is how many times the word occurs in each file"""
                self.cmatrix[i,n] += 1

        #Get each column of the count matrix and store in vect1 as a vector for each file number
        vect1 = zip(*self.cmatrix)
        #Get the first two vectors from vect1 and store as u and v respectively
        u = vect1[0]
        v = vect1[-1]
        #Make cosine distance (cd) accessible from the sim2 function
        global cd
        #Calculate the cosine distance between the two vectors
        cd = numpy.dot(u, v) / (math.sqrt(numpy.dot(u, u)) * math.sqrt(numpy.dot(v, v)))

#Create the output for the user to see in a new window
def sim2(filename1, filename2):
    #Call the SIM class with stopwords as an argument and call the analyse function for each document and call the matrix function on both in one go
    dnsim = SIM(stopwords)
    for d in documents:
        dnsim.analyse(d)
    dnsim.matrix()
    #Get the base name of each file
    filename1 = os.path.basename(filename1)
    filename2 = os.path.basename(filename2)
    filename1 = filename1.capitalize()
    filename2 = filename2.capitalize()
    #Define boundaries of similarities
    if cd <= -0.5 and cd > -1:
        simtxt = "is very dissimilar to"
    elif cd > -0.5 and cd < 0:
        simtxt = "is dissimilar to"
    elif cd >0 and cd <0.5:
        simtxt = "is similar to"
    elif cd >=0.5 and cd <1:
        simtxt = "is very similar to"
    elif cd == -1:
        simtxt = "is independent of"
    elif cd == 0:
        simtxt = "is exactly opposite to"
    elif cd == 1:
        simtxt = "is a duplicate of"
    #Round the cosine distance to 2 decimal places
    value = "{0:.2f}".format(cd)
    value = "Cosine Value is "+str(value)
    f1 = str(filename1)
    f2 = str(filename2)

    #Create a new window for the output
    bk = "white"
    dn_sim = Toplevel(root, bg=bk)
    #Set the title name next to the logo
    dn_sim.title("Data Ninja - Document Similarity")
    #Choose random points within the ranges specified of where the new window will appear
    xs = range(0,295) 
    ys = range(40,280)
    xpos = choice(xs)
    ypos = choice(ys)
    coord = "800x200+"+str(xpos)+"+"+str(ypos)
    dn_sim.geometry(coord)
    #Create labels in the window
    colour = "#42C0FB"
    label = Label(dn_sim, text=simtxt, fg=colour, bg=bk, font=("Helvetica", 30))
    label.pack(side="left", fill="both", expand=1)
    label.place(x=30, y=80)
    label = Label(dn_sim, text=value, fg=colour, bg=bk, font=("Helvetica", 15)) 
    label.pack(side="left", fill="both", expand=1)
    label.place(x=0, y=0)
    label = Label(dn_sim, text=f1, fg=colour, bg=bk, font=("Helvetica", 25)) 
    label.pack(side="left", fill="both", expand=1)
    label.place(x=30, y=30)
    label = Label(dn_sim, text=f2, fg=colour, bg=bk, font=("Helvetica", 25)) 
    label.pack(side="left", fill="both", expand=1)
    label.place(x=30, y=140)


    
#-------------------------------------------------------------------------------------------------------End of doc similarity process for the doc similarity button
#-------------------------------------------------------------------------------------------------------Start of report generation process for the report generation feature and probe


#Function to allow the user to select an audit report and then call the gen function
def report():
    #Dialog box to allow the user to select an audit report
    reportX = tkFileDialog.askopenfilename(parent=root,title='Select an audit report')
    #Check that the audit report is a pdf or docx and that a report was selected
    if len(reportX) > 0 and reportX.endswith('pdf') or reportX.endswith('docx'):
        #Error Handling
        try:
            #Call the gen function to generate the insight log and draft management letter
            gen(reportX)
        #Error Handling
        except Exception:
            #Error box informing the user that processing unsuccessful
            tkMessageBox.showwarning("DN Report Generation", "Unable to open file\n\nEnsure the audit report has a .pdf or .docx extension and contains text")
    else:
        #Error box informing the user that processing unsuccessful
        tkMessageBox.showerror("DN Report Generation", "Unable to open file\n\nPlease select an audit report with a .pdf or .docx extension containing text")
        
#Probe function to check if a specific audit report is in a directory then call the gen function, d_w_report is directory with the report
def probe(d_w_report):
    #Get all files from the directory including files in sub directories
    for directory, subdirectories, files in os.walk(d_w_report):
        for i in files:
            #If a file name contains "XYZ Audit Report.pdf" probe the user
            if "XYZ Audit Report.pdf" in i:
                #Dialog box informing the user that an audit report was found in the directory and asking if the user wants to generate an insight log and draft management letter
                probe1 = tkMessageBox.askyesno("Report Generation?", "I noticed there is an audit report - "+str(i)+" in that directory.  \n\nWould you like to generate an insight log and draft management letter from this report?")
                #If the user clicks Yes, then the file path of the audit report found is fed into the gen function
                if probe1 == True:
                    r_path = os.path.join(directory, i)
                    #Error Handling
                    try:
                        gen(r_path)
                    #Error Handling
                    except Exception:
                        tkMessageBox.showwarning("DN Report Generation", "Unable to open file")
                else:
                    continue
            else:
                continue
            #Same as above but for a docx audit report
            if "XYZ Audit Report.docx" in i:
                probe1 = tkMessageBox.askyesno("Report Generation?", "I noticed there is an audit report - "+str(i)+" in that directory.  \n\nWould you like to generate an insight log and draft management letter from this report?")
                if probe1 == True:
                    r_path = os.path.join(directory, i)
                    #Error Handling
                    try:
                        gen(r_path)
                    #Error Handling
                    except Exception:
                        tkMessageBox.showwarning("DN Report Generation", "Unable to open file")
                else:
                    continue
            else:
                continue

#Make this dictionary accessible from the gen report
global rec
#Matching key phrases found in insights to recommendations to these insights
rec = {
         "penetration testing has not been performed" : "Penetration testing should be conducted every X months.",
         "changes are not approved" :
         "All application, network and os changes should be formally approved by the relevant administrator before being deployed onto the production environment.  Formal system testing should proceed this approval.",
         "changes are not archived" : "Change requests should be stored either electronically or manually and archived for a period of X years for referential purposes.",
         "anti-virus software installed expired" : "A reliable anti-virus software should be installed on the system and should be reviewed every X months.",
         "acceptance testing is not conducted" :
         "It is advisable that user acceptance testing is conducted for application changes because this will confirm whether user's are satisfied with the application or not.",
         "no IT security policy" : "An up to date IT policy should be in place which outlines the use of computer software and hardware by employees.",
         "no disaster recovery policy" : "A detailed up to date Disaster Recover Policy should be in place which outlines actions to be taken by employees and the business in the event of a disaster.",
         "password parameters are inadequate" : "Password parameters should be - minimum password length: X chars, password history: X, minimum password change: X days."
         }

         
#Generate a .doc draft management letter and .xls insight log with insights pulled in from the audit report
def gen(a_report):
    #Dialog box for the user to enter the audit company's name
    compname = tkSimpleDialog.askstring("DN Report Generation - Company Name?", "Enter the auditor's company name")
    compname = str(compname)
    #Capitalise all first letters of the company's name if the company's name is not all upper cases
    if compname.isupper() == False:
        compname = compname.title()
    #If the user typed something into the dialog box, do the next step
    if compname != "None":
        #Dialog box for the user to enter the client company's name
        clientname = tkSimpleDialog.askstring("DN Report Generation - Client Name?", "Enter the client company's name for audit")
        clientname = str(clientname)
        #Capitalise all first letters of the company's name if the company's name is not all upper cases
        if clientname.isupper() == False:
            clientname = clientname.title()
        #If the user typed something into the dialog box, do the next step
        if clientname != "None":
            #Dialog box for the user to enter the audit year
            auyear = tkSimpleDialog.askstring("DN Report Generation - Audit Year?", "Enter the year of audit")
            auyear = str(auyear)
            #If the user typed something into the dialog box, do the next step
            if auyear != "None":
                a_report = str(a_report)
                #If file is a pdf, call the readPdf function to read it
                if a_report.endswith('pdf'):
                    words = readPdf(a_report)
                #If file is a docx, call the readDocx function to read it
                elif report1.endswith('docx'):
                    words = readDocx(a_report)
                #Split the audit report into lines
                lines = re.split(r'\s*[!?.:]\s*', words)
                
                findings = ["InsightDN", "ContactDN", "MitigationDN", "RaiseDN", "AreaDN"]

                #Create an xls workbook with the following text and formatting
                book = xlwt.Workbook(encoding="utf-8")
                sheet1 = book.add_sheet("Findings")
                style = easyxf('font: name Times New Roman size 05 bold on, color dark_blue, height 400;')
                style1 = easyxf('font: color white, height 200; pattern: back_colour dark_blue, pattern thick_forward_diag, fore-colour dark_blue;')
                sheet1.write_merge(0,1,0,10, clientname+" Insights", style)
                sheet1.write(2, 0, "No", style1) 
                sheet1.write(2, 1, "Section", style1) 
                sheet1.write(2, 2, "Finding", style1) 
                sheet1.write(2, 3, "Contact", style1) 
                sheet1.write(2, 4, "Mitigation", style1) 
                sheet1.write(2, 5, "Raise in Draft Letter?", style1)

                #Write to a doc file and name it with the content of the variable draft
                draft = compname+" Draft Management Letter - "+clientname+" Audit "+auyear
                draft1 = dpath+"/"+draft+".doc"
                f = open(draft1, "w")
                time1 = time.strftime("%Y-%m-%d")
                f.write(draft+"\n\nDate: "+time1+"\n\nAuditors: "+compname)

                #Get lines that contain information required for the insight log and draft management letter and add this line to the ins list
                ins = []
                for i in lines:
                    if ("InsightDN" in i) or ("ContactDN" in i) or ("MitigationDN" in i) or ("RaiseDN" in i) or ("AreaDN" in i):
                        ins.append(i)
                #Store the length of the ins list in l
                l = len(ins)
                """Get the first 5 items in ins which will be 1 of each of the "..DN" ids above, add these to the doc then get the next 5"""
                n1 = 0
                n2 = 5
                #end is the number of items in the list ins i.e. the number of ids, divided by 5 which gives the number of sets of ids
                end = (l/5)
                #Create a list with the number of items equalling end (this is the number of times to iterate ins) equal to the number of sets of ids
                num1 = range(0, end)
                for i in num1:
                    #Get the current 5 items from the ins list
                    item1 = ins[n1:n2]
                    #Get the "RaiseDN.." item
                    raise1 = str(item1[3])
                    #Remove the id part of the text
                    raise1 = raise1.replace("RaiseDN ", "")
                    #If the remaining text is a "Yes" then write the insight details to the doc file
                    if raise1 == "Yes":
                        area = str(item1[4])
                        area = area.replace("AreaDN ", "")
                        f.write("\n\n\n\nArea: "+area)
                        #Separator
                        f.write("\n_________________________________________________________________________")
                        insight = str(item1[0])
                        insight = insight.replace("InsightDN ", "")
                        f.write("\n\nInsight: "+insight+".")
                        else1 = "Ensure formal procedures are in place, that prevent unnecessary risks to the system."
                        #Get recommendation for the insight
                        for key in rec:
                            if key in insight:
                                else1 = rec[key]
                        f.write("\n\nRecommendation: "+else1+"")
                        f.write("\n_________________________________________________________________________\n")
                    #Add 5 to get the next 5 items from the ins list on the next iteration, do this until it reaches the end number when the last set is done
                    n1 += 5
                    n2 += 5
                f.write("\n\n\n\n\n\n\nSigned:                                                Dated: "+time.strftime("%d/%m/%Y"))
                #Close doc file
                f.close()

                #Insight number
                txt = 1
                #Positions of insight information
                x1 = 3
                x2 = 3
                x3 = 3
                x4 = 3
                x5 = 3

                #Write Insight info to sheet 1
                for i in lines:
                    if "InsightDN" in i:
                        #Error Handling
                        try:
                            i = i.replace("InsightDN ", "")
                            sheet1.write(x1, 2, i)
                            sheet1.write(x1, 0, txt)
                            x1 += 1
                            txt += 1
                        #Error Handling
                        except Exception:
                            continue
                    else:
                        continue

                #Write Contact info (contact spoken to) to sheet 1
                for i in lines:    
                    if "ContactDN" in i:
                        #Error Handling
                        try:
                            i = i.replace("ContactDN ", "")
                            sheet1.write(x2, 3, i)
                            x2 += 1
                        #Error Handling
                        except Exception:
                            continue
                    else:
                        continue

                #Write Mitigation info to sheet 1
                for i in lines:   
                    if "MitigationDN" in i:
                        #Error Handling
                        try:
                            i = i.replace("MitigationDN ", "")
                            sheet1.write(x3, 4, i)
                            x3 += 1
                        #Error Handling
                        except Exception:
                            continue
                    else:
                        continue

                #Write Raise info to sheet 1
                for i in lines:   
                    if "RaiseDN" in i:
                        #Error Handling
                        try:
                            i = i.replace("RaiseDN ", "")
                            sheet1.write(x4, 5, i)
                            x4 += 1
                        #Error Handling
                        except Exception:
                            continue
                    else:
                            continue

                #Write Area info i.e. section to sheet 1
                for i in lines:
                    if "AreaDN" in i:
                        #Error Handling
                        try:
                            i = i.replace("AreaDN ", "")
                            sheet1.write(x5, 1, i)
                            x5 += 1
                        #Error Handling
                        except Exception:
                            continue
                    else:
                        continue

                #Set column width
                sheet1.col(1).width = 256*20
                sheet1.col(2).width = 256*160
                sheet1.col(3).width = 256*34
                sheet1.col(4).width = 256*92
                sheet1.col(5).width = 256*20
                #Save .xls file with the name specified below
                fname = compname+ " Insight Log - "+clientname+" Audit "+auyear+".xls"
                fname1 = dpath+"/"+fname
                book.save(fname1)
                #Message box to inform the user that the report generation was successful
                tkMessageBox.showinfo("DN Report Generation Complete", "Done!\n\nCheck your desktop for these 2 documents:\n\n"+fname+"\n\n"+draft+".doc")


        
#-------------------------------------------------------------------------------------------------------End of report generation process for the report generation feature


#
def main(b1,b2,b3):
    #Set the colour for the background and buttons in the DNInter class
    colour1(b2)
    colour2(b3)
    #Make the variable root accessible from anywhere within the script
    global root
    root = Tk()
    #Define root as the frame for the DNInter class
    inter = DNInter(root)
    #Automatically maximise window
    root.state("zoomed")

    #Create a toolbar menu in the root window
    menu = Menu(root)
    root.config(menu=menu)
    global reportmenu
    filemenu = Menu(menu, bg=b1)
    menu.add_cascade(label="File", menu=filemenu)
    filemenu.add_command(label="Report Generation", command=report)
    filemenu.add_separator()
    filemenu.add_command(label="Customise Number of Clusters", command=clusterno)
    filemenu.add_command(label="Default Number of Clusters", command=defclustno)
    filemenu.add_separator()
    filemenu.add_command(label="Quit", command=askquit)
    global helpmenu
    helpmenu = Menu(menu, bg=b1)
    menu.add_cascade(label="Help", menu=helpmenu)
    helpmenu.add_command(label="About...", command=about)
    #Get the user's "My Documents" path 
    l_path = os.path.join(os.path.expanduser("~"), "Documents")
    #Get the exact path of where the logo is saved - the user has to save the Data Ninja folder in My documents
    l_path = l_path+"\Data Ninja\dist\dnlogo.ico"
    #Change the logo
    root.iconbitmap(default=l_path)
    #Enter Tkinter's event loop 
    root.mainloop()


#This is the section that runs when the script runs
if __name__ == '__main__':
    #Set toolbar menu, background and button colours and call the main function
    b1 = "gray96"
    colour1("#42C0FB") 
    colour2("white")
    main(b1,b2,b3)  





