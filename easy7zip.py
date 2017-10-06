# -*- coding: cp1252 -*-
import subprocess
import os
import sys
import time
import re

if sys.version_info[0] != 2:
	raise Exception("This module required  Python version 2.7")

__author__ = 'Sreejith KOVILAKATHUVEETTIL CHANDRAN'
__version__ = '0.0.1'
__copyright__ = " Copyright 2017,SREEJITH KOVILAKATHUVEETTIL CHANDRAN"
__email__ = "sreeju_kc@hotmail.com"
__license__ = "Apache License 2.0"
__last_modification__ = '2017.10.03'

"""The intention of this module is to provide an easy and handy way to create, extract,update, delete,hash value and list the 7zip files."""
"""This is module only work on Python 2.7 on Windows paltform and the only prerequest is to install 7zip program on on one of the following locations "C:\7Zip","C:\Program Files","C:\Program Files (x86)"""

class easy7zip(object):
    def __init__(self,path = ""):
        ''' This is module is only for Windows and it works on Python 2.7.
            This required 7zip program to be installed on your machine 
            Please make sure that 7zip program is installed on one of the following locations "C:\7Zip","C:\Program Files","C:\Program Files (x86)'''
        folders = ["C:\\7Zip","C:\\Program Files","C:\\Program Files (x86)"]
        lookfor = "7za.exe"
        for pa in folders:
            for root, dirs, files in os.walk(pa):
                if lookfor in files:
                    path = root
                    self.path = path
                    break
            if  path:
                break
            else:
                continue
        if not path:
            raise Exception ("Please download and install 7zip from 'http://www.7-zip.org/' and try again, please make sure to install it either C drive or C:\Program Files ")
    def AddToArch(self,zipfilename,filetoadd,passwd=None):
        ''' This function will create a 7z archive file either with or without password
            This function takes 3 arguments,1st argument take the file (with full path) name of intended archive file you are about to  create for example "C:\dir\file", this will create a file.7z file in C:\dir directory.
            2nd argument is full path of your source file, you can either give one file or give * for all files in the folder.
            3rd argument is optional for passwords, if 3rd argument is empty  it will create an archive without password, if you need to create an archive  with password then provide the password in the 3rd argument.
            If the overall operation is successful this function will return a Boolean True value.'''
            
        path =  self.path
        path.encode('string_escape')
        if not zipfilename.endswith(".7z"):
            zipfilename = zipfilename+".7z"
        filetoadd = '"{}"'.format(filetoadd)
        filetoadd = filetoadd.encode('string-escape')
        zipfilename = zipfilename.encode('string-escape')
        tempar = zipfilename+".tmp"
        zipfilename = '"{}"'.format(zipfilename)
        if os.path.exists(tempar):
            os.remove(tempar)
        out = ""
        if passwd==None:
            try: 
                cmd = path+"\\7za.exe a -t7z "+"-mx9 "+zipfilename+' '+filetoadd
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try: 
                cmd = path+"\\7za.exe a -t7z -p"+passwd+" -mx9 "+ zipfilename+' '+filetoadd
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        if """Everything is Ok""" in out:
            return True
        elif "The process cannot access the file because it is being used by another process" in out:
            raise Exception ("It seems Zip file is already exists and opend by another process, please close the zip file and try again")
        else:
            raise Exception ("Some error while creating Zip file, make sure file and directory are valid" +'\n'+out)
    def ExtractFromArch(self,zipfile,foldertoextract,passwd=None):
        ''' This function will extract a 7z archive file to a destination folder .
            This function takes 3 arguments,1st argument take the file (with full path) name of the source archive file you are about to  extract for example "C:\dir\file.7z".
            2nd argument is full path of your destination folder for your extracted files, for example �C:\destination�.
            3rd argument is optional for passwords, if the archive file is password protected then you can give your password there, otherwise leave it blank.
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not zipfile.endswith(".7z"):
            zipfile = zipfile+".7z"
        zipfile = '"{}"'.format(zipfile)
        zipfile = zipfile.encode('string-escape')
        foldertoextract = '"{}"'.format(foldertoextract)
        foldertoextract = foldertoextract.encode('string-escape')
        if passwd==None:
            try:
                cmd = path+"\\7za.exe x -aos "+zipfile+' '+"-o"+foldertoextract
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try:
                cmd = path+"\\7za.exe x -aos "+zipfile+' '+"-o"+foldertoextract+" -p"+passwd
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        if """Everything is Ok""" in out:
            return True
        elif "Enter password (will not be echoed)" in out:
             raise Exception ("Some files are password protected, please provide password to extract")
        elif "Scanning the drive for archives:" in out:
             raise Exception ("Either 7z file or destination folder is not valid, please try again "+'\n'+'\n'+'\n'+out) 
        else:
             raise Exception ("Some error while extracting Zip file " +'\n'+out)
    def UpdateToArch(self,archfile,filetoupdate,passwd=None):
        ''' This function will update an existing 7z archive file
            This function takes 3 arguments,1st argument take the file (with full path) name of the source archive file you want to get updated, for example "C:\dir\file.7z".
            2nd argument is full path of your destination file you want to add it to the archive, for example �C:\destination\file.txt�.
            3rd argument is optional , if you want the archive file  password protected then you can give your password there, otherwise leave it blank.
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not archfile.endswith(".7z"):
            archfile = archfile+".7z"
        archfile = archfile.encode('string-escape')
        tempar = archfile+".tmp"
        archfile = '"{}"'.format(archfile)
        if os.path.exists(tempar):
            os.remove(tempar)
        filetoupdate = '"{}"'.format(filetoupdate)
        filetoupdate = filetoupdate.encode('string-escape')
        if passwd==None:
            try:
                cmd = path+"\\7za.exe u -aos "+archfile+' '+filetoupdate
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try:
                cmd = path+"\\7za.exe u -aos "+archfile+' '+filetoupdate+" -mx9 -p"+passwd
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        if """Everything is Ok""" in out:
            return True
        elif "cannot delete the file" in out:
            raise Exception ("Please make sure to close the 7z file before updating")
        else:
            raise Exception ("Some error while updating Zip file " +'\n'+out)

    def DeleteFromArch(self,archfile,filetodelete):
        ''' This function will delete a file from 7z archive file
            This function takes 2 arguments
            1st argument take the file (with full path) name of the source archive file from where you want to delete a file, for example "C:\dir\file.7z".
            2nd argument is the name of the file  you want to get deleted from the archive, for example �file.txt�. 
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not archfile.endswith(".7z"):
            archfile = archfile+".7z"
        archfile = archfile.encode('string-escape')
        tempar = archfile+".tmp"
        archfile = '"{}"'.format(archfile)
        if os.path.exists(tempar):
            os.remove(tempar)
        filetodelete = '"{}"'.format(filetodelete)
        filetodelete = filetodelete.encode('string-escape')
        try:
            cmd = path+"\\7za.exe d "+archfile+' '+filetodelete
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
            out, err = p.communicate()
        except Exception as e:
            print (e.message, e.args)
        
        if """Everything is Ok""" in out:
            return True
        elif "cannot delete the file" in out:
            raise Exception ("Please make sure to close the 7z file before updating")
        else:
            raise Exception ("Some error while deleting the file " +'\n'+out)
            
    def HashValue(self,zfile):
        
        ''' This function will Calculate hash values for files.
            This function take only one argument, full path of 7z file and return a hash value'''
        path =  self.path
        path.encode('string_escape')
        if not zfile.endswith(".7z"):
            zfile = zfile+".7z"
        zfile = '"{}"'.format(zfile)
        zfile = zfile.encode('string-escape')
        try:
            cmd = path+"\\7za.exe h "+zfile
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
            out, err = p.communicate()
        except Exception as e:
            print (e.message, e.args)
        if """Everything is Ok""" in out:
            hashv = re.search("CRC32  for data:\s+(.*)", out).group(1)
            hashv = hashv.rstrip('\r')
            return hashv
        else:
            raise Exception ("Some error in calculating hash value of the file " +'\n'+out)
    def ListFiles(self,zfile):
        
        ''' This function will list all the files inside a 7z archive file .
            This function take only one argument, full path of 7z file and return a list with files'''
        path =  self.path
        path.encode('string_escape')
        if not zfile.endswith(".7z"):
            zfile = zfile+".7z"
        zfile = '"{}"'.format(zfile)
        zfile = zfile.encode('string-escape')
        try:
            cmd = path+"\\7za.exe l "+zfile
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
            out, err = p.communicate()
        except Exception as e:
            print (e.message, e.args)
        if """Name""" in out:
            allfiles = re.findall("\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2}\s.*\s+\d+\s+\d+\s(.*)", out)
            allfiles = allfiles[:-1]
            allfiles = [x.rstrip() for x in allfiles]
            return allfiles
        else:
            raise Exception ("Some error in listing file " +'\n'+out)
