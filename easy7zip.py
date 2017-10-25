# -*- coding: cp1252 -*-
import subprocess
import os
import sys
import time
import re


if sys.platform != 'win32':
        raise Exception("This is a Windows 7-Zip module for Python 2.7")
if sys.version_info[0] != 2:
	raise Exception("This module requires  Python version 2.7")
__author__ = 'Sreejith KOVILAKATHUVEETTIL CHANDRAN'
__version__ = '0.0.11'
__copyright__ = " Copyright 2017,SREEJITH KOVILAKATHUVEETTIL CHANDRAN"
__email__ = "sreeju_kc@hotmail.com"
__license__ = "Apache License 2.0"
__last_modification__ = '2017.10.08'

"""This is a Windows 7-Zip module for Python 2.7"""
"""The intention of this module is to provide an easy and handy way to create, extract, update, delete, hash value and list the 7-Zip files."""
"""This module will only work on Python 2.7 (Windows paltform) and the only prerequisite is to install 7-Zip program in one of the following locations "C:\7Zip","C:\Program Files","C:\Program Files (x86)"""

class easy7zip(object):
    def __init__(self):
        ''' This  module is only for Windows and it works on Python 2.7.
            This required 7-Zip program to be installed on your machine 
            Please make sure that 7-Zip program is installed on one of the following locations "C:\7Zip","C:\Program Files","C:\Program Files (x86)'''
        path = ""
        folders = ["C:\\7Zip","C:\\7-Zip","C:\\7-zip","C:\\7zip","C:\\Program Files","C:\\Program Files (x86)"]
        lookfor = "7za.exe"
        nloofr = "7z.exe"
        for pa in folders:
            for root, dirs, files in os.walk(pa):
                if lookfor in files:
                    path = root+"\\"+lookfor
                    self.path = path
                    break
                elif nloofr in files:
                        path = root+"\\"+nloofr
                        self.path = path
                        break
                        
            if  path:
                break
            else:
                continue
        if not path:
            raise Exception ("Please download and install 7-Zip from 'http://www.7-zip.org/' and try again, please make sure to install it either C drive or C:\Program Files ")
    def AddToArch(self,zipfilename,filetoadd,passwd=None):
        ''' This function will create a 7-Zip archive file either with or without a password
            This function takes 3 arguments, 1st argument take the file (with full path) name of intended archive file you are about to create for example "C:\dir\file", this will create a file.7z file in the C:\dir directory.
            2nd argument is full path of your source file, you can either give one file or give * for all files in the folder.
            3rd argument is optional for passwords, if 3rd argument is empty, it will create an archive without a password, if you need to create an archive  with password then provide the password in the 3rd argument.
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not zipfilename.endswith(('.7z','.zip','.gz')):
                raise Exception("Please provide the file extention .7z,.zip or .gz")
        filetoadd = '"{}"'.format(filetoadd)
        tempar = zipfilename+".tmp"
        zipfilename = '"{}"'.format(zipfilename)
        zipfilenames = pathes(zipfilename)
        filetoadds = pathes(filetoadd)   
        if os.path.exists(tempar):
            os.remove(tempar)
        out = ""
        if passwd==None:
            try: 
                cmd = path+" a -t7z "+"-mx9 "+zipfilenames+' '+filetoadds
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try: 
                cmd = path+" a -t7z -p"+passwd+" -mx9 "+ zipfilenames+' '+filetoadds
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        if """Everything is Ok""" in out:
            return True
        elif "The process cannot access the file because it is being used by another process" in out:
            raise Exception ("It seems 7-Zip file is already exists and opend by another process, please close the zip file and try again")
        else:
            raise Exception ("Some error while creating Zip file, make sure file and directory are valid" +'\n'+out)
    def ExtractFromArch(self,zipfile,foldertoextract,passwd=None):
        ''' This function will extract a 7-Zip archive file to a destination folder .
            This function takes 3 arguments,1st argument take the file (with full path) name of the source archive file you are about to  extract for example "C:\dir\file.7z".
            2nd argument is full path of your destination folder for your extracted files, for example “C:\destination”.
            3rd argument is optional for passwords, if the archive file is password protected then you can give your password there, otherwise leave it blank.
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not zipfile.endswith(('.7z','.zip','.gz')):
                raise Exception("Please provide the file extention .7z,.zip or .gz")
        zipfile = '"{}"'.format(zipfile)
        foldertoextract = '"{}"'.format(foldertoextract)
        zipfiles = pathes(zipfile)
        foldertoextracts = pathes(foldertoextract)  
        if passwd==None:
            try:
                cmd = path+" x -aos "+zipfiles+' '+"-o"+foldertoextracts
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try:
                cmd = path+" x -aos "+zipfiles+' '+"-o"+foldertoextracts+" -p"+passwd
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
             raise Exception ("Some error while extracting 7-Zip file " +'\n'+out)
    def UpdateToArch(self,archfile,filetoupdate,passwd=None):
            
        '''This function will update an existing 7z archive file
            This function takes 3 arguments,1st argument take the file (with full path) name of the source archive file you want to get updated, for example "C:\dir\file.7z".
            2nd argument is full path of your destination file you want to add it to the archive, for example “C:\destination\file.txt”.
            3rd argument is optional, if you want the archive file  password protected then you can give your password there, otherwise leave it blank.
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not archfile.endswith(('.7z','.zip','.gz')):
                raise Exception("Please provide the file extention .7z,.zip or .gz")
        tempar = archfile+".tmp"
        archfile = '"{}"'.format(archfile)
        if os.path.exists(tempar):
            os.remove(tempar)
        filetoupdate = '"{}"'.format(filetoupdate)
        archfiles = pathes(archfile)
        filetoupdates = pathes(filetoupdate) 
        if passwd==None:
            try:
                cmd = path+" u -aos "+archfiles+' '+filetoupdates
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                p = subprocess.Popen(cmd,stdout=subprocess.PIPE, stderr=subprocess.PIPE,startupinfo=si)
                out, err = p.communicate()
            except Exception as e:
                print (e.message, e.args)
        else:
            try:
                cmd = path+" u -aos "+archfiles+' '+filetoupdates+" -mx9 -p"+passwd
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
            1st argument takes the file (with full path) name of the source archive file from where you want to delete a file, for example "C:\dir\file.7z".
            2nd argument is the name of the file you want to get deleted from the archive, for example “file.txt”. 
            If the overall operation is successful this function will return a Boolean True value.'''
        path =  self.path
        path.encode('string_escape')
        if not archfile.endswith(('.7z','.zip','.gz')):
            raise Exception("Please provide the file extention .7z,.zip or .gz")
        tempar = archfile+".tmp"
        archfile = '"{}"'.format(archfile)
        if os.path.exists(tempar):
            os.remove(tempar)
        filetodelete = '"{}"'.format(filetodelete)
        archfiles = pathes(archfile)
        filetodeletes = pathes(filetodelete)
        try:
            cmd = path+" d "+archfiles+' '+filetodeletes
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
        
         ''' This function will Calculate hash values (CRC32) for files.
            This function takes only one argument, full path of 7z file and return a hash value'''
         
         path =  self.path
         path.encode('string_escape')
         if not zfile.endswith(('.7z','.zip','.gz')):
             raise Exception("Please provide the file extention .7z,.zip or .gz")
         zfile = '"{}"'.format(zfile)
         zfiles = pathes(zfile)
         try:
             cmd = path+" h "+zfiles
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
        if not zfile.endswith(('.7z','.zip','.gz')):
            raise Exception("Please provide the file extention .7z,.zip or .gz")
        zfile = '"{}"'.format(zfile)
        zfiles = pathes(zfile)
        try:
            cmd = path+" l "+zfiles
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

def pathes(takepath):
        backslash = { '\a': r'\a', '\b': r'\b', '\f': r'\f', 
                  '\n': r'\n', '\r': r'\r', '\t': r'\t', '\v': r'\v' }
        for key, value in backslash.items():
                takepath = takepath.replace(key, value)
        return takepath
