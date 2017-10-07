python-easy7zip
===========

python-easy7zip is a python module which helps to create, delete, update, extract 7-Zip archive from Python 2.7 (Windows platform only).
It is very easy and handy to use from a python program.
7-Zip supports AES encryption so this can be very useful for security folks.
The only prerequisite  for this module is a preinstalled 7-Zip program in either C:\ or C:\Program Files.
This module is available via pip install, please use pip2.7 to install the python-easy7zip module.(pip install easy7zip)

Typical usage looks like:

>>> from easy7zip import easy7zip

>>> test = easy7zip()                                                               # Create an object from Class easy7zip

>>> test.AddToArch("D:\ZipTest\test","D:\ZipTest\test1.csv","test")                 #This will create a test.7z file and add test1.csv with a security key "test".Return a boolean True if operation is successful.


    True                                                                            
>>> test.AddToArch("D:\ZipTest\test1","D:\ZipTest\test1.csv")                       #This will create a test1.7z file and add test1.csv without a security key..Return a boolean True if operation is successful.

    True 

>>> test.UpdateToArch("D:\ZipTest\test.7z","D:\ZipTest\test.csv","test")            #Update file test.csv to test.7z with a security key "test"..Return a boolean True if operation is successful.

    True

>>> test.UpdateToArch("D:\ZipTest\test","D:\ZipTest\test1.txt")                     #Update file test.csv to test.7z without a security key..Return a boolean True if operation is successful.

    True

>>> test.ExtractFromArch("D:\ZipTest\test","D:\ZipTest\extra","test")               #Extract file to "extra" folder, 'test' is the security key as some of the files are password protected, leave it blank if none.Return a boolean True if operation is successful.

    True

>>> test.HashValue("D:\ZipTest\test")                                               #This function will take a 7zip file and return a hash value of it.

    'A40C0D4B'

>>> test.ListFiles("D:\ZipTest\test")                                               #This will return a list with all the files inside the archive.

    [' test1.txt', ' test2.txt', ' test3.txt', ' test1.csv', ' test.csv']

>>> test.DeleteFromArch("D:\ZipTest\test","test1.txt")                               # This function delete a file from the 7zip archive.Return a boolean True if operation is successful.

  True

>>> test.ListFiles("D:\ZipTest\test")

    [' test2.txt', ' test3.txt', ' test1.csv', ' test.csv']
