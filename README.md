# xconf_to_excel
Program written using Python 3.5 that give a text file with the output from a "xconfiguration" command on the VCS CLI. Note that one of the resulting files from initiating diagnostic logging on the web interface can also be used. The program required of the following third party modules:
* appdirs==1.4.0
* et-xmlfile==1.0.1
* jdcal==1.3
* openpyxl==2.4.2
* packaging==16.8
* pyparsing==2.1.10
* six==1.10.0
 
The program parser the text file and read line from it. The objective is organize the information in Python dictionaries that will be used by a separate module to build the Excel sheet with the information. The script is comprised by two different modules:
xconf_to_dict.py: responsible of parsing the text and build the dictionaries
xconf2excel: module that calls the function from xconf_to_dict and build a Excel sheet with that information.
