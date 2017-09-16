#requires MS Windows and runs on Windows
#sample Code to list intalled printer and set a default printer
#use subprocess to background print a word file to the default printer. 

import types, pprint
import win32api, win32con, win32print
import subprocess

# Get a list of attached printers
def GetPrinters():
  '''Gets local printers into printers list'''
  printers = []
  try:
    plist = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, '')
    for p in plist:
      printers.append(p[2])
    return printers
  except:
    print ('Error Enumerating Printers')

# Change the default printer 
def ChangeDefaultPrinter(printer):
  #if type(printer) == types.StringType:
  try:
    win32print.SetDefaultPrinter(printer)
    print ('  Default Printer set to: %s' % printer)
  except:
    print ('  Specified printer %s not available' % printer)
    raise
  #else:
  #  print ( '  Default printer not set, no printers available')

 if __name__  ==  '__main__':
# get a list of printers on the system
	printers = GetPrinters()
# print out list of printers
	pprint.pprint(printers)
# change printer (example on my work network)
#	ChangeDefaultPrinter(r'HP Photosmart C309a series')
	ChangeDefaultPrinter(r'\\domain\XEROX-1')
# Print a file
	subprocess.Popen([r"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE", r".\docs\\new-file-name.docx"), "/mFilePrintDefault", "/mFileExit"]).communicate()
