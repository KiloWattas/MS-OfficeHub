# (C) 2023 DEIVIDAS STANYS
# AUTOMATIC KMS ACTIVATION SCRIPT
# MS Office 2019 Professional Plus
# PYTHON 3.11.1 - 2023-12-20


from os import system, chdir
import os.path as checkifpath
import ctypes, sys

def whichVersion():
    x86 = checkifpath.exists("c:\Program Files (x86)\Microsoft Office\Office16")
    x64 = checkifpath.exists("c:\Program Files\Microsoft Office\Office16")
    if x86:
        return 1
    elif x64:
        return 2
    else:
        print(" [ERROR] > Could not find Office installation")

def is_admin(): # CHECKS IF CURRENT USER IS ADMIN
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if is_admin(): # PERFORM ACTIVATION IF CURRENT USER IS ADMIN
    option = whichVersion()

    if option == 1:
        chdir("c:\Program Files (x86)\Microsoft Office\Office16")
        system("for /f %x in ('dir /b ..\\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:\"..\\root\Licenses16\%x\"")
        system("cscript ospp.vbs /setprt:1688")
        system("cscript ospp.vbs /unpkey:6MWKP >nul")
        system("cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP")
        system("cscript ospp.vbs /sethst:e8.us.to")
        system("cscript ospp.vbs /act")
        print(" ")
        print(" ")
        print(" >>>   D O N E  !   <<<")
        print(" ")
        input(" PRESS [ENTER] TO EXIT...")
        exit()

    elif option == 2:
        chdir("c:\Program Files\Microsoft Office\Office16")
        system("for /f %x in ('dir /b ..\\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:\"..\\root\Licenses16\%x\"")
        system("cscript ospp.vbs /setprt:1688")
        system("cscript ospp.vbs /unpkey:6MWKP >nul")
        system("cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP")
        system("cscript ospp.vbs /sethst:e8.us.to")
        system("cscript ospp.vbs /act")
        print(" ")
        print(" ")
        print(" >>>   D O N E  !   <<<")
        print(" ")
        input(" PRESS [ENTER] TO EXIT...")
        exit()

    else:
        print(" [ERROR] > Could not activate Your office installation. Maybe the version You have is not supported?")
        input(" PRESS [ENTER] TO EXIT...")
        exit()
else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1) # EVOKE UAC PROMPT AND EXECUTE AS ADMIN
