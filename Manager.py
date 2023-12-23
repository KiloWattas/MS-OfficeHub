# (C) 2023 DEIVIDAS STANYS
# KMS ACTIVATION SCRIPT MANAGER
# PYTHON 3.11.1 - 2023-12-23

from os import system, chdir

def Music():
    #system("mkdir c:\\KMS\\misc")
    #system("copy misc\\audio.wav c:\\KMS\\misc")
    system("cls")
    file = "audio.vbs"
    try:
        with open (file, 'w') as audio:
            audio.write("Set Sound = CreateObject(\"WMPlayer.OCX.7\")\nSound.URL = \"c:\\KMS\\misc\\audio.wav\"\nSound.Controls.play\ndo while Sound.currentmedia.duration = 0\nwscript.sleep 100\nLoop\nwscript.sleep (int(Sound.currentmedia.duration)+1)*1000")
            system("attrib +h +r +s audio.vbs")
    except Exception:
        pass
    system("start audio.vbs")


def Choices():
    while True:
        print(" ")
        choice = input(" [SYS] > What do you want to do? [EXIT - 0 | DOWNLOAD - 1 | ACTIVATE - 2] >>> ")

        if choice == "0":
            system("taskkill /f /im wscript.exe")
            exit()
        else: 
            version = input(" [SYS] > What is your preferred version [OFFICE 2016 - 1 | OFFICE 2019 - 2 | OFFICE 2021 - 3] >>> ")
            print(" ")

        if choice == "1":

            if version == "1":
                print(" [SEL] > Version selected: OFFICE 2016")
                file = "link_office_2016_ProPlus.txt"
                arch = input(" [SYS] > What is your preferred architecture [32-BIT - 1 | 64-BIT - 2] >>> ")
                print(" ")
                if arch == "1": 
                    with open (file, 'r') as doc:
                        link1, link2 = map(str, doc.read().splitlines())
                        del link1
                        print(" [SYS] > Downloading...")
                        system (f"start {link2}")
                elif arch == "2":
                    with open (file, 'r') as doc:
                        link1, link2 = map(str, doc.read().splitlines())
                        del link2
                        print(" [SYS] > Downloading...")
                        system (f"start {link1}")

            elif version == "2":
                print(" [SEL] > Version selected: OFFICE 2019")
                print(" ")
                file = "link_office_2019_ProPlus.txt"
                with open (file, 'r') as doc:
                    link1 = doc.read()
                    print(" [SYS] > Downloading...")
                    system (f"start {link1}")

            elif version == "3":
                print(" [SEL] > Version selected: OFFICE 2021")
                print(" ")
                file = "link_office_2021_ProPlus.txt"
                with open (file, 'r') as doc:
                    link1 = doc.read()
                    print(" [SYS] > Downloading...")
                    system (f"start {link1}")
            else:
                print(" [ERR] > Invalid Office version selected")

        elif choice == "2":
            #system("cmd /c mkdir c:\\KMS")
            if version == "1":
                #system("cmd /c copy activator_office_2016.py c:\\KMS")
                print(" [SYS] > Activating...")
                system("c:\\KMS\\activator_office_2016")
            elif version == "2":
                #system("cmd /c copy activator_office_2019.py c:\\KMS")
                print(" [SYS] > Activating...")
                system("c:\\KMS\\activator_office_2019")
            elif version == "3":
                #system("cmd /c copy activator_office_2021.py c:\\KMS")
                print(" [SYS] > Activating...")
                system("c:\\KMS\\activator_office_2021")
            else:
                print(" [ERR] > Invalid Office version selected")
        
        else:
            print(" [ERR] > Invalid choice")


system ("mode CON: cols=129 lines=35")
Music()

print(" ")
print(" o     o .oPYo.     .oPYo.  d'b  d'b  o                o    o        8          o     o                                         \n 8b   d8 8          8    8  8    8                     8    8        8          8b   d8                                         \n 8`b d'8 `Yooo.     8    8 o8P  o8P  o8 .oPYo. .oPYo. o8oooo8 o    o 8oPYo.     8`b d'8 .oPYo. odYo. .oPYo. .oPYo. .oPYo. oPYo. \n 8 `o' 8     `8     8    8  8    8    8 8    ' 8oooo8  8    8 8    8 8    8     8 `o' 8 .oooo8 8' `8 .oooo8 8    8 8oooo8 8  `' \n 8     8      8     8    8  8    8    8 8    . 8.      8    8 8    8 8    8     8     8 8    8 8   8 8    8 8    8 8.     8     \n 8     8 `YooP'     `YooP'  8    8    8 `YooP' `Yooo'  8    8 `YooP' `YooP'     8     8 `YooP8 8   8 `YooP8 `YooP8 `Yooo' 8     \n ..::::..:.....::::::.....::..:::..:::..:.....::.....::..:::..:.....::.....:::::..::::..:.....:..::..:.....::....8 :.....:..::::\n :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::ooP'.:::::::::::::\n :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::...:::::::::::::::")
print(" ")

Choices()