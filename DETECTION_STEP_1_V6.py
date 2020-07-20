## V1
## V2 
## V3: read txt. better filter weather degree
## V4: removed step for loading final las and all mentions of las or lmp in scripts
## V5: removed steps to save marker files and moved 'clear mrk files' to the end of each script.
## V6: combine V4 and V5. Fixed some bugs

import pywinauto, xlrd, re, time, os, warnings
from time import sleep
from pywinauto.timings import wait_until
from pywinauto.keyboard import send_keys
from pywinauto.application import Application
warnings.filterwarnings("ignore")
##app = Application(backend="win32").start("C:\\Program Files (x86)\\PLS\\pls_cadd\\pls_cadd64.exe")
app = Application(backend="win32").start("C:\\Program Files\\PLS\\pls_cadd\\pls_cadd64.exe")
os.system('cls' if os.name == 'nt' else 'clear')
try:
    app.about_PLS_CADD.OK.click
except:
    pass
app.Dialog.OK.click()
## Gather input values

## Reading info from Veg_Detection_Command.txt
fname='C:\\Scripts\\Veg_Detection_Command.txt'
f = open(fname)
text = []
for line in f:
    text.append(line[:-1])
processDir=text[1]
#processGroup=sheet.cell_value(6,1)
#LineID=str(int(sheet.cell_value(7,1)))
#Kv=str(int(sheet.cell_value(8,1)))
#print("File Path:"+processDir,"Processing Group:"+processGroup,"LineID is "+LineID,"Line KV is "+Kv+"Kv")
#NAME=LineID+"_"+Kv+"kv"
#FINAL_LAS_PATH=text[7]
#FINAL_LAS_PATH=FINAL_LAS_PATH[:-1]
#print(FINAL_LAS_PATH)
BAK_PATH=text[3]
#BAK_PATH=BAK_PATH[:-1]
#print(BAK_PATH)
#FINAL_LAS=text[9]
#FINAL_LAS=FINAL_LAS[:-1]
#print(FINAL_LAS)
BAK=text[5]
#BAK=BAK[:-1]

##    Open bak file
app.PLS_CADD_Project_wizard.Restore_Project_from_PLS_CADD_Backup_file.click()
app.PLS_CADD_Project_wizard.OK.click()
#app.PLS_CADD.menu_select("File->Restore_Backup")
#app[u'Restore_Backup']['File_name:Edit'].set_text(processDir+processGroup+"\\"+NAME+"_qsi2018_detections.bak") #2019
app[u'Restore_Backup']['File_name:Edit'].set_text(BAK_PATH+"\\"+BAK+".bak") 
sleep(2.0)
app.window(best_match='Restore_Backup', top_level_only=True).child_window(best_match='Open').click()
sleep(2.0)
#app[u'Restore_Backup']['Open'].click()
app.Directory_Mapping_For_Restore.Quick_Restore.click()
sleep(1.0)
app[u'Select_Directory_To_Restore_Files_In']['File_name:Edit'].set_text(BAK_PATH)
sleep(1.0)
#app[u'Select_Directory_To_Restore_Files_In']['Select_Folder'].click()
app.window(best_match='Select_Directory_To_Restore_Files_In', top_level_only=True).child_window(best_match='Select_Folder').double_click()
#send_keys('{ENTER}')
app.PLS_CADD.OK.click()
try:
    app.PLS_CADD.OK.click
except:
    pass
try:
    #app[u'Restore_Backup_Of_'+NAME+'_qsi2018_veg_detections_bak']['Always'].click()
    app.top_window().Always.click()
except:
    pass
#app.restore_backup.wait_not('visible', timeout=100)
app.restore_backup.wait_not('exists', timeout=1000)
#app[u'Restore_Backup_Of_'+BAK+".bak"].wait('ready',timeout=30)
sleep(2.0)
app[u'Restore_Backup_Of_'+BAK+".bak"]['Yes'].click()
app.Calculating_terrain_stations_and_offsets.wait_not('exists', timeout=6000)
app.Calculating_terrain_stations_and_offsets.wait_not('exists', timeout=6000)
app.checking_xyz_data.wait_not('exists', timeout=6000)
sleep(2.0)

## 3D view and TIN
app.top_window().menu_select("Window-> 4 3D view")
app.top_window().menu_select("Terrain-> TIN-> Display options")
app.Ground_TIN_Display_Options.render_triangles.check()
app.Ground_TIN_Display_Options.OK.click()
print("Ready to run next script")





