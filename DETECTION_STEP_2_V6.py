#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pywinauto, xlrd, re, time, warnings
from time import sleep
from pywinauto.timings import wait_until
from pywinauto.keyboard import send_keys
from pywinauto.application import Application

## for weather
def weather_value(F):   
    C=0
    if (F==122):
        C=50
    elif (F==140):
        C=60
    elif (F==158):
        C=167
    elif (F==176):
        C=80
    elif (F==194):
        C=90
    elif (F==212):
        C=100
    elif (F==230):
        C=110
    elif (F==248):
        C=120
    elif (F==266):
        C=130
    elif (F==294):
        C=140
    elif (F==302):
        C=150 
    elif (F==320):
        C=160   
    elif (F==338):
        C=170   
    elif (F==392):
        C=200
    else:
        C=F
    return C

## Gather input values
## Read by txt
warnings.filterwarnings("ignore")
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
#print(FINAL_LAS_PATH)
#FINAL_LAS=FINAL_LAS[:-1]
#print(FINAL_LAS)
BAK=text[5]
#BAK=BAK[:-1]
#print("bak IS"+BAK)
findstring="for"
APPNAME="PLS-CADD - "+BAK
CIRCUIT_KV=str(BAK)
if findstring in APPNAME:
    APPNAME=APPNAME[:-15]
    CIRCUIT_KV=BAK[:-23]
    APPNAME=APPNAME+".xyz"
    print("Working on bak:"+APPNAME)
else: 
    APPNAME=APPNAME[:-11]
    CIRCUIT_KV=BAK[:-19]
    APPNAME=APPNAME+".xyz"
    print("Working on bak"+APPNAME)
#print("ckv IS "+CIRCUIT_KV)
RESOURCE_PATH=text[12]
#IMP=text[18]
FEA=text[14]
BLOWOUT_FEA=text[16]

## count index and line names
multiple=BAK.split('_')
kv_index=[i for i, j in enumerate(multiple) if 'kv' in j]

### by excel
###warnings.filterwarnings("ignore")
###fname='C:\\Scripts\\Veg_Detection_Command.xlsx'
###xl=xlrd.open_workbook(fname)
###sheet=xl.sheet_by_index(0)
###processDir=sheet.cell_value(5,1)
#processGroup=sheet.cell_value(6,1)
#LineID=str(int(sheet.cell_value(7,1)))
#Kv=str(int(sheet.cell_value(8,1)))
#print("File Path:"+processDir,"Processing Group:"+processGroup,"LineID is "+LineID,"Line KV is "+Kv+"Kv")
#NAME=LineID+"_"+Kv+"kv"
###FINAL_LAS_PATH=sheet.cell_value(9,1)
###BAK_PATH=sheet.cell_value(10,1)
###FINAL_LAS=sheet.cell_value(11,1)
###BAK=sheet.cell_value(12,1)
###RESOURCE_PATH=sheet.cell_value(14,1)
###IMP=sheet.cell_value(15,1)
###FEA=sheet.cell_value(16,1)
###BLOWOUT_FEA=sheet.cell_value(17,1)
#APPNAME=sheet.cell_value(13,1)
###findstring="for"
###APPNAME="PLS-CADD - "+BAK
###CIRCUIT_KV=str(BAK)
###if findstring in APPNAME:
###    APPNAME=APPNAME[:-15]
###    CIRCUIT_KV=BAK[:-23]
###    APPNAME=APPNAME+".xyz"
###else: 
###    APPNAME=APPNAME[:-11]
###    Circuit_KV=BAK[:-19]
###    APPNAME=APPNAME+".xyz"
#app = Application().connect(title_re="PLS-CADD - ALP_SUC_1_69kV_qsi2019.xyz")
app = Application().connect(title_re=APPNAME)


## close TIN
app.top_window().menu_select("Terrain-> TIN -> Display_Options")
app.Ground_TIN_Display_Options.render_triangles.uncheck()
app.Ground_TIN_Display_Options.OK.click()
sleep(1.0)
## Step 6 Creat MaxOP Detection
app.top_window().menu_select("Line-> Edit")
try:
    dlg = app.window(title_re="Warning")
    app.dlg.no.click()
except:
    pass
app.Line.Copy.click()
try:
    app.Copy_Line.OK.click()
except:
    pass
text_info = app.Line.Edit.window_text()
w_info=app.Line.Edit1.window_text()
app.Line.Info.click()
app[u'Line_Display_Options']['File_name:Edit'].set_text("temp")
sleep(1.0)
app.Line_Display_Options.OK.click()
try:
    app.Line_Display_Options.OK.click()
except:
    pass
app.Line.Select_and_hide_other_lines.click()
Listbox_text=app.Line.ListBoxWrapper.item_texts()
#print("test1"+str(Listbox_text))
#MaxOP_Library=["Max Op - 32", "Max Op - 122", "Max Op - 140", "Max Op - 158", "Max Op - 167", "Max Op - 176", "Max Op - 194", "Max Op - 212", "Max Op - 230", "Max Op - 248", "Max Op - 266", "Max Op - 294", "Max Op - 302", "Max Op - 320","Max Op - 338", "Max Op - 392"]
#Target_MaxOP=[ "Max Op - 266","Max Op - 392"]
find_sub_op="Max Op"
find_sub_as="As Surveyed"
find_sub_combined="combined"
# Line name detection

#if find_sub not in str(Listbox_text):
#    MaxOp_list=[i for i in Listbox_text if find_sub_as in i]
#else:
#    MaxOp_list=[i for i in Listbox_text if find_sub in i]

if find_sub_op not in str(Listbox_text) and find_sub_combined not in str(Listbox_text):
    maxop_list=[i for i in Listbox_text if find_sub_as in i]
elif find_sub_combined in str(Listbox_text):
    i=0
    multi_line_name=[]
    while i <int(kv_index[0]):
        multi_line_name.append(multiple[i])
        i+=1
    maxop_list = [p for p in Listbox_text if any(q in p for q in multi_line_name)]
else:
    maxop_list=[i for i in Listbox_text if find_sub_op in i]
find_star="*"
star_list=[i for i in Listbox_text if find_star in i]
#print("MaxOP_list is"+str(maxop_list))
#print("Star_list is"+ str(star_list))
#print("list is"+str(Start_list))
#maxop_list=[s.strip('*') for s in maxop_list]

for x in maxop_list:
    sleep(2.0)
    app.line.OK.click()
    # app.top_window().menu_select("View-> Markers-> Clear Markers")
    # try:
        # app.Clear_Markers.Yes.click()
    # except:
        # pass
    app.top_window().menu_select("Line-> Edit")
    #print("MaxOp_list is"+ str(MaxOp_list))
    #print("list is"+str(x))
    app.Line.ListBox.select(x)
    app.Line.Select_and_hide_other_lines.click()
    sleep(2.0)
    app.Line.OK.click()
    sleep(2.0)	
    app.top_window().menu_select("Line-> Reports -> Danger Tree Locator")
    tabc = app.Danger_Tree_locator.TabControl.wrapper_object()
    tabc.select(0)
    #app['Danger_Tree_locator'][u'Check vegetation grow-in violations displayed with square marker'].check()
    #app['Danger_Tree_locator'][u'Radial:is violation of total distance to wire is less than'].click() 
    #app['Danger_Tree_locator'][u'check clearance to failling trees violations displayed with circular markers'].check()
    #app.Danger_Tree_locator.Edit.set_text('0')
    #app.Danger_Tree_locator.Edit2.set_text('15')
    #app.Danger_Tree_locator.Edit3.set_text('50')
    #app.Danger_Tree_locator.Edit4.set_text('150')
    app['Danger_Tree_locator'][u'Check vegetation grow-in violations displayed with square marker'].check()
    #app['Danger_Tree_locator'][u'Check vegetation grow-in violations displayed with square marker'].click()
    app['Danger_Tree_locator'][u'Radial:is violation of total distance to wire is less than'].click() 
    app['Danger_Tree_locator'][u'check clearance to failling trees violations displayed with circular markers'].check() 
    app.Danger_Tree_locator.Edit.set_text('0')
    app.Danger_Tree_locator.Edit2.set_text('15')
    app.Danger_Tree_locator.Edit3.set_text('50')
    app.Danger_Tree_locator.Edit4.set_text('150')
    app.Danger_Tree_locator.Button4.click()
    app.Vegtation_feature_code_other_codes_ignored.Slect_None.click()
    app.Vegtation_feature_code_other_codes_ignored.ListBox.select(6)
    app.Vegtation_feature_code_other_codes_ignored.OK.click()
    tabc = app.Danger_Tree_locator.TabControl.wrapper_object()
    tabc.select(2)
    app.Danger_Tree_locator.Combobox.select("CSV file (comma separated value)")
    app.Danger_Tree_locator.Combobox2.select("List violations dense (no blank columns, slower)")
    app.Danger_Tree_locator.checkbox1.check()
    app.Danger_Tree_locator.checkbox1.click()
    app.Danger_Tree_locator.checkbox2.check()
    app.Danger_Tree_locator.checkbox2.click()
    app.Danger_Tree_locator.checkbox3.check()
    app.Danger_Tree_locator.checkbox4.check()
    app.Danger_Tree_locator.checkbox4.click()
    tabc.select(0)
    app.Danger_Tree_locator.Buttum6.click()
    app.Survey_Point_Clearance_Criteria.RadioButton.click()########## 12/3/2019
    #print("Please input the weather condition for line:"+str(x))
    ##weather_name=str(x)
    ##weather_name=weather_name.replace(" ","")
    ##weather_name=weather_name[6:9]
    ##Checkpoint=str(maxop_list)
    #if find_sub not in Checkpoint:
    #    #text = app.Line.Edit.window_text()
    #    weather_name=filter(lambda y: y in '0123456789', text_info)
    #    weather_name=weather_name[0:3]
    #degreef=(weather_value(int(weather_name))*9/5)+32
    print("Working on circuit:"+str(x))
    print("Weather info: "+str(w_info))
    print("Please select the weather condition by using the weather info showing above...") #Duke 2020
    #print("weather_name"+weather_name)
    #print("Please select the weather condition: "+str(weather_value(int(weather_name)))+" Deg C or "+str(degreef)+" Deg F and Max Sag FE ") #### 12/3/2019
    print('IMPORTANT: DO NOT CLICK ANYTHING IN PLS_CADD!!!!!')
    pause1=raw_input("After select the Max_Op weather condition, please press enter to continue: ") #Python 2
    #pause1=input("After selecting the Max_Op weather condition, please press enter to continue: ") #Python 3
    print('IMPORTANT: DO NOT CLICK ANYTHING IN PLS_CADD!!!!!') ##might have warning, select no
    app.Survey_Point_Clearance_Criteria.OK.click()
    app.Danger_Tree_locator.OK.click()
    ###app.Danger_Tree_locator.cancel.click()
    if len(maxop_list) > 1:
        # get the current line's name in multiple lines
        name_for_multi=x.split(' ')
        dash_index=[i for i, j in enumerate(name_for_multi) if '-' in j]
        dash_index[0]=int(dash_index[0])
        c_name=str(name_for_multi[dash_index[0]+1])
        c_kv=CIRCUIT_KV.split('_')
        CIRCUIT_KV=c_name+"_"+str(c_kv[-1])       
        app[u'Save_Report_Comma_Separated_value_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_Veg_Detections.csv")
    if (len(maxop_list) == 1):
        app[u'Save_Report_Comma_Separated_value_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_Veg_Detections.csv")
    #app[u'Save_Report_Comma_Separated_value_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_"+weather_name+"_Veg_Detections.csv")
    app.Save_Report_Comma_Separated_value_file.Save.click()
    try:
        app.confirm_save_as.yes.click()
    except:
        pass
    app.Vegetation_Analysis_Report.wait_not('exists', timeout=60000)
    #app.top_window().menu_select("Window->4 3D view")
    #app.top_window
    sleep(20.0)
    print('IMPORTANT: DO NOT CLICK ANYTHING IN PLS_CADD!!!!!')
    sleep(1.0)
    app.top_window().right_click_input()
    sleep(1.0)
    app.PopupMenu.menu_item(u'Autosize Font').click_input()
    sleep(1.0)
    app.PLS_CADD.Yes.click()
    sleep(1.0)
	# save pdf
    app.top_window().right_click_input()
    sleep(1.0)
    app.PopupMenu.menu_item(u'Print').click_input()
    sleep(1.0)
    app.Print.ComboBox.select(u'Microsoft Print to PDF')
    sleep(1.0)
    app.Print.OK.click()
    sleep(1.0)
    #app_pdf = Application().connect(title_re="Set output file name for PDF Architect 6")
    if len(maxop_list) > 1:
        name_for_multi=x.split(' ')
        dash_index=[i for i, j in enumerate(name_for_multi) if '-' in j]
        dash_index[0]=int(dash_index[0])
        c_name=str(name_for_multi[dash_index[0]+1])
        c_kv=CIRCUIT_KV.split('_')
        CIRCUIT_KV=c_name+"_"+str(c_kv[-1]) 
        app[u'Save_Print_Output_As']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_Veg_Detections.PDF")


    if len(maxop_list) == 1:
        app[u'Save_Print_Output_As']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_Veg_Detections.PDF")
    #app.Save_Print_Output_As.save.click()
    send_keys('{ENTER}')
    try: 
        app.confirm_save_as.yes.click()
    except:
        pass

    #pause2=raw_input("After save the pdf file press enter: ")
    sleep(2.0)
    app.top_window().menu_select("Window->4 3D view")
    #pause3=raw_input("After qucick QC the markers press enter: ")
    #sleep(1.0)
    #app.top_window().menu_select("View-> Markers-> Save Marker File")
    # try:
        # app.warning.OK.click()
    # except:
        # pass
    # if len(MaxOp_list) > 1:
        # app[u'Save_Marker_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_"+weather_name+"_Veg_Detections.MRK")

	# #app.Save_Marker_file.Save.click()
    # if len(MaxOp_list) == 1:
        # app[u'Save_Marker_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_Veg_Detections.MRK")
    # send_keys('{ENTER}')		
    # #app[u'Save_Marker_file']['Edit'].set_text(processDir+"\\"+CIRCUIT_KV+"_MaxOp_"+weather_name+"_Veg_Detections.MRK")
    # #send_keys('{ENTER}')
    # try:
        # app.confirm_save_as.yes.click()
    # except:
        # pass
	#app.Save_Marker_file.Save.click()
    sleep(1.0)
    app.top_window().menu_select("Line-> Edit")
    #Listbox_text=app.Line.ListBoxWrapper.item_texts()
    #MaxOp_list=[i for i in Listbox_text if find_sub in i]
    temp_map=map(str, star_list) #Python 2
    #temp_map=list(map(str, star_list)) #Python 3
    temp_map[0]=temp_map[0].strip("*")
    Listbox_text2=app.Line.ListBoxWrapper.item_texts()
    #print("test list in loop old"+str(Listbox_text))
    #print("test list in loop new"+str(Listbox_text2))
    app.Line.Select_and_hide_other_lines.click()
## Clean
Listbox_text3=app.Line.ListBoxWrapper.item_texts()
#print("Listbox_text3 is"+str(Listbox_text3))
find_temp="temp"
temp_list=[i for i in Listbox_text3 if find_temp in i]
#print(temp_list)
app.Line.ListBox.select(temp_list[0])
app.Line.Delete.click()
app.Delete_Line.OK.click()
##
app.Line.OK.click()
try:
    app.Line.OK.click()
except:
    pass
sleep(0.5)
app.top_window().menu_select("View-> Markers-> Clear Markers")
sleep(0.5)
try:
    app.Clear_Markers.Yes.click()
except:
    pass
print("MaxOp preparation is done...")

 