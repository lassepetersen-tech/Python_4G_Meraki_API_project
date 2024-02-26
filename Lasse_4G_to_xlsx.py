

######from requests.exceptions import InsecureRequestWarning
import datetime
import requests
import json
import xlsxwriter
import urllib3
#####
urllib3.disable_warnings()     ##<---- denne FUNKTION eksekveres her Og ignorer advarsler om certifikat fejl som untrusted f.eks.
#####
##
#
#
##################################################################################################

############################################################################################################
#
############################################################################################################
class class_date_time:
        #
        # Her nedenfor definerer jeg en masse "class variables eller class attributes som det ogsaa kaldes i Python"
        # Det er slet ikke noedvendigt at have "__init__" funktionen naar der ikke skal parses et argument fra instantiering 
        # og det er heller ikke nødvendigt med en instantiering af klassen saa laenge at de her variabler skal benyttes, 
        # saa skriver man bare klassenavn.variablenavn: 
        # se feks.:  class_date_time.
        # x = datetime.datetime.now()
        # print(x.strftime("%c"))       Tue Nov  1 12:18:02 2022
        # 
        Boolean_vinter_tid = False
        #
        var_date_time_now = datetime.datetime.now()
        #
        lowercase_date_today = var_date_time_now.strftime("%c").lower()
        #
        #
        Winter_Month_List = ["nov", "dec", "jan", "feb", "mar"]
        #
        #
        # print(Winter_Month_List)
        #
        reverse_Winter_Month_List = Winter_Month_List[::-1]
        #print(reverse_Winter_Month_List)
        #
        print("for-loop begins now via 'class_date_time:'")
        #
        for Winter_item in Winter_Month_List:
            #
            print(Winter_item)
            #
            if Winter_item in lowercase_date_today:
                Boolean_vinter_tid = True
                print("Boolean = {} and that's Winter time!" .format(Boolean_vinter_tid))
                #    
        #
        #
        def Func_date_time_Inside_Class_date_time(obj, arg2_timetallet: int):
            #
            if obj.Boolean_vinter_tid == True:
                print("Message from Func_date_time_Inside_Class_date_time: Boolean = {} and that's Winter time!" .format(obj.Boolean_vinter_tid))
                obj.vinter_time_tallet = 1 + arg2_timetallet   ################   Sommertid
                return obj.vinter_time_tallet
                #
            elif obj.Boolean_vinter_tid == False:
                print("Message from Func_date_time_Inside_Class_date_time:  Boolean = {} so it is Summer time!" .format(obj.Boolean_vinter_tid))
                obj.sommer_time_tallet = 2 + arg2_timetallet   ################   Sommertid
                return obj.sommer_time_tallet
                #
############################################################################################################
#
#
#       class_create_excel_file 
#                                og funktionen func_create_excel_file(temp_obj):
########################################################################
class class_create_excel_file:
    print("\n")
    #
    obj_date_time = class_date_time()
    date_now = str(obj_date_time.var_date_time_now)[:10] # foerste 10 elementer - 0-9
    #
    print("ser du {} saa betyder det at der er oprettet et objekt/instance af klassen 'class_create_excel_file()'" .format(date_now))
    print("og det betyder at disse statement automatisk oprettes ved instantisering af klassen!\n")
    #   
    def __init__(obj_som_argument) -> None:  # -> betyder "return None"
        # instantiering parser det nye objekt sig selv som det foerste argument
        ########################################################################
        #####################
        #
        # Magic funktionen __init__() vil altid automatisk blive eksekveret saa snart at der oprettes en instance af klassen
        # dvs uden for klassen oprettes:  nyt_obj = class_create_excel_file()
        #
        # Sagt paa en anden maade saa eksekveres alt det der er inden i funktionen __init__()
        #
        lige_nu_tid = datetime.datetime.now()    # behoever kun "obj_som_argument.ny_tid" hvis den skal bruges i en anden funktion
        #
        # "obj_som_argument." skal med paa "str_nytid" for at variablen kan benyttes nede i funktionen create_excel_file
        obj_som_argument.str_nytid = str(lige_nu_tid)[:10] + "_Time" + str(lige_nu_tid)[11:13] + "-" + str(lige_nu_tid)[14:16]  
        obj_nytid = obj_som_argument.str_nytid 
        #
        #
        #
        ### obj_som_argument. skal med for at variablen filnavn_now kan bruges udenfor funktionen __init__
        obj_som_argument.filnavn_now = f"Daily_4G_Signal_Report_{obj_nytid}.xlsx"
        #
        ################################################################################
    def func_create_excel_file(temp_obj):     ##### function create excel file #########
        ##ny_tid = datetime.datetime.now()
        ##str_nytid = str(ny_tid)[:10] + "_Time" + str(ny_tid)[11:13] + "-" + str(ny_tid)[14:16]
        #
        # "str_nytid" er allerede defineret som member i denne her samme klasse i funktionen __init__()
        # og er derfor tilgaengelig for andre funktioner i samme klasse
        #
        pointer_to_excel_file = xlsxwriter.Workbook(f"Daily_4G_Signal_Report_{temp_obj.str_nytid}.xlsx")    
        #        
        #
        filnavn_now = f"Daily_4G_Signal_Report_{temp_obj.str_nytid}.xlsx"
        print(f"Data organiseres i en fil med navnet: {filnavn_now}")
        #
        return pointer_to_excel_file
    #
    #
    #
    #
#
############################################################################################################
#
#       
#
############################################################################################################
#      Funktionen Header
############################################################################################################
def func_Header(API_Key="7777777777abcdefghij"): # her ses default argument value som er sat til random API Key
    # hvis funktionen kaldes uden argument parameter vil argument indeholde "7777777777abcdefghij"
    dict_Headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Cisco-Meraki-API-Key": API_Key
          }  
    return dict_Headers
# 
#  
#
############################################################################################################
#      Funktionen  GET request
############################################################################################################ 
def func_Get_Requests_Network_and_Cellular_List():
    min_Header = func_Header()
    #
    # jeg har sat organizations til et fiktivt nummer for sikkerhed
    #
    Network_url = "https://api.meraki.com/api/v1/organizations/123456/networks"
    Cellular_url = "https://api.meraki.com/api/v1/organizations/123456/cellularGateway/uplink/statuses"
    #
    #
    #
    #
    response_kode = 0    # initialiser koden , for hvis den allerede er 200 vil while statement blive False      
    #
    print("\n>One moment !", "\n" * 2, ">Collecting data from Meraki ! \n")          
    #
    # response_kode = 0  (for at while statement eksekveres for foerste gang og dermed at "GET Request" eksekeveres for foerste gang)
    #
    # Hvis response_kode IKKE er 200 bliver while statement True og loop eksekveres
    #
    while response_kode != 200:   ## hvis status retur kode er = 200 bliver statement False og while loop afbrydes  
            #
            try:
                Networks_Response = requests.get(Network_url, headers=min_Header, timeout=120, verify=False)
                #    # 200 betyder at authentication gik godt.            
                # nedenfor tildeles variablen "response_kode" retur status og checkes ved naeste loop
                response_kode = Networks_Response.status_code
                #
                if response_kode != 200:  
                    print(f">\n>Authentication failure from Meraki cloud ! Error: {response_kode} reived from Meraki.\n")
                    #
                    print(">According to the returned error code %n from Meraki, it could indicate the key is obsolete." % response_kode)
                    #
                    Ny_API_Key = input(f'>\n>Type a valid Meraki API authentication key: \n>')
                    #
                    #
                    # Nedenfor eksekveres funktionen med den nye Ny_API_Key som parameter som overskriver argumentets default
                    # value i funktionen func_Header()
                    #
                    #
                    # min_Header tildeles dictionary fra return statement inde i funktionen func_Header()
                    # hvor key'en "X-Cisco-Meraki-API-Key" nu indeholder argumentets value fra Ny_API_Key
                    min_Header = func_Header(Ny_API_Key) 
                    ## og ved loop benyttes dictionary key'ens "X-Cisco-Meraki-API-Key" nye value fra argumentet
                else:           
                    Networks_Response_json = Networks_Response.text           # .text er json format            
                    Networks_List = json.loads(Networks_Response_json) # konverter json til python List
                    #
                    # 
                    # Cellular Get request er lagt i else statement fordi jeg vil kun eksekvere den hvis Networks Get request er success
                    # 
                    Cellular_Response = requests.get(Cellular_url, headers=min_Header, timeout=120, verify=False)             
                    Cellular_Response_json = Cellular_Response.text           # .text er json format            
                    Cellular_List = json.loads(Cellular_Response_json) # konverter json til python List 
                    #  
                    #  nedenfor returneres der 2 forskellige variabler med hver sin List/array
                    return Networks_List, Cellular_List
                    #
                    #
            #   Hvis der godt kan oprettes en forbindelse til Meraki men hvor at retur koden er forskellig fra 200 vil
            #   while loop fortsaettes
            #
            #   exception  opstaar kun hvis der ikke kan oprettes en socket forbindelse til Meraki
            except:
                print("\nNo network connection !\n")
                return False, False
#
#
############################################################################################################
#
############################################################################################################
#
#    Funktionen  henter JSON data fra filer i hjemmekataloget
#
######################################################################## 
#
def func_Network_and_Cellular_List_from_file():   
    try:  # input fil er json response fra API GET mod Meraki
      with open("Organization-networks.json", "r") as fil_networks:
        read_networks = fil_networks.read()
      with open("Organization_cellular_uplink.json", "r") as fil_cellular:
        read_cellular = fil_cellular.read()
      #print("data type: " + type(read_networks))
      # konverterer JSON response til python list (array)
      Networks_List = json.loads(read_networks)   # konverterer fra json string format til python list/array format
      print("antallet af dictionary elementer i Networks_List: " + str(len(Networks_List)))
      Cellular_List = json.loads(read_cellular)
      print("antallet af dictionary elementer i Cellular_List: " + str(len(Cellular_List)))
      #
      return Networks_List, Cellular_List
      #
    except:
      print("mangler en input fil med json indhold fra Meraki API, skal ligge i C:-Users-lasse.petersen ")
      return False
      #
#
#
########################################################################
#

############################################################################################################

############################################################################################################
#
#
# Class sheet
#
############################################################################################################
#
class class_sheet(class_date_time):  # class_sheet arver de ting som class_date_time har
    # det svarer til extends class_date_time som i java og betyder at class_sheet har de samme properties og funktioner som i class_date_time
    #
    datoen_idag = str(class_date_time.var_date_time_now)[:10] # foerste 10 elementer - 0-9
    #
    #
    #
    def function_template_sheet(nyt_obj, country, excelfil_arg, N_List_arg, C_List_arg):
        #
        Fed_skrift = excelfil_arg.add_format({"bold": True})
        #roed_farve = excelfil_arg.add_format({"bg_color": "#951F06"})
        roed_farve = excelfil_arg.add_format({"bg_color": "#FF1F06"})
        #
        #cell_dict = excelfil_arg.add_format({"bg_color": "#FF1F06", "bold": True, "align": "center"})
        roed_og_bold_dictionary = excelfil_arg.add_format({"bg_color": "#FF1F06", "bold": True})
        #
        #
        # nyt_obj.property er kun til gavn hvis jeg skal benytte property fra denne her funktion ovre i en anden funktion i denne class
        #nyt_obj.country_sheet = excel_fil.add_worksheet(country)
        #nyt_obj.country_sheet.write(0, 0, "Store Name and EPOS")
        #nyt_obj.country_sheet.write(0, 2, "4G Signal Strength")
        #################    
        country_sheet = excelfil_arg.add_worksheet(country)  # opretter et ark for hvert land i excel filen
        #  
        country_sheet.set_column(0, 0, 35)   # (firstcolumn, lastcolumn, columnwidth) 
        country_sheet.set_column(3, 4, 28)   # (firstcolumn, lastcolumn, columnwidth)  
        country_sheet.set_column(6, 6, 15)   # (firstcolumn, lastcolumn, columnwidth)  
        country_sheet.set_column(1, 1, 28)   # (firstcolumn, lastcolumn, columnwidth)  
        #    
        country_sheet.write(0, 0, " Poor signal < -110dbm ")
        country_sheet.write(1, 0, " Ok signal: -110dbm range -103dbm ")
        country_sheet.write(2, 0, " Good signal: -102dbm range -88dbm ")
        country_sheet.write(3, 0, " Great signal: -87dbm range -77dbm ")
        country_sheet.write(4, 0, " Excellent signal > -76dbm ")
        country_sheet.write(7, 0, " Store-Location-Name and EPOS")                            # raekke 4, kolonne 0
        country_sheet.write(7, 1, " Last Reporting Date")
        country_sheet.write(7, 2, " dbm")        # raekke 4, kolonne 2   
        country_sheet.write(7, 3, " SIM card ( ICCID number)")
        country_sheet.write(7, 5, " Roaming ")
        country_sheet.write(7, 6, " APN ")
        country_sheet.write(7, 7, " Connection Type ")
        country_sheet.write(8, 0, " ################################# ")
        country_sheet.write(8, 2, " ###### ")
        country_sheet.write(8, 3, " ##################### ")
        country_sheet.write(8, 4, " ##################### ")
        country_sheet.write(8, 5, " ####### ")
        country_sheet.write(8, 6, " ####### ")
        country_sheet.write(8, 7, " ####### ")
        country_sheet.write(8, 1, " ################### ")
        ###########
        row_nr = 9
        count_No_Signal = 0
        for Networks_dict in N_List_arg:
            #              
            if country in Networks_dict["name"]:
                
                for Cellular_dict in C_List_arg:          
                    if Networks_dict["id"] == Cellular_dict["networkId"]:
                        row_nr += 1 
                        country_sheet.write(row_nr, 0, Networks_dict["name"])
                        #
                        #####print(type(Cellular_dict["lastReportedAt"]))
                        #
                        if type(Cellular_dict["lastReportedAt"]) != str:    #Has never connected to the Meraki cloud
                            country_sheet.write(row_nr, 1, "Has never connected to the Meraki cloud", roed_og_bold_dictionary)
                            country_sheet.write(row_nr, 2, "", roed_og_bold_dictionary)
                            #
                            #
                            ###############
                        elif Cellular_dict["lastReportedAt"][:10] != nyt_obj.datoen_idag:   #### datoen_idag assignes i starten af class
                            # 'lastReportedAt': '2022-10-24T11:57:13Z'
                            # [:10] = fra index 0 - 9      
                            # datoen er  index 0 -9 
                            # index 10 udelades for skal ikke have T'et med
                            # klokken begynder fra index 11 og minutter slutter med 15 dvs 16 skal angives som slut index
                            # Cellular_dict["lastReportedAt"][11:16]
                            ###### [11:] = fra index 11 og til enden
                            #print(type(Cellular_dict["lastReportedAt"]))
                            #print(Cellular_dict["lastReportedAt"][:10])
                            #print(Cellular_dict["lastReportedAt"][11:16])
                            #print(Cellular_dict["lastReportedAt"][11:13])   
                            #print(type(Cellular_dict["lastReportedAt"][11:13])) ## <class 'str'>                        
                            #print(type(int(Cellular_dict["lastReportedAt"][11:13]))) ## <class 'int'>
                            #print(int(Cellular_dict["lastReportedAt"][11:13]))
                            #print(str(2 + int(Cellular_dict["lastReportedAt"][11:13])))
                            #
                            country_sheet.write(row_nr, 1, "Unreachable since " + Cellular_dict["lastReportedAt"][:10], roed_og_bold_dictionary)
                            country_sheet.write(row_nr, 2, "", roed_og_bold_dictionary)
                            #
                            #
                            if "iccid" not in str(Cellular_dict):
                                country_sheet.write(row_nr, 3, "  (No iccid information)  ")
                            if "iccid" in str(Cellular_dict):
                                country_sheet.write(row_nr, 3, str(Cellular_dict["uplinks"][0]["iccid"]))                        
                            if "apn" in str(Cellular_dict):
                                country_sheet.write(row_nr, 6, str(Cellular_dict["uplinks"][0]["apn"]))
                            #  
                        elif Cellular_dict["lastReportedAt"][:10] == nyt_obj.datoen_idag:
                            # 'lastReportedAt': '2022-10-24T11:57:13Z'
                            # [:10] = fra index 0 - 9      
                            # datoen ligger i index 0 -9 
                            # index 10 udelades for skal ikke have T'et med
                            # klokken begynder fra index 11 og minutter slutter med 15 dvs 16 skal angives som slut index
                            # Cellular_dict["lastReportedAt"][11:16]
                            # [11:] = fra index 11 og til enden
                            #print(type(Cellular_dict["lastReportedAt"]))
                            #print("Last Report at: ")
                            #print(Cellular_dict["lastReportedAt"][:10])
                            #print(Cellular_dict["lastReportedAt"][11:16])
                            #print(Cellular_dict["lastReportedAt"][11:13])   
                            #print(type(Cellular_dict["lastReportedAt"][11:13])) ## <class 'str'>                        
                            #print(type(int(Cellular_dict["lastReportedAt"][11:13]))) ## <class 'int'>
                            #print(int(Cellular_dict["lastReportedAt"][11:13]))
                            #print(str(2 + int(Cellular_dict["lastReportedAt"][11:13])))
                            #
                            #  [13:16] = i index 13 ligger ":" tegnet og index 14 og 15 ligger minut tallet
                            str_minutter = Cellular_dict["lastReportedAt"][13:16]
                            #
                            # int(Cellular_dict["lastReportedAt"][11:13])) = konverterer til integer inden at der adderes 2 til timetallet i index 11 og 12
                            integer_timetallet = int(Cellular_dict["lastReportedAt"][11:13])
                            #str_plus2timer = str(1 + integer_timetallet)   ################   Tids forskellen!!!!!!!!!!!!!!!!!!!!!!!!
                            #str_Time = f"{str_plus2timer}{str_minutter}"
                            #
                            #
                            #
                            ##############################################################
                            # herunder eksekveres funktion " Func_date_time_Inside_Class_date_time()" som er fra "class_date_time()"
                            Dansk_timetal = nyt_obj.Func_date_time_Inside_Class_date_time(integer_timetallet)
                            #############################################################
                            #
                            DK_timer_og_minutter = f"{str(Dansk_timetal)}{str_minutter}"
                            #
                            ReportedAt = Cellular_dict["lastReportedAt"][:10] + "   # Time: " + DK_timer_og_minutter
                            #country_sheet.write(row_nr, 1, ReportedAt[:-4])  ## lastReportedAt    - De sidste fire elementer fjernes [-4]
                            country_sheet.write(row_nr, 1, ReportedAt)
                            #                
                            ##if "HerningCentret" in str(Networks_dict):   ############### store to inspect ##############
                                #print("#" * 10 + "HerningCentret")                            
                                ##print(Cellular_dict)
                                #
                            if "iccid" not in str(Cellular_dict):
                                country_sheet.write(row_nr, 3, "  (No iccid information)  ")
                                #print("#" * 50)
                                #print(Cellular_dict)
                                #print("#" * 50)
                                #print(type(Cellular_dict["uplinks"]))
                                #print("#" * 50)
                            if "signalStat" not in str(Cellular_dict):
                                country_sheet.write(row_nr, 4, "  (No Signal)  ")
                                #print("#" * 50)
                                #print(Cellular_dict)
                                #print("#" * 50)
                                #print(type(Cellular_dict["uplinks"]))
                                #print("#" * 50)
                                #
                            if "rsrp" not in str(Cellular_dict):
                                if "rscp" not in str(Cellular_dict):  ## "rscp" eller "rsrp" ikke til stede i dictionary ##
                                    count_No_Signal += 1
                                    country_sheet.write(row_nr, 4, "  (No Signal)  ")
                                    #print("#" * 50)
                                    # print(Cellular_dict)                        
                                    #print("#" * 50 + "TRUE No signal")
                            if "rsrp" in str(Cellular_dict) and Cellular_dict["uplinks"][0]["signalStat"]["rsrp"] == "":
                                country_sheet.write(row_nr, 4, "  (No Signal)  ")                            
                                ##print("  ########## value af rsrp er tom ########## ")
                                #
                                #
                                #
                            if "rscp" in str(Cellular_dict) and Cellular_dict["uplinks"][0]["signalStat"]["rscp"] == "":
                                country_sheet.write(row_nr, 4, "  (No Signal)  ")                            
                                ##print("  ########## value af rscp er tom ########## ")
                                #
                                #
                                #
                            if "rscp" in str(Cellular_dict):
                                # ["rscp"] er en key i dictionary signalStat, og viser valuen som key indeholder
                                country_sheet.write(row_nr, 2, str(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]))  # ["rscp"]
                                if float(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]) <= -111: #signal is poor
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Poor")                          
                                elif -110 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]) <= -103: #signal is OK
                                    country_sheet.write(row_nr, 4, "  Signal Strength is OK")
                                elif -102 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]) <= -88: #signal is Good
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Good")
                                elif -87 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]) <= -77: #signal is Great
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Great")
                                elif -76 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]): #signal is Excellent
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Excellent")
                                #
                            if "rsrp" in str(Cellular_dict):
                                country_sheet.write(row_nr, 2, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]))  # ["rsrp"]
                                if float(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]) <= -111: #signal is poor
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Poor")                          
                                elif -110 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]) <= -103: #signal is OK
                                    country_sheet.write(row_nr, 4, "  Signal Strength is OK")
                                elif -102 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]) <= -88: #signal is Good
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Good")
                                elif -87 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]) <= -77: #signal is Great
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Great")
                                elif -76 <= float(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]): #signal is Excellent
                                    country_sheet.write(row_nr, 4, "  Signal Strength is Excellent")
                            #
                            #
                            if "iccid" in str(Cellular_dict):
                                country_sheet.write(row_nr, 3, str(Cellular_dict["uplinks"][0]["iccid"]))
                            if "provider" in str(Cellular_dict):
                                country_sheet.write(row_nr, 5, str(Cellular_dict["uplinks"][0]["provider"])) 
                            if "apn" in str(Cellular_dict):
                                country_sheet.write(row_nr, 6, str(Cellular_dict["uplinks"][0]["apn"]))
                                #                        
                                #country_sheet.write(row_nr, 6, str(Cellular_dict["uplinks"][0])) # foerste element i array som er flere value elementer af noeglen "uplinks"
                            if "connectionType" in str(Cellular_dict):
                                country_sheet.write(row_nr, 7, str(Cellular_dict["uplinks"][0]["connectionType"]))
                            #
    #print("#" * 50)
    #print("That many times No signal: " + str(count_No_Signal))
    #print("#" * 50)
    #
    #
#
#
#
############################################################################################################
#
#
#                 Main 
############################################################################################################
#        
#
def main():   
    instance_af_class_create_excel_file = class_create_excel_file()
    #
    running_filnavn = instance_af_class_create_excel_file.filnavn_now ## property fra class_create_excel_file
    #
    #
    excelfile_pointer = instance_af_class_create_excel_file.func_create_excel_file() 
    #####################################
    #   TEST:  funktionen func_Network_and_Cellular_List_from_file() er for at teste med en lokal json fil 
    N_List, C_List = func_Network_and_Cellular_List_from_file()
    #
    # NEDENFOR GET Request
    #N_List, C_List = func_Get_Requests_Network_and_Cellular_List()    ###### Her laves https GET request ind i Meraki Cloud Storage via API kald
    #
    ##print("Her ses 'return' value fra funktionen 'func_Get_Requests_Networks()':    ", N_List)
    #
    #    
    obj_sheet = class_sheet()    # opretter nyt objekt (en instance) af klassen class_sheet()
    #
    #
    country_array = ["NL", "DK", "SE", "NO", "FI"]
    #
    if N_List == False or C_List == False:
        print("#" * 50)
        ##print("return fra 'func_Get_Requests_Networks()' N_List: ",  N_List)
        ##print("return fra 'func_Get_Requests_Cellular()' C_List: ",  C_List)
        print("\n>>>>>No network connection to Meraki Cloud<<<<<\n")
        print("#" * 50, "\n" * 1)        
        print(f'''>\n>Check your Network Connection ! \n>
                \n>Check Cisco Meraki "Maintenance service window" \n>''')
        input("\nPress any key to quit the program.> ")
        quit
    else:
        for elm in country_array:
            obj_sheet.function_template_sheet(elm, excelfile_pointer, N_List, C_List)
        #####################################################################
        try:        
            excelfile_pointer.close()
            print("\n" * 5)
            print(f"Cellular 4G report saved in the filename: {running_filnavn} in users home directory on the C: drive! \n\n")
            input("Press any key to quit the program.> ")  ## Det er for at holde DOS vinduet aaben til man trykker paa tastaturet
        except:
            print(f">The file: {running_filnavn} is locked!\n>Check file permissions or if same filename is open.\n")
            print(">Then close the excel file and run the program again.\n")
            input("\nPress any key to quit the program.> ")
            quit
#
#
if __name__ == "__main__":
    #    
    main()
    ## function_test()    
    pass
#
# 
############################### Diverse test nedenfor #####################
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
########################################################################
#
#       function_test(to argumenter)
#
def function_test(Networks_List, Cellular_List):   
    instance_excel_fil = class_create_excel_file()
    local_scope_excelfile = instance_excel_fil.func_create_excel_file()
    # for at tilgaa variabel fra en anden funktion tildeles return value til ny lokal variabel
    #
    #    
    DK_sheet = local_scope_excelfile.add_worksheet("DK")
    #
    DK_sheet.set_column(0, 2, 25)   # (firstcolumn, lastcolumn, columnwidth)
    DK_sheet.write(0, 0, "Store Name and EPOS")
    DK_sheet.write(0, 1, "4G Signal Strength in dbm")
    DK_sheet.write(0, 2, "SIM card ( ICCID number)")
    row_nr = 1
    for Networks_dict in Networks_List:
      for Cellular_dict in Cellular_List:
        if Networks_dict["id"] == Cellular_dict["networkId"]:
          # Nedenfor tester jeg om key "signalStat findes
          #try:                        
              if "DK" in Networks_dict["name"] and "Ryesgade" in str(Networks_dict):
                print("#" * 50)
                print(Cellular_dict)
                print("#" * 50)
                if "signalStat" in Cellular_dict["uplinks"][0]:    ## key "uplink" indeholder en liste/array hvori at element 0 er et dictionary med key "signalStat"
                  print(Cellular_dict["uplinks"][0])
                  print("#" * 50)
                if "rscp" in str(Cellular_dict):
                    DK_sheet.write(row_nr, 1, str(Cellular_dict["uplinks"][0]["signalStat"]["rscp"]))
                if "rsrp" in str(Cellular_dict):
                    DK_sheet.write(row_nr, 1, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]))
                    print(f'{row_nr} | {Networks_dict["name"]} | 4G signal: {str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"])} ')
                print("#" * 50)
                DK_sheet.write(row_nr, 0, Networks_dict["name"])
                print("#" * 50)
                DK_sheet.write(row_nr, 2, str(Cellular_dict["uplinks"][0]["iccid"]))
                print("#" * 50)
                print(" indhold fra signalStat key: " + str(Cellular_dict["uplinks"][0]["signalStat"]))
                print("#" * 50)
                print(" indhold fra uplink element 0 - list: " + str(Cellular_dict["uplinks"][0])) #foerste element i array som er flere value elementer af noeglen "uplinks"
                print("#" * 50)
                print(" indhold fra ## iccid ## key: " + str(Cellular_dict["uplinks"][0]["iccid"]))
                print("#" * 50)
                print(type(Cellular_dict["uplinks"][0]))
                print("#" * 50)
                row_nr += 1
          #except:
            ## continue  # continue statement skal indvendig i for-loop
    local_scope_excelfile.close()

#
#
#  # var_signalStat = {'rsrp': '-73', 'rsrq': '-7'}
##print("#" * 50)
#print(" indhold fra rsrp key: " + str(var_signalStat["rsrp"]))
#print("antal elementer i dictionary var_signalStat: " + str(len(var_signalStat)))
# 
#  
# ######################################################

# TEST:
# Networks_List element 0 og i dette element er dictionary hvor jeg skal bruge indholdet (key value) fra "name"
#
# print(Networks_List[0]['name'])
###############################################################
##################### Test #####################
#for Networks_dict in Networks_List:
#  for Cellular_dict in Cellular_List:
#    if Networks_dict["id"] == Cellular_dict["networkId"]:
#     
#      print(Cellular_dict.keys())  # viser hvilke key elementer der er i hver dictionary i Cellular_List
#



