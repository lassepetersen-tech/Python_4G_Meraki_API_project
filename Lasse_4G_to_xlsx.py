from ast import Break
import json
import xlsxwriter

excel_fil = xlsxwriter.Workbook("store_4G_signal.xlsx")


try:  # input fil er json response fra API GET mod Meraki
  with open("specsavers-networks.json", "r") as fil_networks:
    read_networks = fil_networks.read()
  #
  with open("specsavers_cellular_uplink.json", "r") as fil_cellular:
    read_cellular = fil_cellular.read()
  #
  Networks_List = json.loads(read_networks)   # konverterer fra json string format til python list/array format
  print("antallet af dictionary elementer i Networks_List: " + str(len(Networks_List)))
  Cellular_List = json.loads(read_cellular)
  print("antallet af dictionary elementer i Cellular_List: " + str(len(Cellular_List)))
except:
  print("mangler en input fil med json indhold fra Meraki API")
  Break


def function_test():
  DK_sheet = excel_fil.add_worksheet("DK")
  DK_sheet.set_column(0, 2, 25)   # (firstcolumn, lastcolumn, columnwidth)
  
  DK_sheet.write(0, 0, "Store Name and EPOS")
  DK_sheet.write(0, 1, "4G Signal Strenght in dbm")
  DK_sheet.write(0, 2, "4G Signal Quality")
  row_nr = 1
  for Networks_dict in Networks_List:
    for Cellular_dict in Cellular_List:
      if Networks_dict["id"] == Cellular_dict["networkId"]:
        # Nedenfor tester jeg om key "signalStat findes
        try:
          if "signalStat" in Cellular_dict["uplinks"][0]:    ## key "uplink" indeholder en liste/array hvori at element 0 er et dictionary med key "signalStat"
            if "DK" in Networks_dict["name"]:
              print(f'{row_nr} | {Networks_dict["name"]} | 4G signal: {str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"])} ')
              DK_sheet.write(row_nr, 0, Networks_dict["name"])
              DK_sheet.write(row_nr, 1, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]))
              DK_sheet.write(row_nr, 2, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrq"]))
              row_nr += 1
        except:
          continue  # continue statement skal indvendig i for-loop


class class_sheet:

  def function_template_sheet(nyt_obj, country):
    # nyt_obj.property er kun til gavn hvis jeg skal benytte property fra denne her funktion ovre i en anden funktion i denne class
    #nyt_obj.country_sheet = excel_fil.add_worksheet(country)
    #nyt_obj.country_sheet.write(0, 0, "Store Name and EPOS")
    #nyt_obj.country_sheet.write(0, 1, "4G Signal Strenght")
    #################
    country_sheet = excel_fil.add_worksheet(country)
    country_sheet.set_column(0, 0, 35)   # (firstcolumn, lastcolumn, columnwidth)  
    country_sheet.set_column(1, 2, 25)
    country_sheet.write(0, 0, "Store Name and EPOS")  # raekke 0, kolonne 0
    country_sheet.write(0, 1, "4G Signal Strenght in dbm") # raekke 0, kolonne 1
    country_sheet.write(0, 2, "4G Signal Quality")         # raekke 0, kolonne 2
    ###########
    row_nr = 1
    for Networks_dict in Networks_List:
      for Cellular_dict in Cellular_List:
        if Networks_dict["id"] == Cellular_dict["networkId"]:
          # Nedenfor tester jeg om key "signalStat findes
          try:
            if "signalStat" in Cellular_dict["uplinks"][0]:    ## key "uplink" indeholder en liste/array hvori at element 0 er et dictionary med key "signalStat"
              if country in Networks_dict["name"]:
                #print("test:  " + str(country_sheet))
                #print(f'{row_nr} | {Networks_dict["name"]} | 4G signal: {str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"])} ')
                # Nedenfor skrives records til aktuelle sheet
                country_sheet.write(row_nr, 0, Networks_dict["name"])
                country_sheet.write(row_nr, 1, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrp"]))
                country_sheet.write(row_nr, 2, str(Cellular_dict["uplinks"][0]["signalStat"]["rsrq"]))
                row_nr += 1
          except:
            continue



def main():
    obj_sheet = class_sheet()    
    country_array = ["NL", "DK", "SE", "NO", "FI"]
    for elm in country_array:
        obj_sheet.function_template_sheet(elm)
    
    excel_fil.close()
    print("exit main function")

#function_test()
#excel_fil.close()

if __name__ == "__main__":
  main()



  
############################### Diverse test nedenfor ######################
# 
var_signalStat = {'rsrp': '-73', 'rsrq': '-7'}
print(" indhold fra rsrp key: " + str(var_signalStat["rsrp"]))
print("antal elementer i dictionary var_signalStat: " + str(len(var_signalStat)))


"""
test = "abcdef"
test[0:4:1]     # test[start:end:step]
'abcd'
test = "abc12345abcdefghijklm12345abc"
>>> 
>>> test[test.index("a"):test.index("12345")]
'abc'
>>> test[test.index("a"):test.index("12345abc")]
'abc'
>>> 
"""


# 
# 
"""
      if "Cisco" in output_fra_vty:
        print("\n !!!!!!!!!!!!!!!!!!!!! \n" + " <<<<<<<<<<  Det er en Cisco switch >>>>>>>>>>>>" + "\n" * 2)
       
      else:
        print("\n !!!!!!!!!!!!!!!!!!!!! \n" * 5 + " <<<<<<<<<<  Det er IKKE en Cisco switch >>>>>>>>>>>>" + "\n" * 5)
       
"""


# 
# 
# 
# 
# 
# 
# 
# ######################################################

# TEST:  print(Networks_List[0]['name'])  # List værdi hentes via position/index-nr og dictionary værdi hentes via "key"
###############################################################
#       if Cellular_dict.get["uplinks"] != None:

"""
##################### Test #####################
for Networks_dict in Networks_List:
  for Cellular_dict in Cellular_List:
    if Networks_dict["id"] == Cellular_dict["networkId"]:
     
      print(Cellular_dict.keys())  # viser hvilke key elementer der er i hver dictionary i Cellular_List/array

     
###########################TEST####

for Networks_dict in Networks_List:
  for Cellular_dict in Cellular_List:
    if Networks_dict["id"] == Cellular_dict["networkId"]:
      
      print(Cellular_dict["model"] + Networks_dict["id"] + Cellular_dict["networkId"] + str(Cellular_dict.keys()))

###################################

################ Denne her loop skal bruges: ####################################################################

for Networks_dict in Networks_List:
  for Cellular_dict in Cellular_List:
    if Networks_dict["id"] == Cellular_dict["networkId"]:
      # Nedenfor tester jeg om key "signalStat findes
      if "signalStat" in Cellular_dict["uplinks"][0]:    ## key "uplink" indeholder en liste/array hvori at element 0 er et dictionary med key "signalStat"
        print("found a match: " + Networks_dict["name"] + str(Cellular_dict["uplinks"][0]["signalStat"]))
  


#############################################################################################################################

for Networks_dict in Networks_List:
  for Cellular_dict in Cellular_List:
    if Networks_dict["id"] == Cellular_dict["networkId"]:
      if Cellular_dict["uplinks"][0]["signalStat"]:
        print(Cellular_dict["uplinks"][0]["signalStat"])
      
      
###########################TEST#####################################################################

for Networks_dict in Networks_List:
  for Cellular_dict in Cellular_List:
    if Networks_dict["id"] == Cellular_dict["networkId"]:
      print(Cellular_dict)




################################# TEST #################################################################
#with open("./specsavers-networks.json", "r") as fil_networks:
#  read_networks = fil_networks.read(500) #kan ikke konvertere variabel fra json når ikke alle karakterer er med (her hentes kun 500 karakterer)
#  print(read_networks)
#
#with open("./specsavers_cellular_uplink.json", "r") as fil_cellular:
#  read_cellular = fil_cellular.read(800)     #kan ikke konvertere variabel fra json når ikke alle karakterer er med
#  print(read_cellular)
#
##################################################################################################

"""