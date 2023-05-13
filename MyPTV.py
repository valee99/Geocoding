import pandas as pd
import openpyxl
import http.client, urllib.parse
import json
import requests
from socket import *
import urllib3


#from the Excel File, read the columns named "id", "indirizzo", "cap", "paese", "provincia"
df = pd.read_excel( r"FILE_PATH.xlsx",
    usecols=["id", "indirizzo", "cap", "paese", "provincia"],
)
df["id"] = df["id"].astype(str)
df["cap"] = df["cap"].astype(str)
#delete blank spaces at the end and at the beginning 
df = df.apply(lambda x: x.str.strip(), axis=1)
# delete addresses with length <= 1
df = df[df["indirizzo"].str.len() > 1]
#delete blank spaces at the end and at the beginning
df = df.apply(lambda x: x.str.strip(), axis=1)

id = df["id"]
Street = df["indirizzo"]
Locality = df["paese"]
Postal_Code = df["cap"]

#list where the desired information will be added
geos = []

for id_el, Street_el, Locality_el, Postal_Code_el in zip(id, Street, Locality, Postal_Code):
    
    #here the Country is ITALY
    #if you need to use another country, just edit Country and Country_Filter variables
    Country_Filter = "IT"
    Country = "Italy"
    indirizzo = Street_el + " " + Postal_Code_el + " " + str(Locality_el)

    url = "https://api.myptv.com/geocoding/v1/locations/by-address?country=" + Country + "&locality=" + str(Locality_el) + "&postalCode=" + Postal_Code_el + "&street=" + Street_el + "&countryFilter=" + Country_Filter
    headers = {
        'apiKey': "YOUR_MyPTV_API_KEY"
    }
    response = requests.request("GET", url, headers=headers)
    
    risposta = (response.text.encode('utf-8'))
    #trasformation into a json type
    a = risposta.decode('utf8').replace(" ' " , ' " ')
    json_output = json.loads(a)
    
    try:
        if len(json_output["locations"][0]) > 0: 

            #latitudine = latitude & longitudine = longitude
            latitudine = json_output["locations"][0]["referencePosition"]["latitude"]
            longitudine = json_output["locations"][0]["referencePosition"]["longitude"]
            geos.append([id_el , indirizzo, latitudine, longitudine])
            print("FOUND --> ", Street_el)
    except KeyError as e:
        print("errore di tipo ", e)
        geos.append([id_el , indirizzo, "", ""])
        print("NON FOUND --> ", Street_el)
    except gaierror:
        geos.append([id_el , indirizzo, "", ""])
        print("NOT FOUND --> ", Street_el)
    except (http.client.HTTPException, socket.time, socket.error):
        geos.append([id_el , indirizzo, "", ""])
        print("NOT FOUND --> ", Street_el)
        

#convert the list into a dataframe
dff = pd.DataFrame(
    geos, columns=["ID Cliente", "Indirizzo", "Latitudine", "Longitudine"]
)

#write the dataframe into a new Excel File 
dff.to_excel(r"FILE_PATH_+_EXCEL_NAME.xlsx",
    sheet_name="SHEET_NAME",
)
