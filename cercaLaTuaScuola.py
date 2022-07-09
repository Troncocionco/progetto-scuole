import os
import requests
from bs4 import BeautifulSoup
from pprint import pprint

#Data una lista di codici egon identificativi per una scuola, ricerca sul db del ministero ed estrapola [Indirizzo, telefono_referente, email_referente]
def cercaScuole(__sourceCodScuole, __output = "output.csv"):
    pprint(__sourceCodScuole)
    with open(__sourceCodScuole) as f:
        __lista = f.readlines()
    pprint(__lista)
    file_object = open(__output , 'a')
    
    __header = "CodiceScuola" + "%" + "NomePlesso" + "%" + "Indirizzo" + "%" + "TelefonoReferente" + "%" + "EmailReferente" + "%" + "SedePrincipale" + "\n"
    file_object.write(__header)

    for i in __lista:
        __cod = i.replace('\n','')
        pprint(__cod)
    
        __url='https://cercalatuascuola.istruzione.it/cercalatuascuola/ricerca/risultati?rapida='+ __cod +'&tipoRicerca=RAPIDA&gidf=1'
        #params ={'rapida': ,'tipoRicerca': 'RAPIDA', 'gidf':'1'}
        try:
            __response=requests.get(__url)
            __soup = BeautifulSoup(__response.text, 'html5lib')
            __table = __soup.find_all("tbody")
            __email = __table[0].find_all("td")[4].find("a").contents[0]
            __telefono = __table[0].find_all("td")[3].find("a").contents[0]
            __indirizzo = __table[0].find_all("td")[2].find("span").contents[0]
            __plesso = __table[0].find_all("td")[1].find("a").contents[0]
            __sedePrincipale = __table[0].find_all("td")[0].find("a").contents[0]

            __row = __cod + "%" + __plesso + "%" + __indirizzo + "%" + __telefono + "%" + __email + "%" + __sedePrincipale +"\n"     
        except:
            pprint("Ricerca fallita")
            __row = __cod + "%" + "NA" + "%" + "NA" + "%" + "NA" +"%" + "NA" +"%" + "NA" + "\n"

        file_object.write(__row)

    file_object.close()

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    
    
x = os.getcwd() + "/listaCodici.txt"
cercaScuole(x)
