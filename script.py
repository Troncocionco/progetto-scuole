import requests
from bs4 import BeautifulSoup
from pprint import pprint

#Data una lista di codici egon identificativi per una scuola, ricerca sul db del ministero ed estrapola [Indirizzo, telefono_referente, email_referente]
def cercaScuole(__sourceCodScuole = "/home/pi/test/lista_codici.txt", __output = "output.csv"):

  with open(__sourceCodScuole) as f:
      __lista = f.readlines()
  
  pprint(__lista)
  file_object = open(__output , 'a')
  
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
  
          __row = __cod + "%" + __indirizzo + "%" + __telefono + "%" + __email + "\n"
      
      except:
  
          __row = __cod + "%" + "NA" + "%" + "NA" + "%" + "NA" + "\n"
  
      file_object.write(__row)
  
  file_object.close()


