from pprint import pprint
import logging
from warnings import catch_warnings
import requests
import json
import subprocess
from bs4 import BeautifulSoup
from pprint import pprint
import re

with open('/home/pi/test/numeri.txt') as f:
    lista = f.readlines()

pprint(lista)
file_object = open('output_numeri.csv', 'a')


for i in lista:

    pattern1 = "([A-Z]*)\s*([0-9]{5})\s*([0-9]{2,4})"

    result1 = re.match(pattern1, i)

    if result1:
        row = result1[1] + "\t"+ result1[3]
    else:
        row = i + "\t"+ "NA"

    file_object.write(row)

file_object.close() 