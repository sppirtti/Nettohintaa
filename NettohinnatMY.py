# -*- coding: utf-8 -*-
"""
Created on Wed Mar 23 17:06:11 2022

@author: sampp
"""

print("Ohjelma Käynnistyy...")

import os, re
print("...")
import pandas as pd
print("...")
print("Alustetaan .CSV")
kokous = pd.DataFrame()

print("Haetaan dataa")
nyt = os.getcwd()
print(nyt)

os.chdir(nyt + '/data')
print(os.getcwd())

for fname in os.listdir():
    if re.search('\.xlsx$', fname):
        
        asnro = pd.read_excel(fname)
        asnro.columns = range(15)
        numero = asnro.loc[1][8]
        hinnat = pd.read_excel(fname, sheet_name="Hinnoittelu")
        
        hinnat.columns= range(27)
        netot = hinnat.loc[hinnat[17] >= 0]
        netot = netot.reset_index(drop=True)
        
        for i, rivi in netot.iterrows():  
            
            tuote = [["", str(numero),"", rivi[5], rivi[17],"EUR",rivi[20]]]
            print(tuote)
            kokous=kokous.append(tuote)
          
    
df2 = kokous.set_axis(['tyhja', 'Asiakas','tyhja2', 'Tuote', 'Nettohinta', 'Valuutta', 'yksikko'], axis=1, inplace=False)

os.chdir(nyt)
            
df2.to_csv('netotMY.csv', index=False, sep=str(";"), decimal=(","))

print("luotu netotMY.csv sijaintiin: " + str(nyt))