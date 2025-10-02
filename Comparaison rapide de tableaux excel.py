# -*- coding: utf-8 -*-
"""
Created on Mon Sep 29 08:26:56 2025

@author: francois.fabien
"""

import streamlit as st
import pandas as pd
import io
import numpy as np

def traitementFichier(file) :
    df = pd.read_excel(file)
    df.insert(3, "fichier", file.name)
    df["Prix"] = df["Prix"].astype('float64')
    return df

st.set_page_config(layout="wide")

st.title("Comparaison rapide de tableaux excel")

st.write('''
         Appli pour comparaison rapide des prix de différents fournisseurs en matchant par référence OEM. \n
         Chaque fichier excel doit avoir une seule feuille, avec les colonnes Ref, Ref OEM et Prix. \n
         Pour chaque ref OEM, on garde le prix le plus faible.
         ''')
         
dfList = []

uploaded_files = st.file_uploader("Importer les tableaux excel", accept_multiple_files="directory", type="xlsx")

for uploaded_file in uploaded_files :
    if uploaded_file is not None : 
        dfList.append(traitementFichier(uploaded_file))

if dfList != [] : 
    df = pd.concat(dfList, ignore_index=True)
    df = df.dropna(subset=["Prix"])
    df = df.dropna(subset=["Ref OEM"])
    df['Ref OEM'] = df['Ref OEM'].str.split("|")
    df = df.explode('Ref OEM').reset_index(drop=True)
    df['Ref OEM'] = df['Ref OEM'].map(lambda x: str(x).strip())
    df['Ref OEM'] = df['Ref OEM'].replace('', np.nan)
    df = df.dropna(subset=['Ref OEM'])
    df = df.sort_values(by=["Prix"])
    df = df.drop_duplicates(subset=["Ref OEM"], ignore_index=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer :
        df.to_excel(writer, sheet_name='Résultat comparaison', index=False)
        writer.close()

        st.download_button(label="Télécharger le résultat", data=buffer.getvalue(), file_name = "résultat comparaison prix.xlsx", mime="application/vnd.ms-excel")
