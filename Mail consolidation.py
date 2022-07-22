# -*- coding: utf-8 -*-
"""
Created on Wed Jul 20 14:50:30 2022

@author: lovro
v 0.1.1
"""
import pandas as pd
from datetime import datetime
from pandas import ExcelWriter


def ts_clean(ts):
    if "-" in ts:
        return ts
    if "." in ts:
        return "-".join(ts.split("."))
    return ts[:2]+"-"+ts[2:4]+"-"+ts[4:]


def mailToName(mail):
    name = mail.split("@")[0]
    return " ".join(name.split("."))


colnames = ['ts', 'name', 'mail', 'op']
mail_si = pd.read_csv("mail_list.csv", names=colnames, header=None, parse_dates=['ts'], dayfirst=True)
mail_si.ts = mail_si.ts.apply(ts_clean)
mail_si.ts = pd.to_datetime(mail_si.ts, errors='coerce')
mail_si.drop("op", inplace=True, axis=1)
mail_si['lang'] = "si"
mail_si['type'] = 'ML'
mail_si['typepriority'] = 1

mail_eng = pd.read_csv("mail_list_ENG.csv", names=colnames, header=None, parse_dates=['ts'], dayfirst=True)
mail_eng.ts = pd.to_datetime(mail_eng.ts, errors='coerce')
mail_eng.drop("op", inplace=True, axis=1)
mail_eng['lang'] = "eng"
mail_eng['type'] = 'ML'
mail_eng['typepriority'] = 1

os = pd.read_excel('seznam_zavodov.ods', engine="odf")
os = os[['NAZIV', 'E-NASLOV']]
os['type'] = 'school'
os['lang'] = "si"
os['ts'] = datetime.today()
os.rename(columns={'NAZIV': 'name', "E-NASLOV": "mail"}, inplace=True)
os = os[['ts', 'name', 'mail', 'lang', 'type']]
os['typepriority'] = 2

seminar = pd.read_csv("seminar.csv", usecols=[0, 1, 2, 3], names=['ts', 'name', 'x', 'mail'], sep='\t')
seminar.drop("x", inplace=True, axis=1)
seminar['type'] = 'seminar'
seminar['lang'] = "si"
seminar['typepriority'] = 0

ML = pd.concat([seminar, mail_si, mail_eng, os])
ML = ML[['mail', 'ts', 'name', 'lang', 'type', 'typepriority']]
# for importing into my sql: email, name, lang, type

ML.mail = ML.mail.str.lower()
ML.mail = ML.mail.str.strip()
ML.name = ML.name.str.strip()

ML.sort_values(["mail", 'typepriority', 'ts'], inplace=True, ascending=[True, True, False])
ML = ML[ML['mail'].apply(len) > 2]

ML.loc[ML.name.isna(), 'name'] = ML.loc[ML.name.isna(), 'mail'].apply(mailToName)
ML.drop_duplicates(subset=['mail'], keep="first", inplace=True)
ML = ML[~ML.mail.str.contains("mail.ru")]

ML.drop(['ts', 'typepriority'], axis=1, inplace=True)
ML.reset_index(drop=True, inplace=True)

# =============================================================================
# # To excel
# =============================================================================
excel = ExcelWriter("ML archive.xlsx", engine='xlsxwriter')
ML.to_excel(excel, 'Mail List', index=False)
ML.to_csv("ML archive.csv", index=False, header=False)
excel.save()
