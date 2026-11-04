"""Hartmann Transporeon → InnoTMS. Run: streamlit run app.py."""
from __future__ import annotations

import base64
import hashlib
import io
import json
import re
from collections import defaultdict
from copy import copy
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
import pandas as pd
import requests
import streamlit as st

GOOGLE_URL = 'https://docs.google.com/spreadsheets/d/1bUWaAnw52BPFmiCJOIEtsUOmzpXtq6LVIW45jz1vmrg/edit'
DATA_SCHEMA = 'tariff-postal-rules-v3'
TEMPLATE_B64 = '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'
DEFAULT_ROUTES = [{'country_name': 'Albania', 'route': 'AL001', 'description': 'All ZIP codes'}, {'country_name': 'Austria', 'route': 'AT002', 'description': '10-19,20-24,25-28,30,31,34-37, 70-72,73'}, {'country_name': 'Austria', 'route': 'AT003', 'description': '32, 33, 38-39, 40-49'}, {'country_name': 'Austria', 'route': 'AT003A', 'description': '4550'}, {'country_name': 'Austria', 'route': 'AT003B', 'description': '4550'}, {'country_name': 'Austria', 'route': 'AT004', 'description': '50-57, 89'}, {'country_name': 'Austria', 'route': 'AT005', 'description': '60-66, 99'}, {'country_name': 'Austria', 'route': 'AT006', 'description': '67-69'}, {'country_name': 'Austria', 'route': 'AT007', 'description': '74-75, 80-88, 90-98'}, {'country_name': 'Bosnia&Hercegovina', 'route': 'BA016', 'description': 'All ZIP codes'}, {'country_name': 'Belgium', 'route': 'BE017', 'description': 'All ZIP codes'}, {'country_name': 'Bulgaria', 'route': 'BG019', 'description': '40-49'}, {'country_name': 'Bulgaria', 'route': 'BG020', 'description': '50-59'}, {'country_name': 'Bulgaria', 'route': 'BG246', 'description': '80-89'}, {'country_name': 'Bulgaria', 'route': 'BG248', 'description': '20-29'}, {'country_name': 'Bulgaria', 'route': 'BG249', 'description': '70-79'}, {'country_name': 'Bulgaria', 'route': 'BG253', 'description': '90-99'}, {'country_name': 'Bulgaria', 'route': 'BG266', 'description': '30-39'}, {'country_name': 'Bulgaria', 'route': 'BG267', 'description': '60-69'}, {'country_name': 'Belarus', 'route': 'BY021', 'description': '223043'}, {'country_name': 'Schwitzerland', 'route': 'CH022', 'description': '1000-4999'}, {'country_name': 'Schwitzerland', 'route': 'CH023', 'description': '5000-9999'}, {'country_name': 'Schwitzerland', 'route': 'CH023F', 'description': '8302'}, {'country_name': 'Czech Republic', 'route': 'CZ028', 'description': '10000-59999'}, {'country_name': 'Czech Republic', 'route': 'CZ029', 'description': '60000-79999'}, {'country_name': 'Germany', 'route': 'DE031', 'description': '01-05, 07-09'}, {'country_name': 'Germany', 'route': 'DE031A', 'description': '09221'}, {'country_name': 'Germany', 'route': 'DE031B', 'description': '09221'}, {'country_name': 'Germany', 'route': 'DE032', 'description': '10-19'}, {'country_name': 'Germany', 'route': 'DE032A', 'description': '17179'}, {'country_name': 'Germany', 'route': 'DE032B', 'description': '15741'}, {'country_name': 'Germany', 'route': 'DE033', 'description': '20, 25-27, 29'}, {'country_name': 'Germany', 'route': 'DE033B', 'description': '27243'}, {'country_name': 'Germany', 'route': 'DE034', 'description': '26'}, {'country_name': 'Germany', 'route': 'DE035', 'description': '06,30,31,32,33,38,39'}, {'country_name': 'Germany', 'route': 'DE035B', 'description': '39326'}, {'country_name': 'Germany', 'route': 'DE035F', 'description': '39326'}, {'country_name': 'Germany', 'route': 'DE036', 'description': '34,35,36,37,98,99'}, {'country_name': 'Germany', 'route': 'DE036B', 'description': '99428'}, {'country_name': 'Germany', 'route': 'DE037', 'description': '40-47, 50-52, 57-59'}, {'country_name': 'Germany', 'route': 'DE038', 'description': '48'}, {'country_name': 'Germany', 'route': 'DE039', 'description': '49'}, {'country_name': 'Germany', 'route': 'DE040', 'description': '53-56, 60-69'}, {'country_name': 'Germany', 'route': 'DE040A', 'description': '53881'}, {'country_name': 'Germany', 'route': 'DE040B', 'description': '53881'}, {'country_name': 'Germany', 'route': 'DE041', 'description': '70-79, 88-89'}, {'country_name': 'Germany', 'route': 'DE042', 'description': '81-87, 93-94'}, {'country_name': 'Germany', 'route': 'DE043', 'description': '90-92,95-97'}, {'country_name': 'Denmark', 'route': 'DK060', 'description': '1000-3699, 3800-4999'}, {'country_name': 'Denmark', 'route': 'DK061', 'description': '3700-3799'}, {'country_name': 'Denmark', 'route': 'DK062', 'description': '5000-5999'}, {'country_name': 'Denmark', 'route': 'DK063', 'description': '6000-6799'}, {'country_name': 'Denmark', 'route': 'DK063B', 'description': '6270'}, {'country_name': 'Denmark', 'route': 'DK064', 'description': '6800-8999'}, {'country_name': 'Denmark', 'route': 'DK065', 'description': '9000-9999'}, {'country_name': 'Denmark', 'route': 'DK066', 'description': '6070'}, {'country_name': 'Denmark', 'route': 'DK068', 'description': '7700'}, {'country_name': 'Denmark', 'route': 'DK069', 'description': '9560'}, {'country_name': 'Estonia', 'route': 'EE070', 'description': '67001'}, {'country_name': 'Estonia', 'route': 'EE071', 'description': '68315, 63218'}, {'country_name': 'Estonia', 'route': 'EE072', 'description': '74201'}, {'country_name': 'Spain', 'route': 'ES073', 'description': '08, 17, 22, 25, 43'}, {'country_name': 'Spain', 'route': 'ES074', 'description': '01, 12, 20, 26, 31, 42, 44, 48, 50'}, {'country_name': 'Spain', 'route': 'ES075', 'description': '09, 16, 19, 28, 34, 39, 40, 46, 47'}, {'country_name': 'Spain', 'route': 'ES076', 'description': '02,03,05,10,13,15,24,27,32,33,36,37,4549'}, {'country_name': 'Spain', 'route': 'ES077', 'description': '04,06,11,14,18,21,23,29,30,41'}, {'country_name': 'Spain', 'route': 'ES078', 'description': '07'}, {'country_name': 'Finland', 'route': 'FI079', 'description': '00-09'}, {'country_name': 'Finland', 'route': 'FI080', 'description': '10-19'}, {'country_name': 'Finland', 'route': 'FI081', 'description': '20-21, 23-39, 45-49, 53-56'}, {'country_name': 'Finland', 'route': 'FI081A', 'description': '20760, 23100, 23800'}, {'country_name': 'Finland', 'route': 'FI452', 'description': '32200'}, {'country_name': 'Finland', 'route': 'FI082', 'description': '50-52, 57-59, 76-79'}, {'country_name': 'Finland', 'route': 'FI083', 'description': '22, 40-44, 60-75, 80-83'}, {'country_name': 'Finland', 'route': 'FI084', 'description': '84-95'}, {'country_name': 'Finland', 'route': 'FI085', 'description': '96-99'}, {'country_name': 'France', 'route': 'FR086', 'description': '02, 08, 59, 60, 62, 76, 80'}, {'country_name': 'France', 'route': 'FR087', 'description': '14, 35, 50, 53, 61, 72'}, {'country_name': 'France', 'route': 'FR088', 'description': '10,27,28,45,51,75,77,78,89,91-95, 89'}, {'country_name': 'France', 'route': 'FR088B', 'description': '51310, 51600'}, {'country_name': 'France', 'route': 'FR088F', 'description': '51310'}, {'country_name': 'France', 'route': 'FR089', 'description': '52, 54, 55, 57, 67, 68, 70, 88, 90'}, {'country_name': 'France', 'route': 'FR090', 'description': '01, 21, 25, 39, 42, 69, 71, 74'}, {'country_name': 'France', 'route': 'FR091', 'description': '03, 15, 18, 19, 23, 36, 43, 58, 63, 87'}, {'country_name': 'France', 'route': 'FR092', 'description': '16, 17, 37, 41, 44, 49, 79, 85, 86'}, {'country_name': 'France', 'route': 'FR093', 'description': '04-06, 13, 26, 38, 73, 83, 84'}, {'country_name': 'France', 'route': 'FR094', 'description': '07, 09, 11, 12, 20, 30, 31, 34, 48, 66, 81'}, {'country_name': 'France', 'route': 'FR095', 'description': '24, 32, 33, 40, 46, 47, 64, 65, 82'}, {'country_name': 'France', 'route': 'FR111', 'description': '72650'}, {'country_name': 'France', 'route': 'FR265', 'description': '22,29,56'}, {'country_name': 'Great Britain', 'route': 'GB116', 'description': 'CT, MG, TN, BN, RH, PO, GU, SO, RG, OX, MK, NN, CB, SG, LU, AL, WD, EN, CM, SL, HA, UB, TW, IG, RM, SS, KT, SM, CR, BR, DA, IP'}, {'country_name': 'Great Britain', 'route': 'GB117', 'description': 'PE, LE, DE, NG, S, DN, HU, NR, LN'}, {'country_name': 'Great Britain', 'route': 'GB117A', 'description': 'NR17 2QZ'}, {'country_name': 'Great Britain', 'route': 'GB118', 'description': 'EX, PL,TA, TQ, TR'}, {'country_name': 'Great Britain', 'route': 'GB119', 'description': 'DG, ML, TD, EH, G, KA'}, {'country_name': 'Great Britain', 'route': 'GB120', 'description': 'AB, DD, FK, IV, KW, KY, PA, PH, ZE'}, {'country_name': 'Great Britain', 'route': 'GB120A', 'description': 'AB43 6RQ'}, {'country_name': 'Great Britain', 'route': 'GB121', 'description': 'BT'}, {'country_name': 'Great Britain', 'route': 'GB126', 'description': 'DD8 3EW'}, {'country_name': 'Great Britain', 'route': 'GB127', 'description': 'EH45 8QZ'}, {'country_name': 'Great Britain', 'route': 'GB128', 'description': 'KY1 4DX'}, {'country_name': 'Great Britain', 'route': 'GB129', 'description': 'LN6 9HA  (North Scarle)'}, {'country_name': 'Great Britain', 'route': 'GB130', 'description': 'MELKSHAM.WILTS.SN12 6SP'}, {'country_name': 'Great Britain', 'route': 'GB260', 'description': 'BH, DT, SP, BA, BS, SN, GL, HR, WR, CV, B, DY, WV, WS, TF, ST'}, {'country_name': 'Great Britain', 'route': 'GB260A', 'description': 'WS11 1RR'}, {'country_name': 'Great Britain', 'route': 'GB261', 'description': 'LS, YO, HG, BD, BB,  PR, FY, TS, DL, LA, SR, DH, CA, NE, IM, L, CW, SK, WA, OL, CH, BL, WN, HX, M'}, {'country_name': 'Great Britain', 'route': 'GB262', 'description': 'CF, NP, LD, SY, LL, SA'}, {'country_name': 'Great Britain', 'route': 'GB281', 'description': 'OX29 7RB Standlake'}, {'country_name': 'Greece', 'route': 'GR133', 'description': '20-29'}, {'country_name': 'Greece', 'route': 'GR134', 'description': '50-69'}, {'country_name': 'Greece', 'route': 'GR250', 'description': '30-49'}, {'country_name': 'Croatia', 'route': 'HR120', 'description': '48 000'}, {'country_name': 'Croatia', 'route': 'HR135', 'description': '48 000'}, {'country_name': 'Croatia', 'route': 'HR270', 'description': '10, 40-49'}, {'country_name': 'Croatia', 'route': 'HR271', 'description': '31-35'}, {'country_name': 'Hungary', 'route': 'HR110', 'description': '2941'}, {'country_name': 'Hungary', 'route': 'HU136', 'description': '2941'}, {'country_name': 'Hungary', 'route': 'HU244', 'description': '30-39'}, {'country_name': 'Hungary', 'route': 'HU245', 'description': '90-99'}, {'country_name': 'Hungary', 'route': 'HU272', 'description': '10-19'}, {'country_name': 'Hungary', 'route': 'HU273', 'description': '20-29'}, {'country_name': 'Hungary', 'route': 'HU275', 'description': '40-49'}, {'country_name': 'Hungary', 'route': 'HU276', 'description': '50-59'}, {'country_name': 'Hungary', 'route': 'HU277', 'description': '60-69'}, {'country_name': 'Hungary', 'route': 'HU278', 'description': '9073'}, {'country_name': 'Hungary', 'route': 'HU279', 'description': '80-89'}, {'country_name': 'Ireland', 'route': 'IE137', 'description': 'Dublin,Belfast'}, {'country_name': 'Ireland', 'route': 'IE138', 'description': 'Antrim,Armagh,Down,Cavan,Louth,Longford,Meath,Westmeath,Offaly,Kildare,Laois,Wicklow,Carlow,Kilkenny'}, {'country_name': 'Ireland', 'route': 'IE139', 'description': 'Derry,Tyrone,Fermanagh,Sligo,Leitrim,Roscommon,Galway,Claire,Tipperary,Limerick,Wexford,Waterford,Cork'}, {'country_name': 'Ireland', 'route': 'IE140', 'description': 'Donegal,Mayo,Kerry'}, {'country_name': 'Ireland', 'route': 'IE141', 'description': 'Silver Stream, Tyholland'}, {'country_name': 'Ireland', 'route': 'IE142', 'description': 'Monaghan'}, {'country_name': 'Ireland', 'route': 'IE121', 'description': 'BT'}, {'country_name': 'Italy', 'route': 'IT143', 'description': '0000-06000'}, {'country_name': 'Italy', 'route': 'IT144', 'description': '10000-19999'}, {'country_name': 'Italy', 'route': 'IT145', 'description': '30000-39999'}, {'country_name': 'Italy', 'route': 'IT145F', 'description': '35010'}, {'country_name': 'Italy', 'route': 'IT146', 'description': '50000-59999'}, {'country_name': 'Italy', 'route': 'IT147', 'description': '60000-69999'}, {'country_name': 'Italy', 'route': 'IT148', 'description': '70000-86999'}, {'country_name': 'Italy', 'route': 'IT149', 'description': '87000-89999'}, {'country_name': 'Italy', 'route': 'IT150', 'description': '90000-99999'}, {'country_name': 'Italy', 'route': 'IT151', 'description': '07000-09999'}, {'country_name': 'Italy', 'route': 'IT255', 'description': '20000-29999'}, {'country_name': 'Italy', 'route': 'IT256', 'description': '40000-49999'}, {'country_name': 'Lithauen', 'route': 'LT157', 'description': 'All ZIP codes'}, {'country_name': 'Latvia', 'route': 'LV158', 'description': 'All ZIP codes'}, {'country_name': 'Moldova', 'route': 'MD159', 'description': 'All ZIP codes'}, {'country_name': 'Montenegro', 'route': 'ME247', 'description': 'All ZIP codes'}, {'country_name': 'Macedonia', 'route': 'MK160', 'description': 'All ZIP codes'}, {'country_name': 'Holland', 'route': 'NL161', 'description': 'All ZIP codes'}, {'country_name': 'Holland', 'route': 'NL161B', 'description': '7030'}, {'country_name': 'Norway', 'route': 'NO166', 'description': '1-20, 32'}, {'country_name': 'Norway', 'route': 'NO167', 'description': '21-22, 30-31, 33-39, 46-48'}, {'country_name': 'Norway', 'route': 'NO168', 'description': '23-29, 40, 43-45, 49'}, {'country_name': 'Norway', 'route': 'NO169', 'description': '41-42, 50-52, 55-56, 58'}, {'country_name': 'Norway', 'route': 'NO170', 'description': '53-54, 59-66, 70, 72-73, 75'}, {'country_name': 'Norway', 'route': 'NO171', 'description': '67-69, 71, 76-79'}, {'country_name': 'Norway', 'route': 'NO172', 'description': '80-94'}, {'country_name': 'Norway', 'route': 'NO173', 'description': '95-99'}, {'country_name': 'Norway', 'route': 'NO174', 'description': '1400'}, {'country_name': 'Norway', 'route': 'NO175', 'description': '1891'}, {'country_name': 'Norway', 'route': 'NO177', 'description': '4367'}, {'country_name': 'Norway', 'route': 'NO178', 'description': '7080'}, {'country_name': 'Poland', 'route': 'PL179', 'description': '20-29, 30-39'}, {'country_name': 'Poland', 'route': 'PL180', 'description': '00-08, 10-19'}, {'country_name': 'Poland', 'route': 'PL180B', 'description': '05-311, 05-806'}, {'country_name': 'Poland', 'route': 'PL181', 'description': '09'}, {'country_name': 'Poland', 'route': 'PL182', 'description': '40-44-49, 90-99'}, {'country_name': 'Poland', 'route': 'PL182B', 'description': '48-317'}, {'country_name': 'Poland', 'route': 'PL183', 'description': '43'}, {'country_name': 'Poland', 'route': 'PL183B', 'description': '58-312'}, {'country_name': 'Poland', 'route': 'PL184', 'description': '50-59, 60-65-69'}, {'country_name': 'Poland', 'route': 'PL185', 'description': '62'}, {'country_name': 'Poland', 'route': 'PL186', 'description': '63'}, {'country_name': 'Poland', 'route': 'PL187', 'description': '64'}, {'country_name': 'Poland', 'route': 'PL188', 'description': '70-79, 80-89'}, {'country_name': 'Portugal', 'route': 'PT263', 'description': '2'}, {'country_name': 'Portugal', 'route': 'PT264', 'description': '3'}, {'country_name': 'Romania', 'route': 'RO192', 'description': '0 - 24'}, {'country_name': 'Romania', 'route': 'RO192B', 'description': '137165'}, {'country_name': 'Romania', 'route': 'RO193', 'description': '30 - 33'}, {'country_name': 'Romania', 'route': 'RO194', 'description': '12, 60 - 62, 70 - 73, 80 - 81'}, {'country_name': 'Romania', 'route': 'RO257', 'description': '40 - 45'}, {'country_name': 'Romania', 'route': 'RO257B', 'description': '417598'}, {'country_name': 'Romania', 'route': 'RO258', 'description': '50 - 55'}, {'country_name': 'Romania', 'route': 'RO259', 'description': '82, 90 - 92'}, {'country_name': 'Republic Serbia', 'route': 'RS200', 'description': '21-25'}, {'country_name': 'Republic Serbia', 'route': 'RS200B', 'description': '21471'}, {'country_name': 'Republic Serbia', 'route': 'RS201', 'description': '11, 13, 15'}, {'country_name': 'Republic Serbia', 'route': 'RS202', 'description': '19, 34, 36, 37'}, {'country_name': 'Republic Serbia', 'route': 'RS203', 'description': '16, 18'}, {'country_name': 'Sweden', 'route': 'SE210', 'description': '20-26'}, {'country_name': 'Sweden', 'route': 'SE211', 'description': '27-29'}, {'country_name': 'Sweden', 'route': 'SE211A', 'description': '28433'}, {'country_name': 'Sweden', 'route': 'SE212', 'description': '30-32'}, {'country_name': 'Sweden', 'route': 'SE213', 'description': '33-39'}, {'country_name': 'Sweden', 'route': 'SE213A', 'description': '38473'}, {'country_name': 'Sweden', 'route': 'SE214', 'description': '40-47, 50-54'}, {'country_name': 'Sweden', 'route': 'SE215', 'description': '55-59'}, {'country_name': 'Sweden', 'route': 'SE216', 'description': '65-67'}, {'country_name': 'Sweden', 'route': 'SE217', 'description': '60-64, 69-70'}, {'country_name': 'Sweden', 'route': 'SE218', 'description': '71-73'}, {'country_name': 'Sweden', 'route': 'SE219', 'description': '10-19, 74-75'}, {'country_name': 'Sweden', 'route': 'SE220', 'description': '68, 77-78'}, {'country_name': 'Sweden', 'route': 'SE221', 'description': '79, 82'}, {'country_name': 'Sweden', 'route': 'SE222', 'description': '80-81'}, {'country_name': 'Sweden', 'route': 'SE223', 'description': '84-86'}, {'country_name': 'Sweden', 'route': 'SE224', 'description': '83, 87-89'}, {'country_name': 'Sweden', 'route': 'SE225', 'description': '90-94'}, {'country_name': 'Sweden', 'route': 'SE226', 'description': '95-98'}, {'country_name': 'Slovenia', 'route': 'SI231', 'description': '20-29, 90'}, {'country_name': 'Slovenia', 'route': 'SI232', 'description': '10-19, 30-39, 80-89'}, {'country_name': 'Slovenia', 'route': 'SI233', 'description': '40-49, 50-59, 60-69'}, {'country_name': 'Slovakia', 'route': 'SK235', 'description': '80-99'}, {'country_name': 'Slovakia', 'route': 'SK236', 'description': '00-10'}, {'country_name': 'Ukrain', 'route': 'UA239', 'description': '03138,08606, 07534, 09400, 02089,19603'}, {'country_name': 'Ukrain', 'route': 'UA240', 'description': '32325'}, {'country_name': 'Ukrain', 'route': 'UA241', 'description': '62361'}, {'country_name': 'Ukrain', 'route': 'UA242', 'description': '57300, 75024'}, {'country_name': 'Ukrain', 'route': 'UA243', 'description': '94320, 97560'}, {'country_name': 'Kosovo', 'route': 'XK001', 'description': 'Kosovo'}]
SOURCE_FIELDS = ['Fuvar sz.', 'Szállítmánysz.', 'Szállítmány-ID', 'Súly', 'Súly egysége',
                 'Mennyiség', 'Térfogat egysége', 'Cég neve - Kirakodási hely',
                 'Ország - Kirakodási hely', 'Postai irányítószám - Kirakodási hely',
                 'Berakodás dátuma ettől', 'Berakodás dátuma eddig',
                 'Kirakodás dátuma (ettől)', 'Kirakodás dátuma (eddig)']
DATE_FIELDS = ['Berakodás dátuma ettől', 'Berakodás dátuma eddig',
               'Kirakodás dátuma (ettől)', 'Kirakodás dátuma (eddig)']
ADDITIVE = ['Súly', 'Mennyiség', 'Rakodási egységek']
IDENTIFIERS = ['Szállítmánysz.', 'Szállítmány-ID']
ZIP_WIDTH = {'AT': 4, 'BE': 4, 'CH': 4, 'DE': 5, 'FR': 5, 'IT': 5,
             'NL': 4, 'HU': 4, 'CZ': 5, 'SK': 5, 'PL': 5}


def text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def number(value):
    if isinstance(value, bool) or value is None or text(value) == '':
        raise ValueError('Hiányzó számérték.')
    raw = str(value).replace('\u00a0', '').replace(' ', '').replace(',', '.')
    try:
        result = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f'Érvénytelen szám: {value}') from exc
    if not result.is_finite():
        raise ValueError(f'Érvénytelen szám: {value}')
    return result


def timestamp(value):
    if value is None or text(value) == '':
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time())
    try:
        return datetime.fromisoformat(text(value))
    except ValueError as exc:
        raise ValueError(f'Nem értelmezhető dátum: {value}') from exc


def read_rows(data, required, preferred=None):
    book = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        sheets = sorted(book.worksheets, key=lambda s: s.title != preferred) if preferred else book.worksheets
        for sheet in sheets:
            # Transporeon declares A1:A1 even when the XML contains the full table.
            sheet.reset_dimensions()
            iterator = iter(sheet.values)
            for _ in range(20):
                row = next(iterator, None)
                if row is None:
                    break
                headers = [text(v) for v in row]
                if set(required).issubset(headers):
                    records = [dict(zip(headers, r)) for r in iterator if any(text(v) for v in r)]
                    return records, sheet.title
    finally:
        book.close()
    raise ValueError('Nem található megfelelő munkalap. Elvárt oszlopok: ' + ', '.join(required))


def aggregate(data):
    records, sheet = read_rows(data, SOURCE_FIELDS)
    groups = defaultdict(list)
    seen = set()
    for row in records:
        key = text(row['Fuvar sz.'])
        if not key:
            raise ValueError('A forrásban üres fuvarszám szerepel.')
        identity = text(row['Szállítmány-ID']) or text(row['Szállítmánysz.'])
        if not identity:
            raise ValueError(f'{key}: hiányzó szállítmányazonosító.')
        if (key, identity) in seen:
            raise ValueError(f'{key}: a(z) {identity} szállítmány többször szerepel. A forrást ellenőrizni kell.')
        seen.add((key, identity))
        groups[key].append(row)
    result = []
    ignored = set(IDENTIFIERS + ADDITIVE + DATE_FIELDS)
    for key, rows in groups.items():
        errors = []
        sums = {}
        for name in ADDITIVE:
            try:
                sums[name] = sum((number(r.get(name)) for r in rows), Decimal(0))
                if sums[name] < 0:
                    errors.append(f'Negatív érték: {name}')
            except ValueError as exc:
                sums[name] = None
                errors.append(f'{name}: {exc}')
        if {text(r.get('Súly egysége')).lower() for r in rows} != {'kg'}:
            errors.append('A súly mértékegysége nem egységes kg.')
        if {text(r.get('Térfogat egysége')).lower() for r in rows} != {'cbm'}:
            errors.append('A térfogat mértékegysége nem egységes cbm.')
        dates, conflicts = {}, []
        for label, fields in [('Berakodás', DATE_FIELDS[:2]), ('Kirakodás', DATE_FIELDS[2:])]:
            try:
                pairs = [(timestamp(r.get(fields[0])), timestamp(r.get(fields[1]))) for r in rows]
                dates[label] = sorted({d.date() for pair in pairs for d in pair if d})
                if len(set(pairs)) > 1:
                    conflicts.append(label)
            except ValueError as exc:
                dates[label] = []
                errors.append(str(exc))
        different = [name for name in rows[0] if name not in ignored
                     and len({text(r.get(name)) for r in rows}) > 1]
        result.append({'id': key, 'rows': rows, 'sums': sums, 'dates': dates,
                       'date_conflicts': conflicts, 'different': different, 'errors': errors})
    return result, len(records), sheet


def read_routes(data=None):
    if data is None:
        return DEFAULT_ROUTES
    rows, _ = read_rows(data, ['Country', 'Route', 'Description (ZIP codes)'], 'Routes Overview (Road)')
    return [{'country_name': text(r['Country']), 'route': text(r['Route']).upper(),
             'description': text(r['Description (ZIP codes)'])} for r in rows if text(r['Route'])]


def normalize_zip(value, country):
    raw = text(value).upper()
    raw = re.sub(r'^[A-Z]{1,3}\s*-\s*', '', raw)
    compact = re.sub(r'[\s-]', '', raw)
    if compact.isdigit():
        compact = compact.zfill(ZIP_WIDTH.get(country, len(compact)))
    return compact


def exact_zip_matches(description, postal, country):
    """Match complete postal codes in lists, without padding short area prefixes.

    Numeric ranges must not become exact ZIPs when punctuation is removed:
    e.g. Austrian 10-19 is a range, not the complete postal code 1019.
    Polish 05-311 is a complete postal code, as are spaced codes such as 48 000.
    """
    for part in re.split(r'[,;]', text(description).upper()):
        token = part.strip().replace('–', '-').replace('—', '-')
        token = re.sub(r'^' + re.escape(country) + r'\s*-\s*', '', token)
        if '-' in token and not (country == 'PL' and re.fullmatch(r'\d{2}-\d{3}', token)):
            continue
        if not re.fullmatch(r'[A-Z0-9\s-]+', token):
            continue
        if re.sub(r'[\s-]', '', token) == postal:
            return True
    return False


def zip_matches(description, postal, country):
    """ZIP prefixes and inclusive ranges; unknown patterns never get guessed."""
    raw = text(description).upper().replace('–', '-').replace('—', '-')
    if exact_zip_matches(raw, postal, country):
        return True
    if raw.startswith('ALL ZIP'):
        return True
    for part in re.split(r'[,;]', raw):
        part = part.strip()
        interval = re.fullmatch(r'(\d+)\s*-\s*(\d+)', part)
        if interval:
            left, right = interval.groups()
            if len(left) == len(right) and postal[:len(left)].isdigit():
                if int(left) <= int(postal[:len(left)]) <= int(right):
                    return True
            elif country == 'IT' and postal.isdigit() and max(len(left), len(right)) >= 4:
                if int(left) <= int(postal) <= int(right):
                    return True
        else:
            token = re.sub(r'\s+', '', part)
            if token and re.fullmatch(r'[A-Z0-9]+', token):
                # British outward letter areas must stop at a digit, not another letter.
                if token.isalpha():
                    if re.match(re.escape(token) + r'(?=\d|$)', postal):
                        return True
                elif postal.startswith(token):
                    return True
    return False


def parse_rates(values):
    header_index = next((i for i, row in enumerate(values[:20])
                         if 'Sales Freight Route Name' in row and 'In CBM Interval' in row), None)
    if header_index is None:
        raise ValueError('A Rates munkalap fejléce nem található.')
    headers = [text(v) for v in values[header_index]]
    route_col = headers.index('Sales Freight Route Name')
    band_col = headers.index('In CBM Interval')
    # Use the last populated Price column globally. Never fall back per row.
    price_cols = [i for i, h in enumerate(headers) if h.lower().startswith('price')
                  and any(len(r) > i and text(r[i]) for r in values[header_index + 1:])]
    if not price_cols:
        raise ValueError('Nincs kitöltött Price ároszlop.')
    price_col = price_cols[-1]
    table, blocks = {}, []
    current = None
    for excel_row, row in enumerate(values[header_index + 1:], header_index + 2):
        if len(row) <= max(route_col, band_col):
            continue
        route, band = text(row[route_col]).upper(), text(row[band_col])
        if not route or not re.fullmatch(r'\d+(?:[.,]\d+)?\s*-\s*(?:\d+(?:[.,]\d+)?)?', band):
            continue
        loading = text(row[headers.index('Loading')]).upper()
        unloading = text(row[headers.index('Unloading')]).upper()
        # The agreed 10 / 20 / ... / 70 cutoffs use the upper endpoints;
        # 10.001 belongs to the next band without rounding the shipment volume.
        lower, upper = re.split(r'\s*-\s*', band)
        upper_value = number(upper) if upper else None
        band_key = str(upper_value) if upper_value is not None else 'FTL'
        group_key = (loading, unloading, route)
        if current is None or current['key'] != group_key or band_key in current['entries']:
            current = {'key': group_key, 'entries': {}, 'descriptions': []}
            blocks.append(current)
        description = text(row[headers.index('Description (ZIP codes)')])
        if description:
            current['descriptions'].append(description)
        try:
            price = number(row[price_col]) if len(row) > price_col else None
        except ValueError:
            price = None
        current['entries'][band_key] = {'price': price, 'band': band, 'upper': upper_value, 'row': excel_row}
    totals = defaultdict(int)
    for block in blocks:
        totals[block['key']] += 1
    indices = defaultdict(int)
    for block in blocks:
        loading, unloading, route = block['key']
        indices[block['key']] += 1
        code = route if totals[block['key']] == 1 else f'{route} [{indices[block["key"]]}]'
        description = ' / '.join(block['descriptions'])
        for band_key, entry in block['entries'].items():
            table[(loading, unloading, code, band_key)] = {
                **entry, 'description': description, 'base_route': route,
                'postal_descriptions': tuple(block['descriptions'])}
    if not table:
        raise ValueError('A díjtáblában nincs feldolgozható díjsor.')
    return table, headers[price_col], openpyxl.utils.get_column_letter(price_col + 1)


def read_rate_xlsx(data):
    book = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        if 'Rates' not in book.sheetnames:
            raise ValueError('A feltöltött díjfájlból hiányzik a Rates munkalap.')
        book['Rates'].reset_dimensions()
        return [list(r) for r in book['Rates'].values]
    finally:
        book.close()


def fetch_google(url, credentials=None):
    parsed = urlparse(url)
    match = re.search(r'/spreadsheets/d/([A-Za-z0-9_-]+)', parsed.path)
    if parsed.hostname != 'docs.google.com' or not match:
        raise ValueError('Érvényes docs.google.com/spreadsheets hivatkozást adj meg.')
    spreadsheet_id = match.group(1)
    if credentials:
        from google.oauth2.service_account import Credentials
        from google.auth.transport.requests import Request
        auth = Credentials.from_service_account_info(dict(credentials), scopes=[
            'https://www.googleapis.com/auth/spreadsheets.readonly'])
        auth.refresh(Request())
        response = requests.get(
            f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Rates!A1:ZZ5000',
            params={'valueRenderOption': 'UNFORMATTED_VALUE'},
            headers={'Authorization': f'Bearer {auth.token}'}, timeout=30)
        response.raise_for_status()
        return response.json().get('values', [])
    response = requests.get(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq',
                            params={'tqx': 'out:json', 'sheet': 'Rates', 'headers': '1'}, timeout=30)
    response.raise_for_status()
    if '<html' in response.text[:1000].lower() or 'accounts.google.com' in response.url:
        raise ValueError('A Google-táblázat hitelesítést kér.')
    # The JSON 'v' member holds the full value; formatted CSV could round prices.
    payload = response.content.decode('utf-8-sig')
    start, end = payload.find('{'), payload.rfind('}')
    if start < 0 or end < start:
        raise ValueError('A Google-válasz nem értelmezhető táblázatként.')
    table = json.loads(payload[start:end + 1]).get('table')
    if not table:
        raise ValueError('A Google-válaszban nincs olvasható Rates munkalap.')
    return [[c.get('label', '') for c in table['cols']]] + [
        [c.get('v') if c else None for c in r['c']] for r in table['rows']]


def supplementary_postal_rules(description, country):
    """Read ZIP rules from the Rates description, never monetary amounts.

    Structured ZIP lists/ranges remain rules. In prose, only an explicit full
    ZIP at the start of a list item is accepted, e.g. '06049 Perugia 1300 EUR'.
    A following currency/distance/unit makes a bare number ineligible as a ZIP.
    """
    raw = text(description).upper().replace('–', '-').replace('—', '-')
    if raw.startswith('ALL ZIP'):
        return ['All ZIP codes']
    rules = []
    width = ZIP_WIDTH.get(country)
    # Mixed-type Google columns may return numeric cells as Hungarian strings.
    # Never split 2,50 into ZIP prefixes 2 and 50. Only complete ZIP codes are
    # accepted as standalone supplemental numbers; area prefixes stay in master.
    decimal = re.fullmatch(r'[+-]?(\d+)[.,](\d+)', raw)
    if decimal and ('.' in raw or width is None or
                    any(len(part) != width for part in decimal.groups())):
        return []
    if re.fullmatch(r'[+-]?\d+(?:[.,]\d+)?\s*(?:EUR|HUF|€|KM|KG|CBM|M3|%)', raw):
        return []
    for item in re.split(r'[,;]', raw):
        item = re.sub(r'^' + re.escape(country) + r'\s*-\s*', '', item.strip())
        if re.fullmatch(r'\d+\s*-\s*\d+', item):
            rules.append(item)
        elif re.fullmatch(r'\d+', item):
            if width is not None and len(item) == width:
                rules.append(item)
        elif width:
            match = re.match(r'^(\d{' + str(width) + r'})(?![A-Z0-9])', item)
            if match:
                suffix = item[match.end():]
                if not re.match(r'^\s*(?:EUR\b|HUF\b|€|KM\b|KG\b|CBM\b|M3\b|%)', suffix):
                    rules.append(match.group(1))
        elif country == 'GB':
            match = re.match(r'^([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})(?![A-Z0-9])', item)
            if match:
                rules.append(match.group(1))
    return rules


def postal_match_strength(description, postal, country):
    if exact_zip_matches(description, postal, country):
        return 3
    if not zip_matches(description, postal, country):
        return 0
    return 1 if text(description).upper().startswith('ALL ZIP') else 2


def candidate_routes(routes, rates, loading, country, postal):
    # Rates defines eligible lanes. Reference-only codes (e.g. IT145F, CH023F)
    # cannot suppress a real tariff for this loading/unloading country pair.
    eligible = defaultdict(list)
    for key, value in rates.items():
        if key[:2] == (loading, country):
            eligible[key[2]].append(value)
    master_rules = defaultdict(list)
    for row in routes:
        master_rules[row['route']].append(row['description'])
    scores = {}
    for code, entries in eligible.items():
        base = entries[0].get('base_route', code.split(' [')[0])
        rules = list(master_rules.get(base, []))
        descriptions = {desc for entry in entries for desc in entry.get('postal_descriptions', ())}
        for description in descriptions:
            rules.extend(supplementary_postal_rules(description, country))
        scores[code] = max((postal_match_strength(rule, postal, country) for rule in rules), default=0)
    strongest = max(scores.values(), default=0)
    return sorted(code for code, strength in scores.items() if strongest > 0 and strength == strongest)


def calculate(rates, loading, country, route, cbm):
    if cbm is None or cbm <= 0:
        raise ValueError('A térfogatnak pozitívnak kell lennie.')
    bands = [v for k, v in rates.items() if k[:3] == (loading, country, route)]
    if cbm > 70:
        found = [b for b in bands if b['upper'] is None]
    else:
        found = sorted((b for b in bands if b['upper'] is not None and b['upper'] >= cbm),
                       key=lambda b: b['upper'])[:1]
    if len(found) != 1:
        raise ValueError('Nincs egyértelmű díjsáv ehhez az útvonalhoz és térfogathoz.')
    band = found[0]
    if band['price'] is None or band['price'] <= 0:
        raise ValueError('Az utolsó ároszlop megfelelő cellája üres vagy nem pozitív szám.')
    raw = band['price'] if cbm > 70 else cbm * band['price']
    return raw.quantize(Decimal('.01'), rounding=ROUND_HALF_UP), band


ASSIGNMENT_FIELD = 'Odaítélés dátuma (Fuvar szintje)'
DEADLINE_FIELD = 'Kirakodás dátuma (eddig)'


def source_days(groups, field):
    days = set()
    for group in groups:
        for row in group['rows']:
            try:
                value = timestamp(row.get(field))
                if value:
                    days.add(value.date())
            except ValueError:
                pass
    return sorted(days)


def filter_groups(groups, assigned_range=None, deadline_range=None, weeks=None):
    """Filter whole tours by original source dates, preserving all shipment amounts.

    All active conditions must match at least one single source shipment row.
    ISO weeks include their ISO year, and both date endpoints are inclusive.
    Invalid/missing source dates in an active filter are reported explicitly.
    """
    selected, unavailable, partial = [], set(), []
    for group in groups:
        matches = []
        for row in group['rows']:
            valid = True
            for field, bounds, week_set in [(ASSIGNMENT_FIELD, assigned_range, None),
                                             (DEADLINE_FIELD, deadline_range, weeks)]:
                if bounds is None and week_set is None:
                    continue
                try:
                    value = timestamp(row.get(field))
                except ValueError:
                    value = None
                if value is None:
                    unavailable.add(group['id'])
                    valid = False
                    continue
                day = value.date()
                if bounds is not None and not bounds[0] <= day <= bounds[1]:
                    valid = False
                if week_set is not None and tuple(day.isocalendar()[:2]) not in week_set:
                    valid = False
            matches.append(valid)
        if any(matches):
            selected.append(group)
            if not all(matches):
                partial.append(group['id'])
    return selected, sorted(unavailable), partial


def week_label(week):
    start = date.fromisocalendar(week[0], week[1], 1)
    end = start + timedelta(days=6)
    return f'{week[0]} / {week[1]:02d}. hét · {start:%Y.%m.%d.} – {end:%Y.%m.%d.}'


@st.cache_data(show_spinner=False)
def brand_asset64(name):
    return base64.b64encode((Path(__file__).parent / 'assets' / name).read_bytes()).decode()


def render_brand():
    root = Path(__file__).parent
    st.set_page_config(page_title='Hartmann fuvarimport | Bábolna Sped',
                       page_icon=str(root / 'assets' / 'logo_ui.png'), layout='wide')
    fonts = ''.join(f'@font-face{{font-family:{family};src:url(data:font/ttf;base64,{brand_asset64(filename)})}}'
                    for family, filename in [('Fira', 'FiraSans-Regular.ttf'), ('Montserrat', 'Montserrat.ttf')])
    st.markdown('<style>' + fonts + '''
    html,body,.stApp{font-family:Fira,sans-serif;color:#2A3756}
    .stApp{background:#F6F7F9}
    h1,h2,h3{font-family:Montserrat,sans-serif!important;color:#2A3756}
    [data-testid="stSidebar"]{background:#FFF;border-right:1px solid #D1D5DB}
    [data-testid="stSidebar"] [data-testid="stImage"]{background:white;padding:8px;border-radius:8px}
    [data-testid="stSidebar"] p,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] label{color:#2A3756}
    [data-testid="stMetric"]{background:white;padding:22px;border-radius:14px;border:1px solid #D1D5DB}
    [data-testid="stMetricValue"]{font-size:30px;font-weight:700;letter-spacing:-0.5px}
    [data-testid="stMetricLabel"]{font-size:13px;color:#657087}
    [data-testid="stFileUploader"]{background:white;padding:16px;border:1px solid #D1D5DB;border-radius:14px}
    [data-testid="stAlert"],[data-testid="stExpander"]{border-radius:12px}
    div.stButton>button[kind="primary"]{background:#F2E47D;color:#2A3756;border:0;font-weight:700}
    div.stDownloadButton>button{border:1px solid #D1D5DB;border-radius:9px;background:#F2E47D;color:#2A3756;padding:12px 24px}
    .block-container{padding-top:2rem;max-width:1700px}
    .brandbar{background:#2A3756;color:white;padding:16px 24px;border-radius:12px;margin-bottom:22px}
    .brandbar span{color:#F2E47D}
    </style>''', unsafe_allow_html=True)
    with st.sidebar:
        st.image(str(root / 'assets' / 'logo_ui.png'), width=230)
        st.markdown('## Hartmann fuvarimport')
        st.caption('TRANSPOREON · INNOTMS')
        st.divider()
    st.markdown('<div class="brandbar">BÁBOLNA SPED <span>· HARTMANN FUVARIMPORT</span></div>', unsafe_allow_html=True)
    st.title('Hartmann fuvarimport')
    st.caption('Válassz időszakot, ellenőrizd a fuvarokat, majd töltsd le az InnoTMS-importot.')


def create_export(template, records):
    """Preserve the supplied 30 columns, duplicate Weight header and hidden TP-Meta."""
    book = openpyxl.load_workbook(io.BytesIO(template))
    sheet = book['Odaítélt-fuvarok']
    headers = [text(c.value) for c in sheet[1]]
    if len(headers) != 30 or headers[0] != 'Fuvar sz.' or headers[-1] != 'Fuvardíj':
        raise ValueError('Az importsablon nem a várt 30 oszlopos formátum.')
    row_styles = [copy(c._style) for c in sheet[2]]
    sample_height = sheet.row_dimensions[2].height if 2 in sheet.row_dimensions else None
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    for rec in records:
        r, g = rec['source'], rec['group']
        loading_date, unloading_date = rec['loading_date'], rec['unloading_date']
        t0, t1 = timestamp(r.get(DATE_FIELDS[2])), timestamp(r.get(DATE_FIELDS[3]))
        time_range = f'{t0:%H:%M} - {t1:%H:%M}' if t0 and t1 else ''
        def station(prefix):
            return '-'.join([text(r.get(f'Ország - {prefix}')), text(r.get(f'Postai irányítószám - {prefix}'))]) + ' ' + text(r.get(f'Város - {prefix}'))
        values = [g['id'], rec.get('extra', ''), text(r.get('Cég neve - Kirakodási hely')),
                  text(r.get('Ország - Kirakodási hely')), text(r.get('Postai irányítószám - Kirakodási hely')),
                  text(r.get('Város - Kirakodási hely')), text(r.get('Rendsz.')),
                  loading_date.isoformat(), unloading_date.isoformat(), time_range,
                  '', '', '', float(g['sums']['Súly']), float(g['sums']['Súly']), 'kg',
                  float(g['sums']['Mennyiség']), 'cbm', len(g['rows']), '', '',
                  text(r.get('Megjegyzés - Kirakodási hely')), text(r.get('Megbízó')),
                  station('Berakodási hely'), station('Kirakodási hely'),
                  text(r.get('Cég neve - Berakodási hely')), text(r.get('Postai irányítószám - Berakodási hely')),
                  text(r.get('Város - Berakodási hely')), text(r.get('Ország - Berakodási hely')),
                  float(rec['price'])]
        sheet.append(values)
        # Treat source text literally, including strings beginning with '='.
        for cell, value, style in zip(sheet[sheet.max_row], values, row_styles):
            cell._style = copy(style)
            if isinstance(value, str):
                cell.data_type = 's'
        if sample_height is not None:
            sheet.row_dimensions[sheet.max_row].height = sample_height
    out = io.BytesIO()
    book.save(out)
    book.close()
    return out.getvalue()


def remember_table_edits(editor_key, override_key, ids):
    """Keep explicit edits when an upstream route/date choice rebuilds the grid."""
    saved = st.session_state.setdefault(override_key, {})
    for row_index, changes in st.session_state.get(editor_key, {}).get('edited_rows', {}).items():
        index = int(row_index)
        if 0 <= index < len(ids):
            saved.setdefault(ids[index], {}).update(changes)


def main():
    render_brand()
    with st.sidebar:
        st.header('Forrásfájlok')
        source = st.file_uploader('1. Transporeon-export', type=['xlsx'], key='source')
        st.subheader('2. Aktuális fuvardíjak')
        uploaded_rates = None
        with st.expander('Google Sheets-kapcsolat', expanded=True):
            url = st.text_input('Díjtáblázat hivatkozása', value=GOOGLE_URL, key='google_url')
            fetch = st.button('Díjtábla frissítése', key='google_refresh')
            # Attempt once at session startup, on URL change, or explicit refresh.
            # Ordinary edits and filters reuse this session's successful read.
            if st.session_state.get('google_attempted_url') != url or fetch:
                st.session_state['google_attempted_url'] = url
                st.session_state.pop('live_rates', None)
                st.session_state.pop('google_error', None)
                try:
                    with st.spinner('Google-díjtábla betöltése…'):
                        try:
                            credentials = st.secrets.get('google_service_account')
                        except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
                            credentials = None
                        google_values = fetch_google(url, credentials)
                        _, header, column = parse_rates(google_values)
                        st.session_state['live_rates'] = {
                            'values': google_values, 'url': url,
                            'at': datetime.now().isoformat(timespec='seconds'),
                            'price_header': header, 'price_column': column}
                except Exception:
                    st.session_state['google_error'] = 'Nem sikerült a Google-díjtábla beolvasása.'
            live = st.session_state.get('live_rates', {})
            if live.get('url') == url and not st.session_state.get('google_error'):
                st.success('Google-díjtábla betöltve.')
                st.caption('Használt ároszlop: ' + live['price_column'] + ' – ' + live['price_header'])
            else:
                st.warning('Nem sikerült a Google-díjtábla beolvasása. Próbáld újra a frissítést, vagy töltsd fel az XLSX-változatát.')
                uploaded_rates = st.file_uploader('Google-díjtábla XLSX-változata', type=['xlsx'],
                    key='fallback_rates:' + hashlib.sha256(url.encode()).hexdigest()[:12],
                    help='A Hartmann fuvardíjak táblázatból letöltött Excel, Rates munkalappal.')
                st.caption('Zárt táblázathoz olvasási hozzáférésű Google-szolgáltatásfiók állítható be a Streamlit Secretsben.')
        with st.expander('Útvonalak és importsablon'):
            route_upload = st.file_uploader('Új Final External ratesheet (opcionális)', type=['xlsx'])
            st.caption('Alapértelmezés: a megadott 2026-os fájl Routes Overview (Road) adatai.')
            template_upload = st.file_uploader('Másik InnoTMS-importminta (opcionális)', type=['xlsx'])
            st.caption('Az eredeti importminta fejlécei és rejtett TP-Meta lapja be vannak építve.')
    try:
        if uploaded_rates is not None:
            rate_bytes = uploaded_rates.getvalue()
            values = read_rate_xlsx(rate_bytes)
            rate_identity = hashlib.sha256(rate_bytes).hexdigest()
            origin = f'Feltöltött díjtábla: {uploaded_rates.name}'
        elif st.session_state.get('live_rates', {}).get('url') == url:
            live = st.session_state['live_rates']
            values = live['values']
            rate_identity = hashlib.sha256(json.dumps(values, default=str).encode()).hexdigest()
            origin = 'Google Sheets · beolvasva: ' + live['at']
        else:
            values = None
        if not source or values is None:
            if not source:
                st.info('Töltsd fel a Transporeon-exportot a bal oldalon.')
            else:
                st.info('A Google-kapcsolat hibáját javítsd a frissítéssel, vagy töltsd fel a díjtábla XLSX-változatát a bal oldalon.')
            st.markdown('**A feldolgozás menete**\n\n1. Fuvarszám szerinti összevonás.\n2. Dátum- és útvonaleltérések rendezése.\n3. Fuvardíjak ellenőrzése és az importfájl letöltése.')
            return
        source_bytes = source.getvalue()
        route_bytes = route_upload.getvalue() if route_upload else None
        template_bytes = template_upload.getvalue() if template_upload else base64.b64decode(TEMPLATE_B64)
        # Keep uploaded business data in this user's session, not a shared cache.
        context = hashlib.sha256(DATA_SCHEMA.encode() + source_bytes + (route_bytes or b'') + template_bytes + rate_identity.encode()).hexdigest()[:16]
        if st.session_state.get('context') != context:
            st.session_state['context'] = context
            st.session_state['prepared'] = aggregate(source_bytes)
            st.session_state['routes'] = read_routes(route_bytes)
            st.session_state['rates'] = parse_rates(values)
        groups, count, _ = st.session_state['prepared']
        routes = st.session_state['routes']
        rates, price_header, price_column = st.session_state['rates']
    except Exception as exc:
        st.error(f'A fájlok feldolgozása sikertelen: {exc}')
        return
    key = lambda value: f'{context}:{value}'
    all_groups = groups
    award_days = source_days(all_groups, ASSIGNMENT_FIELD)
    deadline_days = source_days(all_groups, DEADLINE_FIELD)
    assigned_range, deadline_range, selected_weeks = None, None, None
    st.subheader('Feltöltési időszak')
    with st.container(border=True):
        award_column, deadline_column = st.columns(2)
        with award_column:
            by_award = st.toggle('Szűrés az odaítélés dátumára', key=key('filter_award'))
            if by_award:
                if not award_days:
                    st.warning('A forrásban nincs értelmezhető odaítélési dátum.')
                    return
                chosen = st.date_input('Odaítélés: mettől meddig', value=(award_days[0], award_days[-1]),
                    min_value=award_days[0], max_value=award_days[-1], format='YYYY.MM.DD', key=key('award_range'))
                if len(chosen) != 2:
                    st.info('Válaszd ki az odaítélési időszak kezdő- és zárónapját.')
                    return
                assigned_range = tuple(chosen)
            else:
                st.caption('Minden odaítélési időpont.')
        with deadline_column:
            delivery_mode = st.radio('Szűrés a kirakodási határidőre',
                ['Nincs szűrés', 'Dátumtartomány', 'Hétválasztás'], horizontal=True, key=key('delivery_filter_mode'))
            if delivery_mode != 'Nincs szűrés':
                if not deadline_days:
                    st.warning('A forrásban nincs értelmezhető kirakodási határidő.')
                    return
                if delivery_mode == 'Dátumtartomány':
                    chosen = st.date_input('Kirakodási határidő: mettől meddig',
                        value=(deadline_days[0], deadline_days[-1]), min_value=deadline_days[0],
                        max_value=deadline_days[-1], format='YYYY.MM.DD', key=key('deadline_range'))
                    if len(chosen) != 2:
                        st.info('Válaszd ki a kirakodási időszak kezdő- és zárónapját.')
                        return
                    deadline_range = tuple(chosen)
                else:
                    available_weeks = sorted({tuple(d.isocalendar()[:2]) for d in deadline_days})
                    chosen = st.multiselect('Teljesítési hetek', options=available_weeks,
                        format_func=week_label, key=key('delivery_weeks'), placeholder='Válassz egy vagy több hetet')
                    if not chosen:
                        st.info('Válassz legalább egy teljesítési hetet.')
                        return
                    selected_weeks = set(chosen)
        st.caption('A két szűrő együtt is használható. A teljesítési hét az eredeti „Kirakodás dátuma (eddig)” mező ISO-hete, hétfőtől vasárnapig. A kezdő- és zárónap is beleszámít.')
    groups, unavailable, partial = filter_groups(all_groups, assigned_range, deadline_range, selected_weeks)
    st.caption(f'A szűrésben {len(groups)} / {len(all_groups)} fuvar szerepel. A szűrés a forrás dátumai alapján történik; a későbbi kézi dátumválasztás ezt nem változtatja meg.')
    if unavailable:
        st.warning('Hiányzó vagy hibás dátum az aktív szűréshez: ' + ', '.join(unavailable))
    if partial:
        st.info('Ezeknél a fuvaroknál csak a szállítmányok egy része esik az időszakba, ezért a teljes fuvart megtartottuk: ' + ', '.join(partial))
    if not groups:
        st.info('A kiválasztott időszakban nincs fuvar. Módosítsd a szűrőket.')
        return
    count = sum(len(g['rows']) for g in groups)
    def route_label(code):
        if ' [' not in code:
            return code
        entry = next(v for k, v in rates.items() if k[2] == code)
        return code + ' · ' + entry['description']
    st.caption(f'{origin} · Használt ároszlop: {price_column} – {price_header} · Pénznem: EUR')
    metrics = st.columns(4)
    metrics[0].metric('Szűrt szállítmánysor', count)
    metrics[1].metric('Fuvar a szűrésben', len(groups))
    metrics[2].metric('Összevont fuvar', sum(len(g['rows']) > 1 for g in groups))
    metrics[3].metric('Eltérő dátumú fuvar', sum(bool(g['date_conflicts']) for g in groups))
    st.caption('A berakodás és kirakodás alapértelmezett napja a forrás megfelelő „ettől” dátuma. A lista dátumai szerkeszthetők. 70 cbm-ig egységár, 70 cbm felett fix díj; kerekítés csak a végösszegnél.')
    decisions = {}
    exceptions = [g for g in groups if g['date_conflicts'] or g['different'] or g['errors']]
    if exceptions:
        with st.expander(f'Ellenőrizendő eltérések · {len(exceptions)} fuvar', expanded=True):
            for g in exceptions:
                st.markdown(f'**{g["id"]}** · {text(g["rows"][0].get("Cég neve - Kirakodási hely"))}')
                for error in g['errors']:
                    st.error(error)
                representative = 0
                if g['different']:
                    st.warning('Eltérő mezők: ' + ', '.join(g['different']))
                    st.dataframe(pd.DataFrame([{**{n: text(r.get(n)) for n in IDENTIFIERS},
                                               **{n: text(r.get(n)) for n in g['different']}} for r in g['rows']]), hide_index=True)
                    representative = st.selectbox('Melyik szállítmány törzsadatait vegyük át?',
                        range(len(g['rows'])), index=None,
                        format_func=lambda i, g=g: text(g['rows'][i]['Szállítmánysz.']), key=key('row:'+g['id']))
                decisions[g['id']] = {'representative': representative}
                columns = st.columns(2)
                for column, label in zip(columns, ['Berakodás', 'Kirakodás']):
                    if label in g['date_conflicts']:
                        column.warning(f'Eltérő {label.lower()}i dátumok')
                        chosen = column.selectbox(label + ' napja', g['dates'][label], index=None,
                             format_func=lambda d: d.strftime('%Y.%m.%d.'), key=key(label+g['id']))
                        decisions[g['id']][label] = chosen
    # One explicit route decision is shared by identical loading-country/destination ZIPs.
    destination_info = {}
    for g in groups:
        chosen = decisions.get(g['id'], {}).get('representative', 0)
        r = g['rows'][chosen if chosen is not None else 0]
        loading = text(r.get('Ország - Berakodási hely')).upper()
        country = text(r.get('Ország - Kirakodási hely')).upper()
        postal = normalize_zip(r.get('Postai irányítószám - Kirakodási hely'), country)
        destination = (loading, country, postal)
        if destination not in destination_info:
            candidates = candidate_routes(routes, rates, *destination)
            destination_info[destination] = {'candidates': candidates, 'city': text(r.get('Város - Kirakodási hely'))}
    route_decisions = {}
    unresolved = {k: v for k, v in destination_info.items() if len(v['candidates']) != 1}
    if unresolved:
        with st.expander(f'Útvonalválasztás szükséges · {len(unresolved)} célterület', expanded=True):
            st.caption('Átfedő vagy hiányzó irányítószám-szabály esetén válassz útvonalat. Az azonos célterület fuvarjaira együtt alkalmazzuk.')
            for destination, info in unresolved.items():
                loading, country, postal = destination
                available = sorted({k[2] for k in rates if k[0] == loading and k[1] == country})
                label = f'{loading} → {country} {postal} {info["city"]}'
                if info['candidates']:
                    st.caption(label + ' · Találatok: ' + ', '.join(info['candidates']))
                else:
                    st.caption(label + ' · Nincs illeszkedő irányítószám-szabály.')
                if available:
                    route_decisions[destination] = st.selectbox(label, available, index=None, format_func=route_label,
                                                                key=key('route:'+'|'.join(destination)))
                else:
                    st.warning(label + ': nincs díjazott útvonal. A fuvarnál adj meg kézi díjat indoklással, vagy vedd ki az importból.')
    rows, records = [], {}
    for g in groups:
        decision = decisions.get(g['id'], {})
        selected = decision.get('representative', 0)
        r = g['rows'][selected if selected is not None else 0]
        errors = list(g['errors'])
        if selected is None:
            errors.append('Válassz átveendő törzsadatot.')
        selected_dates = {}
        for label, field in [('Berakodás', DATE_FIELDS[0]), ('Kirakodás', DATE_FIELDS[2])]:
            if label in g['date_conflicts']:
                d = decision.get(label)
                if d is None:
                    errors.append(label + ': válassz dátumot az eltéréseknél.')
            else:
                try:
                    ts = timestamp(r.get(field))
                    d = ts.date() if ts else None
                except ValueError:
                    d = None
            selected_dates[label] = d
        loading = text(r.get('Ország - Berakodási hely')).upper()
        country = text(r.get('Ország - Kirakodási hely')).upper()
        postal = normalize_zip(r.get('Postai irányítószám - Kirakodási hely'), country)
        destination = (loading, country, postal)
        candidates = destination_info[destination]['candidates']
        route = candidates[0] if len(candidates) == 1 else route_decisions.get(destination)
        price, band, price_errors = None, None, []
        if not route:
            price_errors.append('Válassz útvonalkódot, vagy adj meg indokolt kézi díjat.')
        else:
            try:
                price, band = calculate(rates, loading, country, route, g['sums']['Mennyiség'])
            except ValueError as exc:
                price_errors.append(str(exc))
        records[g['id']] = {'source': r, 'group': g, 'price': price, 'route': route,
                            'errors': errors, 'price_errors': price_errors, 'band': band}
        rows.append({'Import': True, 'Fuvar sz.': g['id'], 'Címzett': text(r.get('Cég neve - Kirakodási hely')),
                     'Ország': country, 'Irányítószám': postal, 'Város': text(r.get('Város - Kirakodási hely')),
                     'Berakodás': selected_dates['Berakodás'], 'Kirakodás': selected_dates['Kirakodás'],
                     'Súly (kg)': float(g['sums']['Súly']) if g['sums']['Súly'] is not None else None,
                     'Térfogat (cbm)': float(g['sums']['Mennyiség']) if g['sums']['Mennyiség'] is not None else None,
                     'Szállítmányok': len(g['rows']), 'Útvonal': route or '',
                     'Díjsáv': band['band'] if band else '', 'Fuvardíj (EUR)': float(price) if price is not None else None,
                     'Kézi fuvardíj (EUR)': None, 'Kézi díj indoka': '',
                     'Megbízó további számai': '', 'Jelzés': ' · '.join(errors + price_errors)})
    st.subheader('Fuvarlista ellenőrzése')
    st.caption('Az Import jelöléssel választhatod ki a feltöltendő fuvarokat. A dátumok és a „Megbízó további számai” mező szerkeszthetők. Az utóbbihoz nincs azonosított forrásmező, ezért kezdetben üres. Kézi fuvardíj csak indoklással adható meg; ilyenkor ez kerül az exportba a számított díj helyett.')
    frame = pd.DataFrame(rows)
    for c in ['Berakodás', 'Kirakodás']:
        frame[c] = pd.to_datetime(frame[c])
    frame['Kézi fuvardíj (EUR)'] = pd.Series([None] * len(frame), dtype='Float64')
    editor_key = key('editor:'+hashlib.sha256(frame.to_json(date_format='iso').encode()).hexdigest()[:12])
    override_key, initial_key = key('table_overrides'), key('editor_initial')
    if st.session_state.get(initial_key, {}).get('key') != editor_key:
        for i, row in frame.iterrows():
            for column, value in st.session_state.get(override_key, {}).get(row['Fuvar sz.'], {}).items():
                if column in ['Berakodás', 'Kirakodás']:
                    value = pd.to_datetime(value) if value is not None else pd.NaT
                frame.at[i, column] = value
        st.session_state[initial_key] = {'key': editor_key, 'frame': frame.copy()}
    frame = st.session_state[initial_key]['frame'].copy()
    edited = st.data_editor(frame, hide_index=True, width='stretch', height=500,
        disabled=[c for c in frame.columns if c not in ['Import', 'Berakodás', 'Kirakodás', 'Megbízó további számai', 'Kézi fuvardíj (EUR)', 'Kézi díj indoka']],
        column_config={'Import': st.column_config.CheckboxColumn('Import'),
                       'Berakodás': st.column_config.DateColumn(format='YYYY.MM.DD'),
                       'Kirakodás': st.column_config.DateColumn(format='YYYY.MM.DD'),
                       'Súly (kg)': st.column_config.NumberColumn(format='%.3f'),
                       'Térfogat (cbm)': st.column_config.NumberColumn(format='%.3f'),
                       'Kézi fuvardíj (EUR)': st.column_config.NumberColumn(format='%.2f', min_value=0.01, step=0.01),
                       'Fuvardíj (EUR)': st.column_config.NumberColumn(format='%.2f')},
        key=editor_key, on_change=remember_table_edits,
        args=(editor_key, override_key, list(frame['Fuvar sz.'])))
    approved, problems = [], []
    for row in edited.to_dict('records'):
        if not row['Import']:
            continue
        rec = dict(records[row['Fuvar sz.']])
        errors = list(rec['errors'])
        manual = row['Kézi fuvardíj (EUR)']
        if manual is not None and not pd.isna(manual):
            if not text(row['Kézi díj indoka']):
                errors.append('A kézi díjhoz indoklás szükséges.')
            try:
                manual_price = number(manual)
                if manual_price <= 0:
                    raise ValueError('A kézi díjnak pozitívnak kell lennie.')
                rec['price'] = manual_price.quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                rec['manual_reason'] = text(row['Kézi díj indoka'])
            except ValueError as exc:
                errors.append(str(exc))
        else:
            errors.extend(rec['price_errors'])
        for col, output in [('Berakodás', 'loading_date'), ('Kirakodás', 'unloading_date')]:
            value = row[col]
            rec[output] = None if pd.isna(value) else pd.Timestamp(value).date()
            if rec[output] is None:
                errors.append(col + ': hiányzó dátum.')
        if rec['loading_date'] and rec['unloading_date'] and rec['unloading_date'] < rec['loading_date']:
            errors.append('A kirakodás a berakodás előtt van.')
        for name in ['Megbízó', 'Cég neve - Berakodási hely', 'Ország - Berakodási hely',
                     'Postai irányítószám - Berakodási hely', 'Város - Berakodási hely',
                     'Cég neve - Kirakodási hely', 'Ország - Kirakodási hely',
                     'Postai irányítószám - Kirakodási hely', 'Város - Kirakodási hely']:
            if not text(rec['source'].get(name)):
                errors.append('Hiányzó adat: ' + name)
        if errors:
            problems.append(row['Fuvar sz.'] + ': ' + ' · '.join(dict.fromkeys(errors)))
        else:
            rec['extra'] = text(row['Megbízó további számai'])
            approved.append(rec)
    left, right = st.columns(2)
    left.metric('Kijelölt, hibamentes fuvar', len(approved))
    total = sum((r['price'] for r in approved), Decimal(0))
    right.metric('Fuvardíj összesen', f'{total:,.2f} EUR'.replace(',', ' '))
    manual_count = sum('manual_reason' in r for r in approved)
    if manual_count:
        st.warning(f'{manual_count} fuvar kézzel megadott fuvardíjjal szerepel az összegben és az exportban.')
    if problems:
        st.error(f'{len(problems)} kijelölt fuvar ellenőrzést igényel. Rendezd az eltéréseket, vagy vedd ki ezeket az importból.')
        with st.expander('Exportot akadályozó tételek', expanded=True):
            for problem in problems:
                st.write(problem)
    elif approved:
        st.success('A kijelölt fuvarok exportálhatók.')
        try:
            output = create_export(template_bytes, approved)
            st.download_button('InnoTMS-import letöltése', output,
                file_name=f'Hartmann_InnoTMS_{date.today():%Y%m%d}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary', on_click='ignore')
        except Exception as exc:
            st.error(f'Az importfájl elkészítése sikertelen: {exc}')
    else:
        st.info('Jelölj ki legalább egy fuvart.')
    with st.expander('Számítás és forrásadatok ellenőrzése'):
        tour = st.selectbox('Fuvarszám', [g['id'] for g in groups], key=key('audit'))
        rec = records[tour]
        st.dataframe(pd.DataFrame(rec['group']['rows']), hide_index=True)
        if rec['band']:
            b = rec['band']
            volume = rec['group']['sums']['Mennyiség']
            formula = f'{volume} cbm × {b["price"]} EUR/cbm' if volume <= 70 else f'Fix díj: {b["price"]} EUR'
            st.write(f'{rec["route"]} · Rates!{price_column}{b["row"]} · {formula} = {rec["price"]} EUR')
    st.caption('Az alkalmazás letölthető fájlt készít; az InnoTMS-be történő feltöltést te végzed. Az ismeretlen státuszmezők és a módosítás időpontja üresen maradnak.')


if __name__ == '__main__':
    main()
