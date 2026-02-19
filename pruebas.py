from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import pandas as pd
import os
import psycopg2
from collections import defaultdict
from psycopg2 import errors

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ruta = os.path.join(BASE_DIR, 'datos', 'Inventario_(19).xlsx')
ruta_envio = os.path.join(BASE_DIR,'datos','ENVIO(7).xlsx')
ruta_re = os.path.join(BASE_DIR,'datos','reasignacion.xlsx')

doc_inventario = pd.read_excel(ruta, sheet_name='Inventario')
doc_aliados = pd.read_excel(ruta_envio, sheet_name="DESPACHO ALIADOS 2026")
doc_reasignado = pd.read_excel(ruta_re, sheet_name="Hoja1")

doc_aliados['NºSerieFab'] = doc_aliados['NºSerieFab'].astype(str).str.strip()
doc_inventario['Serial1'] = doc_inventario['Serial1'].astype(str).str.strip()

OTP = []
OTH = []
ESTADO_OTP_CRM = []
ESTADO_OTH_CRM = []
Centro = []
ALMACEN = []
Aliado = []
Nombre_Almacén = []
CLIENTE = []
Tipo_de_OT = []
Asignado = []
Codigo_Sap = []
Descripción = []
CANTIDAD = []
SERIAL = []
lote = []
Estado = []
Estatus = []
Fecha_cambio_de_estatus = []
Fecha_de_reporte = []
Fecha_de_ingreso = []
Días_en_Reserva = []

numero = 0

doc_aliados["Fecha"] = pd.to_datetime(
    doc_aliados["Fecha"],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

doc_inventario['FechaOrden'] = pd.to_datetime(
    doc_inventario['FechaOrden'],
    errors="coerce"
).dt.strftime("%d/%m/%Y")


for i in range(len(doc_inventario['Serial1'])):
    serial = doc_inventario['Serial1'].iloc[i]

    if serial in doc_aliados['NºSerieFab'].values:
        
        for x in range(len(doc_aliados['NºSerieFab'])):
            if serial == doc_aliados['NºSerieFab'].iloc[x]:
                OTP.append(doc_aliados['OTP'].iloc[x])
                OTH.append(doc_aliados['OTH'].iloc[x])
                Centro.append(doc_aliados['COD CENTRO'].iloc[x])
                ALMACEN.append(doc_aliados['COD ALM'].iloc[x])
                Aliado.append(doc_aliados['Destino'].iloc[x])
                CLIENTE.append(doc_aliados['CLIENTE'].iloc[x])
                valor_prc = str(doc_aliados['PRC/SOLPED'].iloc[x]).strip()
                if valor_prc.upper().startswith("PRC"):
                    Tipo_de_OT.append(valor_prc)
                else:
                    Tipo_de_OT.append("INSTALACIONES")
                Codigo_Sap.append(doc_aliados['Material'].iloc[x])
                Descripción.append(doc_aliados['Texto breve de material'].iloc[x])
                CANTIDAD.append('1')
                SERIAL.append(serial)
                lote.append(doc_aliados['LOTE'].iloc[x])
                Estado.append('FUNCIONAL')
                Estatus.append('RESERVADO')
                Fecha_de_ingreso.append(doc_aliados['Fecha'].iloc[x])
                Fecha_cambio_de_estatus.append(doc_aliados['Fecha'].iloc[x])


                numero +=1
                break   # 🔥 importante para que no duplique
                
    else:
        OTP.append('N/A')
        OTH.append('N/A')
        Centro.append('C903')
        ALMACEN.append('A500')
        Aliado.append('ALGARTECH')
        CLIENTE.append('N/A') 
        Tipo_de_OT.append('BASE')
        Codigo_Sap.append(doc_inventario['NumeroParte'].iloc[i])
        Descripción.append(doc_inventario['NombreParte'].iloc[i])
        CANTIDAD.append(doc_inventario['CantidadDisponible'].iloc[i])
        SERIAL.append(serial)
        lote.append('VALORADO')
        Estado.append('FUNCIONAL')
        Estatus.append('STOCK')
        Fecha_de_ingreso.append(doc_inventario['FechaOrden'].iloc[i])
        Fecha_cambio_de_estatus.append(doc_inventario['FechaOrden'].iloc[i])

    ESTADO_OTP_CRM.append('N/A')
    ESTADO_OTH_CRM.append('N/A')
    Asignado.append('N/A')
    Fecha_de_reporte.append('17/02/2026')

resultado = pd.DataFrame({
    "OTP": OTP,
    "OTH": OTH,
    "ESTADO_OTP_CRM": ESTADO_OTP_CRM,
    "ESTADO_OTH_CRM": ESTADO_OTH_CRM,
    "Centro": Centro,
    "ALMACEN": ALMACEN,
    "Aliado": Aliado,
    "CLIENTE": CLIENTE,
    "Tipo_de_OT": Tipo_de_OT,
    "Asignado": Asignado,
    "Codigo_Sap": Codigo_Sap,
    "Descripción": Descripción,
    "CANTIDAD": CANTIDAD,
    "SERIAL": SERIAL,
    "Lote": lote,
    "Estado": Estado,
    "Estatus": Estatus,
    "Fecha cambio de estatus": Fecha_cambio_de_estatus,
    "Fecha de reporte": Fecha_de_reporte,
    "Fecha de ingreso": Fecha_de_ingreso
})

doc_reasignado['FECHA_CAMBIO'] = pd.to_datetime(
    doc_reasignado['FECHA_CAMBIO'],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

for i in range (len(doc_reasignado)):
    for x in range(len(resultado)):
        if doc_reasignado["SERIAL"].iloc[i] in resultado["SERIAL"].iloc[x]:
            resultado.loc[x, "OTP"] = doc_reasignado["OTP NUEVA"].iloc[i]
            resultado.loc[x, "CLIENTE"] = doc_reasignado["NUEVO_CLIENTE"].iloc[i]
            resultado.loc[x, "Fecha cambio de estatus"] = doc_reasignado["FECHA_CAMBIO"].iloc[i]
            resultado.loc[x, "Tipo_de_OT"] = "INSTALACIONES"
            resultado.loc[x, "Estatus"] = "REASIGNADO"
            resultado.loc[x,"ALMACEN"] = "Q500"



resultado.to_excel("resultado_final.xlsx", index=False)

