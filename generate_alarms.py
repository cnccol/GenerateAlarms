import pandas as pd
import numpy as np
import datetime

df = pd.read_excel("BASE DE CONTRATADOS 2019 28-10-2019.xlsx")

def get_last(row):
    
    data = [row[x] for x in df.columns if x.startswith("FINALIZACIÓN DE CONTRATO")]
    data = data[::-1]
    
    for x in data:
        if x != np.nan and not pd.isnull(x):
            return x

df['FINALIZACION_ULTIMO_CONTRATO'] = df.apply(get_last, axis = 1)
df['FINALIZACION_ULTIMO_CONTRATO'] = pd.to_datetime(df['FINALIZACION_ULTIMO_CONTRATO'], errors='coerce')


good_data = df[df['FINALIZACION_ULTIMO_CONTRATO'].notnull()]

bad_data = df[df['FINALIZACION_ULTIMO_CONTRATO'].isnull()]
bad_data = bad_data[['#','PRIMER APELLIDO','SEGUNDO APELLIDO','PRIMER NOMBRE','SEGUNDO NOMBRE','NÚMERO DE DOCUMENTO']]


today = pd.Timestamp(datetime.date.today())
date_before = datetime.date.today() + datetime.timedelta(days=7)
date_before = pd.Timestamp(date_before)

to_expire = good_data[good_data['FINALIZACION_ULTIMO_CONTRATO'].between(today,date_before)]
to_expire = to_expire[['#','PRIMER APELLIDO','SEGUNDO APELLIDO','PRIMER NOMBRE','SEGUNDO NOMBRE','NÚMERO DE DOCUMENTO','FINALIZACION_ULTIMO_CONTRATO']]


bad_data.to_excel("errors.xlsx", index = False)
to_expire.to_excel("expiring_contracts.xlsx", index = False)
