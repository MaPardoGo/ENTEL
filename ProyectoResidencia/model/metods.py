"""Librerias utilizadas"""
import os
import requests
from dotenv import load_dotenv
from datetime import datetime
import pandas as pd
import requests.auth

# Cargar las variables de entorno desde archivo .env
load_dotenv()

# Configuración de la instancia y credenciales
INSTANCE = os.getenv('INSTANCE')
USER = os.getenv('USER')
PASS = os.getenv('PASS')
URL = f'https://{INSTANCE}.service-now.com/api/now/table'

ubicaciones_dict = {
    "5c36b3fa47f07d50d9d3a03a536d43fa" :    "Bodega Central",
    "59d48d98eba896504efefc6ccad0cd40" :    "Servicio de Evaluaciones Laborales",
    "11dae7de476c3110d9d3a03a536d4370" :    "Teletrabajo",
    "0141dfda47e83110d9d3a03a536d4342" :    "Agencia Antofagasta",
    "037b1b5a472c3110d9d3a03a536d43a4" :	"Agencia Caldera",
    "037b1b5a472c3110d9d3a03a536d43a9" :	"Agencia Coquimbo",
    "077b1b5a472c3110d9d3a03a536d43a5" :	"Agencia Copiapo",
    "0b7b1b5a472c3110d9d3a03a536d43a6" :	"Agencia Los Loros",
    "0b7b1b5a472c3110d9d3a03a536d43e8" :	"Agencia Puerto Natales",
    "0f7b1b5a472c3110d9d3a03a536d43a2" :	"Agencia Calama",
    "0f7b1b5a472c3110d9d3a03a536d43a7" :	"Agencia Coyhaique",
    "137b1b5a472c3110d9d3a03a536d43eb" :	"Agencia San Felipe",
    "1f7b1b5a472c3110d9d3a03a536d43e9" :	"Agencia Chillan",
    "202daf7347ac3910d9d3a03a536d4329" :	"Hospital Del Trabajador",
    "25036756476c3110d9d3a03a536d4308" :	"Agencia San Antonio",
    "437b1b5a472c3110d9d3a03a536d43d2" :	"Agencia Victoria",
    "437b1b5a472c3110d9d3a03a536d43d7" :	"Agencia Castro",
    "477b1b5a472c3110d9d3a03a536d43d3" :	"Agencia Angol",
    "477b1b5a472c3110d9d3a03a536d43d8" :	"Agencia Calbuco",
    "4b2bf65647683110d9d3a03a536d43da" :	"Agencia La Calera",
    "4b7b1b5a472c3110d9d3a03a536d43cf" :	"Agencia La Serena",
    "4b7b1b5a472c3110d9d3a03a536d43d4" :	"Agencia Laja",
    "4b7b1b5a472c3110d9d3a03a536d43d9" :	"Agencia Osorno",
    "4f7b1b5a472c3110d9d3a03a536d43d0" :	"Agencia Vicuña",
    "4f7b1b5a472c3110d9d3a03a536d43d5" :	"Agencia Nacimiento",
    "4f7b1b5a472c3110d9d3a03a536d43da" :	"Agencia La Union",
    "61036756476c3110d9d3a03a536d430f" :	"Agencia Talcahuano",
    "642e769a47683110d9d3a03a536d43c9" :	"Agencia Los Andes",
    "65036756476c3110d9d3a03a536d430b" :	"Agencia Viña Del Mar",
    "65036756476c3110d9d3a03a536d4310" :	"Agencia Rengo",
    "69036756476c3110d9d3a03a536d430c" :	"Agencia Cañete",
    "69036756476c3110d9d3a03a536d4311" :	"Agencia San Fernando",
    "6d036756476c3110d9d3a03a536d430d" :	"Agencia Coronel",
    "71036756476c3110d9d3a03a536d4314" :	"Agencia Parral",
    "71036756476c3110d9d3a03a536d4319" :	"Agencia Maipu",
    "71036756476c3110d9d3a03a536d431e" :	"Agencia Parque Las Americas",
    "71036756476c3110d9d3a03a536d4323" :	"Agencia Paine",
    "71036756476c3110d9d3a03a536d4328" :	"Agencia Esachs Bustamente",
    "75036756476c3110d9d3a03a536d4315" :	"Agencia Curico",
    "75036756476c3110d9d3a03a536d431a" :	"Agencia Puente Alto",
    "75036756476c3110d9d3a03a536d431f" :	"Agencia Melipilla",
    "79036756476c3110d9d3a03a536d4316" :	"Agencia Linares",
    "79036756476c3110d9d3a03a536d431b" :	"Agencia Las Condes",
    "79036756476c3110d9d3a03a536d4320" :	"Agencia Talagante",
    "79036756476c3110d9d3a03a536d4325" :	"Centro Medico",
    "7d036756476c3110d9d3a03a536d4312" :	"Agencia Santa Cruz",
    "7d036756476c3110d9d3a03a536d4317" :	"Agencia Talca",
    "7d036756476c3110d9d3a03a536d431c" :	"Agencia Quilicura",
    "7d036756476c3110d9d3a03a536d4321" :	"Agencia Santiago",
    "7d036756476c3110d9d3a03a536d4326" :	"Otec Servicios",
    "7e39f0591bddb910d5166392b24bcb87" :	"Agencia Rancagua 333",
    "802dba1a47683110d9d3a03a536d43b0" :	"Agencia La Ligua",
    "837b1b5a472c3110d9d3a03a536d43a7" :	"Agencia Vallenar",
    "877b1b5a472c3110d9d3a03a536d43a0" :	"Agencia Tocopilla",
    "877b1b5a472c3110d9d3a03a536d43a3" :	"Agencia Arica",
    "877b1b5a472c3110d9d3a03a536d43a8" :	"Agencia Puerto Aysen",
    "8b7b1b5a472c3110d9d3a03a536d43a4" :	"Agencia Chañaral",
    "8b7b1b5a472c3110d9d3a03a536d43a9" :	"Agencia Illapel",
    "8f7b1b5a472c3110d9d3a03a536d43a5" :	"Agencia El Salvador",
    "937b1b5a472c3110d9d3a03a536d43e9" :	"Agencia Punta Arenas",
    "95a8df56472c3110d9d3a03a536d43dd" :	"Agencia Mejillones",
    "977b1b5a472c3110d9d3a03a536d43ea" :	"Agencia Iquique",
    "c37b1b5a472c3110d9d3a03a536d43d0" :	"Agencia Ovalle",
    "c37b1b5a472c3110d9d3a03a536d43d5" :	"Agencia Los Angeles",
    "c37b1b5a472c3110d9d3a03a536d43da" :	"Agencia Rio Bueno",
    "c77b1b5a472c3110d9d3a03a536d43d1" :	"Agencia Temuco",
    "c77b1b5a472c3110d9d3a03a536d43d6" :	"Agencia Ancud",
    "c77b1b5a472c3110d9d3a03a536d43db" :	"Agencia Valdivia",
    "cb7b1b5a472c3110d9d3a03a536d43d2" :	"Agencia Villarrica",
    "cb7b1b5a472c3110d9d3a03a536d43d7" :	"Agencia Quellon",
    "cf7b1b5a472c3110d9d3a03a536d43d3" :	"Agencia Cabrero",
    "cf7b1b5a472c3110d9d3a03a536d43d8" :	"Agencia Puerto Montt",
    "e04d9a0f1be53510d5166392b24bcb78" :	"Centro De Atencion Ambulatoria",
    "e1036756476c3110d9d3a03a536d430d" :	"Agencia Concepcion",
    "e1036756476c3110d9d3a03a536d4312" :	"Agencia San Vicente",
    "e5036756476c3110d9d3a03a536d430e" :	"Agencia Curanilahue",
    "e9036756476c3110d9d3a03a536d430f" :	"Agencia Rancagua",
    "ed036756476c3110d9d3a03a536d430b" :	"Agencia Arauco",
    "ed036756476c3110d9d3a03a536d4310" :	"Agencia La Rosa Poli",
    "f1036756476c3110d9d3a03a536d4317" :	"Agencia San Javier",
    "f1036756476c3110d9d3a03a536d431c" :	"Agencia Colina",
    "f1036756476c3110d9d3a03a536d4321" :	"Agencia San Miguel",
    "f1036756476c3110d9d3a03a536d4326" :	"Agencia Providencia",
    "f3377b3e47f07d50d9d3a03a536d43d7" :	"Sala Cuna",
    "f44a63de476c3110d9d3a03a536d43c2" :	"Agencia Valparaiso",
    "f5036756476c3110d9d3a03a536d4313" :	"Agencia Cauquenes",
    "f5036756476c3110d9d3a03a536d4318" :	"Agencia Alameda",
    "f5036756476c3110d9d3a03a536d431d" :	"Agencia Vespucio Oeste",
    "f5036756476c3110d9d3a03a536d4322" :	"Agencia Buin",
    "f5036756476c3110d9d3a03a536d4327" :	"Esachs Servicios",
    "f8dc277347ac3910d9d3a03a536d43e6" :	"Casa Central",
    "f9036756476c3110d9d3a03a536d4314" :	"Agencia Constitucion",
    "f9036756476c3110d9d3a03a536d4319" :	"Agencia La Florida",
    "f9036756476c3110d9d3a03a536d4323" :	"Agencia San Bernardo",
    "fd036756476c3110d9d3a03a536d4315" :	"Agencia Hualañe",
    "fd036756476c3110d9d3a03a536d431a" :	"Agencia La Reina",
    "fd036756476c3110d9d3a03a536d431f" :	"Agencia Peñaflor",
    "786b4bd547b6b590d9d3a03a536d43d8" :	"Agencia Servicio de Examenes Laborales",

}

def registers(table,params):
    url = f"{URL}/{table}"
    try:
        response = requests.get(url, auth=requests.auth.HTTPBasicAuth(USER, PASS)
                                , params=params, verify=False, timeout=20)
        if response.status_code == 200:
            records = response.json().get('result', [])
            for record in records:
                location_id = record.get('location')
                if isinstance(location_id, dict):  # Verifica si location_id es un diccionario
                    location_id = location_id.get('value')  # Extrae el valor
                if location_id in ubicaciones_dict:
                    record['location'] = ubicaciones_dict[location_id]
            print(f'Registros obtenidos de {table}: {len(records)}')
            return records
        else:
            print(f"{response.status_code} al obtener registros de {table}: {response.text}")
            return []
    except requests.exceptions.Timeout:
        print(f"Timeout alcanzado al intentar obtener ubicaciones de {table}")
        return []
    
def registers_wfm(table, params):
    url = f"{URL}/{table}"
    try:
        response = requests.get(url, auth=requests.auth.HTTPBasicAuth(USER, PASS)
                                , params=params, verify=False, timeout=20)
        if response.status_code == 200:
            records = response.json().get('result', [])
            print(f'Registros obtenidos de {table}: {len(records)}')

            # Modifica el valor de 'task_for' con <value> en cada registro
            for record in records:
                if 'parent' in record and 'value' in record['parent']:
                    record['parent'] = record['parent']['value']
                else:
                    record['parent'] = None  # O cualquier valor por defecto que prefieras
            return records
        else:
            print(f"{response.status_code} al obtener registros: {response.text}")
            return []
    except requests.exceptions.Timeout:
        print("Timeout alcanzado al intentar obtener registros")
        return []




# Guardar lista en planilla.
def save_to_excel(dataframes, output_file='report/Cons_pend_Residencia.xlsx'):
    """ [Crea y guarda registros en excel] """

    # Filtra solo los dataframes no vacíos
    non_empty_dataframes = {sheet_name: df for sheet_name, df in dataframes.items() if not df.empty}
    # Si no hay DataFrames para guardar, salir
    if not non_empty_dataframes:
        print("No hay datos para guardar.")
        return

    # Guardar los dataframes no vacíos en el archivo Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, dataframe in non_empty_dataframes.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f'Hoja {sheet_name} guardada con éxito.')
    print(f"Archivo guardado en {output_file}")

def categorizarPorDias(df):
    df['sys_created_on'] = pd.to_datetime(df['sys_created_on'], errors='coerce')

    #obtener fecha actual
    current_date = datetime.now()

    #calcular diferencia en días
    df['dias_diferencia'] = (current_date - df['sys_created_on']).dt.days

    df['categoria_dias'] = pd.cut(df['dias_diferencia'],
                                  bins=[-1,0,1,3,5,15,30],
                                  labels=['0 día', '1 día', '2 a 3 días', '3 a 5 días', '5 a 15 días', '15 a 30 días'],
                                  include_lowest=True)
    return df

def save_to_excel_global(dataframes, output_file='report/Cons_pend_Global.xlsx'):
    """ [Crea y guarda registros en excel] """

    # Filtra solo los dataframes no vacíos
    non_empty_dataframes = {sheet_name: df for sheet_name, df in dataframes.items() if not df.empty}
    # Si no hay DataFrames para guardar, salir
    if not non_empty_dataframes:
        print("No hay datos para guardar.")
        return

    # Guardar los dataframes no vacíos en el archivo Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, dataframe in non_empty_dataframes.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f'Hoja {sheet_name} guardada con éxito.')
    print(f"Archivo guardado en {output_file}")