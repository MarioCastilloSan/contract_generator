import pandas as pd
from word_gen import example_contract
import os
from zipfile import ZipFile



def readData()->pd.DataFrame:
    df = pd.read_excel('Datos_Aleatorios.xlsx')
    return df 

def role_salary(df: pd.DataFrame, rol:str)->str:
    if rol == 'Asistente':
        salary = '1.000.000'
    elif rol == 'Contador':
        salary = '1.500.000'
    elif rol == 'Desarrollador':
        salary = '2.000.000'
    elif rol == 'Supervisor':
        salary = '2.500.000'
    return salary

def singular_data_to_contract(df: pd.DataFrame, index_row:int)->str:
    sub_df = df.iloc[index_row]
    date = sub_df['Fecha']
    rol = sub_df['Rol']
    address = sub_df['Residencia']
    rut = sub_df['RUT']
    full_name = sub_df['nombre_completo']
    nationality = sub_df['Nacionalidad']
    birth_date = sub_df['Fecha de nacimiento']
    profession = sub_df['Profesion']
    salary=role_salary(df, rol)
    end_path=example_contract(date, rol, address, rut, full_name, nationality, birth_date, profession, salary)
    return end_path
    



def multiple_data_to_contract(df: pd.DataFrame, start:int , end:int)->str:
    sub_df = df.iloc[start:end]
    end_paths=[]
    for  index,row in sub_df.iterrows():
        date = row['Fecha']
        rol = row['Rol']
        address = row['Residencia']
        rut = row['RUT']
        full_name = row['nombre_completo']
        nationality = row['Nacionalidad']
        birth_date = row['Fecha de nacimiento']
        profession = row['Profesion']
        salary=role_salary(sub_df, rol)
        end_path=example_contract(date, rol, address, rut, full_name, nationality, birth_date, profession, salary)
        end_paths.append(end_path)
    return end_paths

def singular_path_clean(end_path:str):
    os.remove(end_path)

def multiple_path_clean(end_paths):
    for path in end_paths:
        os.remove(path)

def create_zip_file(paths)->str:
    zip_path = 'contratos.zip'
    with ZipFile(zip_path, 'w') as zipf:
        for path in paths:
            zipf.write(path, os.path.basename(path))
    return zip_path