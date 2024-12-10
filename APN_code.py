import pandas as pd
import numpy as np

# Función para aplicar la lógica de Excel en Python
def process_parcel(parcel):
    if isinstance(parcel, str) and "-" in parcel:
        # Si contiene un guion, devolver como está
        return parcel
    elif not isinstance(parcel, (int, float)) or pd.isna(parcel) or parcel == 0:
        # Si no es número, es NaN o es 0, devolver como está
        return parcel
    else:
        # Convertir a texto con ceros iniciales asegurando 12 caracteres
        parcel_str = f"{int(parcel):012d}"
        # Formatear con guiones
        return f"{parcel_str[:3]}-{parcel_str[3:6]}-{parcel_str[6:9]}-{parcel_str[9:]}"

# Leer el archivo de Excel
input_file = "RENTAL PERMIT DATA FILE.xlsx"  # Cambia esto por el nombre de tu archivo
output_file = "archivo_procesado.xlsx"

# Cargar los datos
try:
    df = pd.read_excel(input_file, engine='openpyxl')
except FileNotFoundError:
    print(f"El archivo {input_file} no se encontró.")
    exit()

# Verificar que la columna "PARCEL #" exista
if "PARCEL #" not in df.columns:
    print("La columna 'PARCEL #' no existe en el archivo.")
    exit()

# Aplicar la función a la columna "PARCEL #"
df["PARCEL #"] = df["PARCEL #"].apply(process_parcel)

# Guardar el resultado en un nuevo archivo de Excel
df.to_excel(output_file, index=False, engine='openpyxl')
print(f"Archivo procesado guardado como {output_file}.")
