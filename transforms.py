import pandas as pd
from pandas import ExcelWriter
import numpy as np
import openpyxl

## Variables Auxiliares
meses = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
  "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

## Tabla 1
# crear nueva tabla con columnas numero de mes / nombre mes/ numero año

#! Preguntar al cliente si hay que agregar días también

df1 = pd.read_excel("./DatosEjemplo.xlsx", sheet_name="Hoja1")
dictFechas = {"Numero de mes": [], "Nombre mes": [], "Numero de año": []}
for x in df1['Fecha'].values:
  fecha = str(x)[:10].split("-") #Eliminar los tildes?, separar con punto?
  nMes = fecha[1]
  nombreMes = meses[int(nMes)-1]
  nAnho = fecha[0]
  dictFechas["Numero de mes"].append(nMes)
  dictFechas["Nombre mes"].append(nombreMes)
  dictFechas["Numero de año"].append(nAnho)

dfFechas = pd.DataFrame(dictFechas) 

## Tabla 2
# separar el nombre y apellido en distintas columnas
# columna de correo electronico
# generar aleatoriamente una columna n de contacto

df2 = pd.read_excel("./DatosEjemplo.xlsx", sheet_name="Hoja2")
emails = []
apellidos = []
numeros= []

for x in df2['Representante'].values:
  names = x.split(" ")  #Eliminar los tildes?, separar con punto?
  email = f"{names[0]}{names[1]}@work.com"
  numero = f"+569 {str(np.random.randint(1, 99999999)).zfill(8)}"
  apellidos.append(names[1]) 
  emails.append(email)
  numeros.append(numero) 

df2["Correo electrónico"] = emails
df2["Apellidos"] = apellidos
df2["Numero de Contacto"] = numeros

# tabla 3
# AGREGAR una columna de precio-venta y costo-producciónter, sheet_name='Tabla ', index=False)
df3 = pd.read_excel("./DatosEjemplo.xlsx", sheet_name="Hoja3")
PrecioVentas = []
CostoProducciones = []
numeros= []
CostoN = np.random.randint(100, 9999)


for x in df3['CódigoProducto'].values:
  names = x.split(" ")  #!Eliminar los tildes?, separar con punto?
  CostoProduccion = CostoN
  CostoProducciones.append(CostoProduccion) 
  CostoN = np.random.randint(1, 500)
  VentaN = CostoProduccion + CostoN
  PrecioVenta = VentaN
  PrecioVentas.append(PrecioVenta)
df3["Precio Venta"] = PrecioVentas
df3["Costo Produccion"] = CostoProducciones
print(df3.head(10))


with pd.ExcelWriter('./DatosArreglados.xlsx', engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Hoja 1', index=False)
    df2.to_excel(writer, sheet_name='Hoja 2', index=False)
    df3.to_excel(writer, sheet_name='Hoja 3', index=False)
    dfFechas.to_excel(writer, sheet_name='Hoja Fechas', index=False)