import pandas as pd
from tkinter import Tk, Checkbutton, Button, IntVar, Toplevel
from tkinter.filedialog import askopenfilename, asksaveasfilename

# Inicializar Tkinter
Tk().withdraw()

# Seleccionar los archivos Excel
print("Seleccione el primer archivo l_gobierno")
excel1_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

print("Seleccione el segundo archivo l_reportes")
excel2_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

# Cargar los datos de los dos archivos Excel
df1 = pd.read_excel(excel1_path)
df2 = pd.read_excel(excel2_path)

# Después de cargar los DataFrames, agregar prints de debug
print("\nÁreas en l_gobierno:")
print(df1['Area de Datos'].unique())
print("\nÁreas en l_reportes:")
print(df2['Area de Datos'].unique())

# Imprimir los roles únicos para debug
print("\nRoles únicos en l_gobierno:")
print(df1['Rol'].unique())

# Normalizar los nombres de las áreas
df1['Area de Datos'] = df1['Area de Datos'].str.strip().str.lower()
df2['Area de Datos'] = df2['Area de Datos'].str.strip().str.lower()

# Normalizar los roles
df1['Rol'] = df1['Rol'].str.strip().str.upper()

# Extraer las columnas deseadas de cada DataFrame
df1_selected = df1[['Mail', 'Rol', 'Area de Datos']]
df2_selected = df2[['Title', 'Responsable', 'Area de Datos', 'Sello', 'InforReporte']]

# Filtrar los roles de interés (ahora en mayúsculas)
df1_filtered = df1_selected[df1_selected['Rol'].isin(['DATA OWNER', 'DATA STEWARD'])]

# Verificar los datos filtrados
print("\nEjemplo de registros en gobierno:")
print(df1_filtered.head())

# Crear un diccionario para almacenar los datos organizados
data = {
    'Area de Datos': [],
    'Titulo': [],
    'InforReporte': [],
    'Desarrollador': [],
    'Sello': [],
    'Data Owner': [],
    'Data Steward 1': [],
    'Data Steward 2': [],
    'Data Steward 3': [],
    'Data Steward 4': [],
    'Data Steward 5': [],
    'Data Steward 6': []
}

# Iterar sobre cada área de datos en df2_selected
for area in df2_selected['Area de Datos'].unique():
    # Filtrar los datos por área de datos
    reportes = df2_selected[df2_selected['Area de Datos'] == area]
    gobierno = df1_filtered[df1_filtered['Area de Datos'] == area]
    
    print(f"\nProcesando área: {area}")
    print(f"Registros en gobierno para esta área:")
    print(gobierno)
    
    # Obtener los datos del reporte
    for _, reporte in reportes.iterrows():
        data['Area de Datos'].append(area)
        data['Titulo'].append(reporte['Title'])
        data['InforReporte'].append(reporte['InforReporte'])
        data['Desarrollador'].append(reporte['Responsable'])
        data['Sello'].append(reporte['Sello'])
        
        # Obtener los Data Owners y Data Stewards con más detalle
        data_owners = gobierno[gobierno['Rol'] == 'DATA OWNER']['Mail'].tolist()
        data_stewards = gobierno[gobierno['Rol'] == 'DATA STEWARD']['Mail'].tolist()
        
        print(f"Data Owners encontrados para {area}: {data_owners}")
        print(f"Data Stewards encontrados para {area}: {data_stewards}")
        
        # Asignar los Data Owners y Data Stewards
        data['Data Owner'].append(data_owners[0] if data_owners else '')
        for i in range(1, 7):
            data[f'Data Steward {i}'].append(data_stewards[i-1] if i-1 < len(data_stewards) else '')

# Crear un DataFrame con los datos organizados
result_df = pd.DataFrame(data)

# Crear ventana de selección de sellos
def seleccionar_sellos():
    ventana_sellos = Toplevel()
    ventana_sellos.title("Seleccionar Sellos a Filtrar")
    
    sellos_vars = {
        'Negocio': IntVar(),
        'Tecnología': IntVar(),
        'Seguridad': IntVar()
    }
    
    for sello in sellos_vars:
        Checkbutton(ventana_sellos, text=sello, variable=sellos_vars[sello]).pack()
    
    def confirmar():
        sellos_seleccionados = [sello for sello, var in sellos_vars.items() if var.get()]
        ventana_sellos.seleccion = sellos_seleccionados
        ventana_sellos.quit()
    
    Button(ventana_sellos, text="Confirmar", command=confirmar).pack()
    ventana_sellos.mainloop()
    return ventana_sellos.seleccion

# Obtener sellos seleccionados
sellos_filtrar = seleccionar_sellos()

if not sellos_filtrar:
    print("No se seleccionó ningún sello")
    exit()

# Filtrar por sellos pendientes
mascara = pd.Series(True, index=result_df.index)
for sello in sellos_filtrar:
    mascara &= ~result_df['Sello'].str.contains(sello, case=False, na=False)

reportes_pendientes = result_df[mascara]

# Guardar los reportes pendientes
if not reportes_pendientes.empty:
    print(f"\nSe encontraron {len(reportes_pendientes)} reportes pendientes de los sellos: {', '.join(sellos_filtrar)}")
    print("Seleccione la ubicación para guardar el archivo filtrado")
    output_path = asksaveasfilename(defaultextension=".xlsx", 
                                  filetypes=[("Excel files", "*.xlsx *.xls")])
    
    reportes_pendientes.to_excel(output_path, index=False)
    print(f"Archivo guardado en {output_path}")
else:
    print(f"\nNo se encontraron reportes pendientes de los sellos seleccionados")
