import os
import glob
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import shutil
from openpyxl.styles import Border, Font, Alignment, Side

#colocar rutas
ruta_actual = os.getcwd()
ubicacion = str(ruta_actual)+'\\data'
archivo_copiar = str(ruta_actual)+'\\Formatos\\Reporte Diario Avisos Nuevos .xlsx'
ubi_output = str(ruta_actual)+'\\Reportes excel diarios'
os.chdir(r""+ubicacion)
extension = 'xlsx'
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
#print('Rutas creadas...')

#lista de columnas    
lista_col = ['FECHA','MEDIO','EMISORA','CATEGORIA','MARCA','VERSION','DURACION','TIPO','CAMPAÑA','AGENCIA','ANUNCIANTE']

#lista meses en castellano
lista_meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

cont = 0
for f in all_filenames:
    df=pd.read_excel(f,skiprows=3)
    print('Archivo original: '+ubi_output+'\\'+str(f))

    #elegir columnas
    df=df[lista_col] 
    
    #sacar fechas
    dia_mes=''
    fechaI=min(df['FECHA'])
    fechaF=max(df['FECHA'])
    dia1=fechaI.strftime("%d")
    dia2=fechaF.strftime("%d")
    mes_num1 = int(fechaI.strftime("%m"))
    mes_num2 = int(fechaF.strftime("%m"))
    mes_nombre1 = str(lista_meses[mes_num1-1]).lower()
    mes_nombre2 = str(lista_meses[mes_num2-1]).lower()
    año1 = fechaI.strftime("%Y")
    año2 = fechaF.strftime("%Y")
    if fechaI == fechaF:
        dia_mes=str(dia1)+' de '+str(mes_nombre1)
        dia_mes_año=dia_mes+' de '+str(año1)
    else:
        if mes_num1==mes_num2:
            dia_mes=str(dia1)+'-'+str(dia2)+' de '+str(mes_nombre1)
            dia_mes_año=dia_mes+' de '+str(año1)
        else:
            if año1==año2:
                dia_mes=str(dia1)+' de '+str(mes_nombre1)+' - '+str(dia2)+' de '+str(mes_nombre2)
                dia_mes_año=dia_mes+' de '+str(año1)
            else:
                dia_mes=str(dia1)+'.'+str(mes_num1)+'.'+str(año1)+' - '+str(dia2)+'.'+str(mes_num2)+'.'+str(año2)
                dia_mes_añoI=str(dia1)+' de '+str(mes_nombre1)+' de '+str(año1)
                dia_mes_añoF=str(dia2)+' de '+str(mes_nombre2)+' de '+str(año2)                
                dia_mes_año=dia_mes_añoI+' - '+dia_mes_añoF

    #print('Dia mes: '+dia_mes)
    #print('Dia mes año: '+dia_mes_año)
    
    #copiar nuevo archivo    
    nombre_archivo_nuevo = 'Reporte Diario Avisos Nuevos '+dia_mes_año+'.xlsx'
    ruta_archivo_nuevo = ubi_output+'\\'+nombre_archivo_nuevo
    shutil.copyfile(archivo_copiar, ruta_archivo_nuevo)

    #modificar nuevo archivo
    wbkName = ruta_archivo_nuevo
    wbk = openpyxl.load_workbook(wbkName)    
    total_filas_df = int(df['FECHA'].count())
    for myRow in range(0, total_filas_df):
        col_excel = 1
        for myCol in lista_col:         
            fila_excel = myRow + 10
            celda = wbk['x'].cell(row=fila_excel, column=col_excel)
            celda.value = df[myCol][myRow]

            if myCol=='FECHA':
                celda.number_format = 'dd/mm/yyyy'
            celda.font = Font(name='Tahoma',size=8.5)
            celda.alignment = Alignment(horizontal="center", vertical="center")
            thin = Side(border_style="thin",color='A6A6A6')
            celda.border = Border(top=thin, left=thin, right=thin, bottom=thin)            
                        
            col_excel += 1

    wbk_hoja = wbk['x']
    wbk_hoja.title = dia_mes
    wbk.save(wbkName)
    wbk.close

    #agregar lista marcas a excel
    df_marcas = df['MARCA']
    df_marcas.drop_duplicates(inplace=True)
    with pd.ExcelWriter(ruta_archivo_nuevo,mode='a') as writer:
        df_marcas.to_excel(writer, sheet_name='marcas',index=False)

    print('Archivo final: '+ruta_archivo_nuevo)

a = input('presiona "Enter" para terminar...')
