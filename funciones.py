    ### Este codigo fue construido utilizando:   ###
    ### Anaconda: v2020.02
    ### Python:   v3.7.6
    ### OS: windows 10-64bits
import pandas as pd
import math
'''
movimientos = r"C:/Users/Bryan/Dropbox/Datos/Datos Mauricio Con_Con/movientos bodega 2019.xlsx"
valorizacion = r"C:/Users/Bryan/Dropbox/Datos/Datos Mauricio Con_Con/ValorizadoMovimientos2019.xlsx"
archivo_tasa = r"C:/Users/Bryan/Dropbox/Datos/Datos Mauricio Con_Con/Tasa demanda datos Con Con.xlsx"
archivo_codigos = r"C:/Users/Bryan/Dropbox/Datos/Datos Mauricio Con_Con/Codigo y precio productos Con Con.xlsx"
'''
# t_d_concon(archivo_consumo,archivo_precio,archivo_tasa,archivo_codigos)
def t_d_concon(movimientos, valorizacion, archivo_tasa, archivo_codigos):
    '''Obtencion tasa consumo para datos del minsal'''
    
    ## Se leen los archivos de movimientos y valorizacion ##
    consumo=pd.read_excel(movimientos,
                   sheet_name=None)
    precio=pd.read_excel(valorizacion,
                         header=12,
                         sheet_name=None,
                         skiprow= list(range(12)))
    
    ## Se crea encabezado en base a las variables que estan en el excel 
    ## original y que son relevantes para el estudio
    hoja=list(consumo.keys())
    dic=dict()
    arti=None
    encabezado=["Fecha/Hora Digitación","Origen/Destino","Tipo Movimiento",
                "Saldo Anterior","Entradas","Saldo Corregido"]

    ## Cada producto leido del archivo excel se guarda en un diccionario 
    ## donde la llave corresponde al nombre del articulo y luego se genera
    ## una df correspondiente al articulo la cual se va llenando por filas
    for ws in hoja:
        for i,row in consumo[ws].iterrows():
            if 'Artículo' in str(row[0]):
                dic[consumo[ws].loc[i+1,"Unnamed: 5"]]= None
                arti=consumo[ws].loc[i+1,"Unnamed: 5"]
                df= pd.DataFrame()
            if not pd.isna(row[1]): ## si no es NA devuelve true ##
                df=df.append(row)
                dic[arti]=df
    
    ## Las df se agrupan por fecha para obtener la demanda mensual de cada producto 
    ## y se almcenan en un dicionario: suma contiene la suma de todas las columnas
    ## numericas agrupadas por mes. compras tiene la suma de todas las columnas
    ## numericas agrupadas por mes y por tipo movimiento
    ## suma: diccionario con df de los productos y su consumo (key = articulo)
    ## compras: diccionario con df de los productos y su maxima compra en el año
    suma=dict()
    compras=dict()
    productos=pd.Series(dtype=object)
    cont=0
    for k in dic:
        dic[k].dropna(axis=1, how="all", inplace=True)
        dic[k].columns=encabezado
        cont +=1
        suma[k]=dic[k].groupby([pd.Grouper(key='Fecha/Hora Digitación', freq='1M')]).sum()
        #suma[k]=suma[k].to_frame()
        suma[k]=suma[k].reset_index(drop=False)
        compras[k]=dic[k].groupby([pd.Grouper(key='Fecha/Hora Digitación', freq='1M'),'Tipo Movimiento']).sum()
        compras[k].reset_index(inplace=True, drop=False)
        compras[k]=compras[k].drop(compras[k][compras[k]['Tipo Movimiento']!='Recepción de artículos'].index)
        compras[k]['P'+str(cont)]=compras[k]['Saldo Corregido'] - compras[k]['Saldo Anterior']
        compras[k].drop(['Tipo Movimiento', 'Saldo Anterior', 'Entradas', 'Saldo Corregido'], axis=1, inplace=True)
        compras[k].set_index('Fecha/Hora Digitación', inplace=True)
        compras[k]=compras[k].max()
        suma[k].drop(['Saldo Anterior','Saldo Corregido'], axis=1, inplace=True)
        suma[k].set_index('Fecha/Hora Digitación', inplace=True)
        suma[k].rename(columns = {'Entradas':"P"+str(cont)}, inplace=True)
        productos["P"+str(cont)]=k
        
    ## Se concatenas las df dentro de compras para generar una df con el nombre
    ## del producto (ej: nombre = P1) y su Tope presupuesto de compra
    compras_f = pd.concat(compras.values(), axis=0, sort=False)   
    compras_f=compras_f.to_frame()
    compras_f.reset_index(inplace=True, drop=False) 

    ## Se genera el df con la informacion de los productos       
    precio["Hoja1"].drop(["IVA","CONSUMO","COSTO","STOCK ACTUAL",
                       "BODEGA"], inplace=True, axis=1)
    precio=precio["Hoja1"]
    productos=productos.to_frame()
    productos.reset_index(inplace=True, drop = False)
    productos.rename(columns={0:'ARTÍCULO'}, inplace=True)
    productos_f=productos.join(precio.set_index('ARTÍCULO'), on='ARTÍCULO')
    productos_f=productos_f.join(compras_f.set_index("index"), on="index")
    productos_f[0]= productos_f[0]*productos_f['PRECIO']
    productos_f.columns=['Nombre','Artículo','Valor promedio','Unidad','Tope presupuesto compra']
    productos_f.set_index('Nombre', inplace=True) 
    
    ## Se genera df con la tasa de demanda de los productos
    final = pd.concat(suma.values(), axis=1, sort=False)
    final.fillna(0, inplace=True)
    
    ## Se exportan las df a un archivo excel
    productos_f.to_excel(archivo_codigos+'.xlsx', engine='xlsxwriter', index_label="Nombre")
    final.to_excel(archivo_tasa+'.xlsx', engine='xlsxwriter', index_label="Meses")
    return()
#%%
## Cambiar segun hospital que se quiera obtener tasa ##
#data_consumo = r'C:/Users/Bryan/Dropbox/Datos/Datos_Hospital_San_Camilo/CONSUMO HOSPITAL SAN CAMILO 2017-2018.xls'
#data_compra = r'C:/Users/Bryan/Dropbox/Datos/Datos_Hospital_San_Camilo/Ordenes de compra 2017-2018 HOSCA.xls'

## Ubicacion y nombre que se quiere dar a los archivos de tasa demanda y de 
## archivo con codigos y precios.
#direc_tasa = "nombre del archivo de tasa"
#direc_info = "nombre del archivo de informacion

# t_d_hospital(data_consumo,data_compra,Ubi_nombre_tasa,Ubi_nombre_info)
def t_d_hospital(data_consumo, data_compra, direc_tasa, direc_info):
  ''' Obtencion tasa consumo para hospitales '''
                  ## Lectura de excel ##
  Consumo_Hospital = pd.read_excel(data_consumo,
                                  sheet_name=None,
                                  usecols="A:O",
                                  skiprows = [1])
    
  Compra_Hospital = pd.read_excel(data_compra,
                                    sheet_name=None,
                                    usecols="L:N,P,Q",
                                    header=1)

        ## Se obtiene dataframe por año de los archivos ##
        ## y se completa la demanda NA por 0, ademas se ##
        ## obtienen demandas de productos unificados    ##
        ## junto con su costo promedio                  ##
  fechas=list(Consumo_Hospital.keys())
  fechas.sort()
  meses=Consumo_Hospital[fechas[0]].columns[3:15]
  contador=0
  for i in fechas:
      for j in meses:
          Consumo_Hospital[i].loc[:,j].fillna(0, inplace=True)
      Consumo_Hospital[i]=Consumo_Hospital[i].groupby(["Codigo","Descripcion","Unidad"], axis=0)[meses].sum().reset_index(drop=False)

    ## Obtencion del costo promedio por producto ##
  for k in range(len(fechas)-1):
    if k==0:
        union_p=pd.concat([Compra_Hospital[fechas[k]], Compra_Hospital[fechas[k+1]]])
        union_s=Consumo_Hospital[fechas[k]].set_index(["Codigo","Unidad"]).join(Consumo_Hospital[fechas[k+1]].set_index(["Codigo","Unidad"]), how="right", rsuffix=str(contador))
    else:
        contador +=1
        union_p.concat(Compra_Hospital[fechas[k+1]], inplace=True)
        union_s=union_s.set_index(["Codigo","Unidad"]).join(Consumo_Hospital[fechas[k+1]].set_index(["Codigo","Unidad"]), how="right", rsuffix=str(contador))

    ## Se conserva la descripción de los productos del ultimo año ##
    ## Los productos nuevos son rellenados con valor 0 en demanda ##
    ## de años anteriores                                         ##
  union_s["Descripcion"]=union_s["Descripcion"+str(contador)] 
  for h in range(len(fechas)-1):
    union_s.drop(["Descripcion"+str(contador)], axis=1, inplace=True)
    contador -=1
  promedio = union_p.groupby(["Codigo Articulo","Nombre Articulo","Unidad"], axis=0)["Valor Unitario"].mean()
  promedio = promedio.to_frame()
  promedio.reset_index(inplace=True, drop=False)
  maximo = union_p.groupby(["Codigo Articulo","Nombre Articulo","Unidad"], axis=0)["Subtotal"].max()
  maximo = maximo.to_frame()
  maximo.reset_index(inplace=True, drop=False)
  promedio = promedio.set_index(["Codigo Articulo","Unidad"]).join(maximo.set_index(["Codigo Articulo","Unidad"]), how="right", rsuffix="_")
  union_s.reset_index(inplace=True, drop=False)
  promedio.reset_index(inplace=True, drop=False)
  promedio.rename(columns={"Codigo Articulo":"Codigo","Subtotal":"Tope"}, inplace=True)
  promedio.drop(["Nombre Articulo","Nombre Articulo_"], inplace=True, axis=1)
                 
      ## Se juntan los productos unicos con su precio promedio ##
      ## y se comienza a generar la estructura final de la tasa##
      ## de demanda ##
  suma=union_s.set_index(["Codigo","Unidad"]).join(promedio.set_index(["Codigo","Unidad"]), how="left")
  suma.reset_index(inplace=True, drop=False)
  data_final=suma.T
  informacion=suma[["Codigo","Unidad","Descripcion","Valor Unitario","Tope"]].copy()
  for n in data_final.columns:
    data_final.rename(columns={n:"P"+str(n+1)}, inplace=True)
    informacion.loc[n,"Nombre"]="P"+str(n+1)
  
  data_final.drop(["Unidad", "Descripcion", "Valor Unitario", "Codigo","Tope"], inplace=True, axis=0)
  data_final.reset_index(inplace= True, drop=False)
  cont=0
  for p in fechas:
    num_mes=1
    for q in meses:
        if q != "Octubre" and q != "Noviembre" and q!= "Diciembre":
            data_final.loc[cont,"index"]= str(cont+1)+"("+str(p)+"/0"+str(num_mes)+")"
        else:
            data_final.loc[cont,"index"]= str(cont+1)+"("+str(p)+"/"+str(num_mes)+")"
        cont +=1
        num_mes +=1

  ## se obtiene el archivo excel de la tasa de demanda del hospital ##
  ## con la estructura solicitdada ##
  data_final.set_index(["index"], inplace=True)
  data_final.fillna(0, inplace=True)
  data_final.to_excel(direc_tasa+".xlsx", engine='xlsxwriter', index_label="Meses")
   ## Se genera un archivo excel que relaciona el codigo del producto ##
  ## con su descripcion, precio promedio y Tope presupuestario ##

  ## Debido a que hay productos que son lo mismo pero se trabajan con unidades
  ## distintas, se debe realizar un ajuste al costo de compra de de ellos, por
  ## ejemplo: P1 se compra en unidades de KG, pero se consume tanto en KG como 
  ## en Gr (P1 esta en KG y P2 es el mismo producto pero en gramos). Se decide 
  ## para este caso particular, transformar el costo, es decir, pasar de $/KG
  ## a $/Gr. #### Se aplicara lo mismo para las demas unidades ###
  for t in range(0,len(informacion["Valor Unitario"])):
      if (math.isnan(informacion.loc[t,"Valor Unitario"])):
            for t2 in range(0,len(promedio["Codigo"])):
                if (promedio.loc[t2,"Codigo"] == informacion.loc[t,"Codigo"] and 
                    informacion.loc[t,"Unidad"]=="GR" and promedio.loc[t2,"Unidad"]=="KG"):
                       informacion.loc[t,"Valor Unitario"]=promedio.loc[t2,"Valor Unitario"]/1000
                       informacion.loc[t,"Tope"]=promedio.loc[t2,"Tope"]  

  ## Archivo excel con informacion de productos ##
  informacion.rename(columns={'Valor Unitario':'Valor promedio','Tope':'Tope presupuesto compra'},inplace=True)
  informacion.to_excel(direc_info+".xlsx", engine='xlsxwriter')
  return()
#%%
# data_consumo = r'C:\Users\Bryan\Dropbox\Datos\Datos_SAN_HCVB\cc_resumen_egresos.xlsx'
def t_d_casino(data_consumo, direc_tasa, direc_info):
                ## Lectura de excel mediante pandas ##
    ## Se recogieron las variables que, previamente en la visualizacion, ##
    ## se determinaron que son relevantes para el trabajo
    data_fns = pd.read_excel(data_consumo,
                              usecols="B:D,F,H,J,K,L:M")
    
    ## Este paso se debio incluir ya que las fechas recientes no son datos
    ## reales (son ejemplos que se crearon para demostrar el funcionamiento
    ## del programa)
    data_fns_e=data_fns.copy()
    data_fns_i=data_fns.copy()
    data_fns_e.drop('fecha_ingreso', axis=1, inplace=True)
    data_fns_i.drop('fecha_egreso', axis=1, inplace=True)
    ## Los datos estan por día, por lo que se agrupan los productos para 
    ## obtener su demanda DIARIA // cambiar freq='1M' para hacer mensual
    demanda=data_fns_e.groupby([pd.Grouper(key='fecha_egreso', freq='1D'),"cod_bar_ingreso","desc_prod","desc_unidad"])["cant_egreso"].sum()
    demanda=demanda.to_frame()
    demanda.reset_index(inplace=True, drop=False)
    ## Se obtiene el precio promedio de los productos 
    precio_promedio=data_fns_i.groupby(["cod_bar_ingreso","desc_prod","desc_unidad"])["costo_ingreso"].mean()
    precio_promedio=precio_promedio.to_frame()
    
    ## Se obtiene el tope de los productos, el mes con la maxima compra
    tope = data_fns_i.groupby([pd.Grouper(key='fecha_ingreso', freq='1D'),"cod_bar_ingreso","desc_prod","desc_unidad","hora_ingreso"])["cant_ingreso"].mean()
    tope=tope.to_frame()
    tope.reset_index(inplace = True, drop=False)
    presu=dict()
    for t in tope['cod_bar_ingreso'].unique():
        presu[t]=tope[tope["cod_bar_ingreso"]==t]["cant_ingreso"].max()
    ## Se genera el archivo con la estructura final, utilizando como indice
    ## los meses de informacion existentes en el excel
    dti = pd.date_range(start=min(data_fns_e["fecha_egreso"]), end=max(data_fns_e["fecha_egreso"]), closed="right", freq='D')
    final = pd.DataFrame(index=dti)
    
    ## Se crean dataframes temporales con cada una de las demandas mensuales
    ## de cada producto y se concatenan con el objeto "final". Se incluye 
    ## las siglas "P#" para cada producto.
    cont=1
    for i in demanda["cod_bar_ingreso"].unique():
        aux = demanda[demanda["cod_bar_ingreso"]==i][["fecha_egreso","cant_egreso"]]
        aux.set_index("fecha_egreso", inplace=True)
        aux.columns=["P"+str(cont)]
        final = pd.concat([final, aux], axis=1)
        precio_promedio.loc[i,"Nombre"]="P"+str(cont)
        cont +=1
    
    ## Meses sin demanda son reemplazados con 0 ## 
    final.fillna(0, inplace=True)
    
    ## Se agrega el valor tope de compra para cada producto ##
    precio_promedio.reset_index(inplace=True, drop=False)
    presu_df=pd.DataFrame([[key, presu[key]] for key in presu.keys()], columns=['cod_bar_ingreso', 'Tope presupuesto compra'])
    precio_promedio=precio_promedio.join(presu_df.set_index('cod_bar_ingreso'), on='cod_bar_ingreso')
    precio_promedio["Tope presupuesto compra"]=precio_promedio["Tope presupuesto compra"]*precio_promedio["costo_ingreso"].astype(int)
    ## Se renombran variables para obtener formato similar al ejemplo 
    ## de Datos HCVB
    precio_promedio.rename(columns={'Codigo':'Nombre','costo_ingreso':'Valor promedio'}, inplace=True)
    precio_promedio["Ordenar"]=precio_promedio["Nombre"]
    for orden in range(0,len(precio_promedio["Ordenar"])):
        aux=precio_promedio.loc[orden,"Ordenar"]
        aux=aux[1::]
        precio_promedio.loc[orden,"Ordenar"]=int(aux)
    precio_promedio.sort_values('Ordenar', inplace=True)
    precio_promedio.set_index("Ordenar", inplace=True, drop=True)
    precio_promedio.index.names=["Index"]
    ## Se generan los archivos excel ## 
    precio_promedio.to_excel(direc_info+".xlsx", engine='xlsxwriter')
    final.to_excel(direc_tasa+".xlsx", engine='xlsxwriter', index_label="Meses")
    return()

#%%
# consumo_anterior='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/Tasa demanda datos Con Con 2018.xlsx'
# consumo_actual='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/Tasa demanda datos Con Con 2019.xlsx'
# info_anterior='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/Codigo y precio productos Con Con 2018.xlsx'
# info_actual='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/Codigo y precio productos Con Con 2019.xlsx'
# guardar='probando'
def fusion(info_anterior,info_actual,consumo_anterior,consumo_actual,guardar='Tasa Total demanda datos Con Con'):
    '''Función que combina los datos de consumo de un año 
    con los datos de consumo del año anterior, SOLO para 
    información del MINSAL'''
    #### Para más de 2 años, se usar de forma retroactiva   ####
    #### Es decir:                                          ####
    #### datos1= Tasa total demanda datos Con Con           ####
    #### datos2
    import pandas as pd
    ## Se leen los dos archivos correspondientes a la tasa demanda de un año y 
    ## su año predecesor, junto con los archivos de informacion de productos de
    ## cada año.
    pasado=pd.read_excel(consumo_anterior)
    actual=pd.read_excel(consumo_actual)
    prod_pas=pd.read_excel(info_anterior)
    prod_act=pd.read_excel(info_actual)
    
    ## Se unen las df de informacion del año actual con el año anterior, utilizando
    ## el nombre del producto ('Artículo') como metodo de unión.
    listado=prod_act.join(prod_pas.set_index('Artículo'),on='Artículo', rsuffix='_2')
    
    ## Dado que los productos del año actual pueden ser distintos al anterior,
    ## se modifica la tasa de demanda del año anterior para que los productos 
    ## coincidan, por ej: que el P1 del año actual sea el mismo que el P1 del año
    ## anterior. 
    comp=list(listado['Nombre_2'])    
    for x in range(len(listado)):
        aux=listado.loc[x,'Nombre']
        if not (aux in comp):
            if (aux in pasado):
                del (pasado[aux])
            ## Hasta aquí, puede ocurrir que se pierda informacion ya que el 
            ## producto eliminado no necesariamente a sido descontinuado, sino
            ## que puede haber tenido algun cambio de nombre, y al no tener 
            ## codigo de identificacion de productos en el archivo original, se 
            ## dificulta el seguimiento.
        if (listado.loc[x,'Nombre'] != listado.loc[x,'Nombre_2']):
            if (type(listado.loc[x,'Nombre_2']) == str):
                actualizar=listado.loc[x,'Nombre']
                anterior=listado.loc[x,'Nombre_2']
                pasado.rename(columns={anterior:actualizar+'_'},inplace=True)
    
    ## Se usa el prefijo "_" para evitar que se dupliquen algunos nombres
    ## de las columnas en el proceso. Ahora son removidos los prefijos
    for k in pasado:
        if ('_' in k):
            borrar = k[:len(k)-1]
            pasado.rename(columns={k:borrar},inplace=True)
    
    #### Para saber si hay algun nombre de columna duplicada     ####
    #### Si len(a) es igual a a.nunique() no hay data duplicados ####
    # a=pd.Series(pasado.columns)
    # a.nunique()
    
    # Se agrega la demanda del año pasado a los productos actuales #
    # Los productos que no tenian demanda el año anterior son 
    # reemplazados por 0. Se exporta el archivo. 
    final=pd.concat([actual,pasado])
    final.sort_values('Meses',inplace=True)
    final.fillna(0, inplace=True)
    final.set_index('Meses', inplace=True)
    final.to_excel(guardar+'.xlsx', engine='xlsxwriter')
    return()
#%%
    ### Este codigo fue construido utilizando:   ###
    ### Anaconda: v2020.02
    ### Python:   v3.7.6
    ### rpy2:     v2.9.4
    ### R:        v3.6.1
    ### OS: windows 10-64 bits
## Paquetes que se requieren para usar rpy2 en conda ##
## conda install -c r rpy2
## conda install -c conda-forge tzlocal

## Importando paquetes necesarios para ejecutar el algoritmo de R 
## mediante python, generando una funcion.
def optimizacion_est(tasa,promedio,O,S,H,guardar,threshold,cuart,prod=None):
    import rpy2.robjects as robjects
    r = robjects.r
    # import rpy2.robjects.packages as rpackages
    # from rpy2.robjects.vectors import StrVector
    from rpy2.robjects import pandas2ri
    pandas2ri.activate()
    from rpy2.rinterface import RRuntimeError
    import pandas as pd
    import math
    ############################################################################
    #                ## Para poder instalar paquetes ##
    ## conda install -c conda-forge r-foreign ## en la consola ya que no se instala con los comandos de más adelante
    ## Si bien esto es más rapido y practico, se recomienda instalar 
    ## las librerias por la shell de anaconda ya que no todas las librerias
    ## son instaladas correctamentes con la secuencia de adelante.
    #utils = rpackages.importr('utils')
    # base = rpackages.importr('base')
    # select a mirror for R packages
    # utils.chooseCRANmirror(ind=0) # select the first mirror in the list
    # R package names // Esto es para instalar todas las librerias 
    # que se requieren de forma más practica
    # packnames = ('lpSolveAPI','gamlss.util','gamlss.dist','RcmdrMisc','e1071',
    #             # 'glarma','Rcmdr','forecast','effects','foreign','gamlss',
    #             # 'gamlss.data','car','carData','base','zoo','MASS','nlme',
    #             # 'sandwich','readxl')
    # Selectively install what needs to be install.
    # names_to_install = [x for x in packnames if not rpackages.isinstalled(x)]
    # if len(names_to_install) > 0:
    #    # utils.install_packages(StrVector(names_to_install))
    ############################################################################
    
    ## Se carga el archivo que contiene el algoritmo de optimizacion como
    ## una funcion definida en R, y se integra a python para luego ser llamada
    ## como una funcion más de python 
    
          ## Recordar tener los archivos en el mismo directorio de trabajo ## 
    with open('prog_est_2.r', 'r', encoding="utf-8") as file:
        content = file.read()
    r(content)
    r_estocastica=robjects.globalenv['estocastica'] 
  
    ## Se lee el archivo excel con los datos de los productos que se quiere aplicar
    ## el algoritmo de optimizacion
    ## Dataset: se coloca el archivo con las tasa de demandas de los productos 
    ## Dataprecios: se coloca el archivo con los codigos y precios de los productos
    # tasa='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/PRUEBA tasa HOSPITAL PSIQUIATRICO.xlsx'
    # promedio='C:/Users/Bryan/Desktop/UAI/Proyecto Titulación/Codigos/PRUEBA info HOSPITAL PSIQUIATRICO.xlsx'
    Dataset=pd.read_excel(tasa)
                         # sheet_name='Hoja1')
    #Dataset=pd.read_excel(r"C:/Users/Bryan/Dropbox/Datos/Datos_HCVB_Ejemplo/tasa de demanda_ejemplo.xlsx",
    #                      sheet_name='Hoja1')
    
    ## Con la variable threshold se filtran los productos que tengan un porcentaje
    ## mensual de demanda menor al señalado por el usuario. 
    # threshold=0
    eliminar=list()
    for m in Dataset.iloc[:,1::]:
        cont=0
        for o in Dataset[m]:
            if o!=0:
                cont +=1
        if cont/len(Dataset) < threshold:
            eliminar.append(m)
    Dataset.drop(eliminar, inplace=True, axis=1)
    
    #Dataprecios=pd.read_excel(r"C:/Users/Bryan/Dropbox/Datos/Datos_HCVB_Ejemplo/valorprom.xlsx",
     #                        sheet_name="Hoja2")
    Dataprecios=pd.read_excel(promedio)
                           #  sheet_name="Hoja2")
    
    ## Para hacer match entre la tasa de demanda de los productos y la informacion
    ## que se tiene de ellos en caso de presentar algun desorden los datos
    ## Se eliminan los productos de los cuales no se tiene infomación del 
    ## valor promedio o del Tope presupuesto compra.
    ## Finalmente se deja ambas df con los productos que tienen en comun, el resto
    ## es eliminado. 
    productos = list(Dataprecios["Nombre"])
    nom_pro=list(Dataset.columns[1::])
    
    for delet in productos:
        aux=list(Dataprecios.loc[Dataprecios[Dataprecios["Nombre"]==delet].index,"Valor promedio"])
        aux2=list(Dataprecios.loc[Dataprecios[Dataprecios["Nombre"]==delet].index,"Tope presupuesto compra"])
        if (math.isnan(aux[0]) or math.isnan(aux2[0])):
            Dataprecios.drop(Dataprecios[Dataprecios["Nombre"]==delet].index, inplace=True)
            if delet in nom_pro:
                Dataset.drop([delet],axis=1,inplace=True)
        if not delet in nom_pro:
            Dataprecios.drop(Dataprecios[Dataprecios["Nombre"]==delet].index, inplace=True)
   
    productos = list(Dataprecios["Nombre"])
    nom_pro=list(Dataset.columns[1::])
    for delet2 in nom_pro:
        if not delet2 in productos:
            Dataset.drop([delet2],axis=1,inplace=True)
    
    ## Los productos a los cuales se ejecutara el algoritmo ##
    prod_act=list(Dataset.columns[1::])
        
    ## Ya que se trabaja con matrix de vectores, se obtiene la ubicación
    ## de las variables "Valor promedio" y "Tope presupuesto compra" 
    DataPcol = list(Dataprecios.columns)
    col1=DataPcol.index('Valor promedio')
    col2=DataPcol.index('Tope presupuesto compra') 
    
    ## Se convierte la pandas df en una robject vector df (matriz de vectores)
    ## De esta manera se puede manipular tanto por codigo python como de R
    robjects.globalenv['Dataset'] = Dataset 
    Dataset=pandas2ri.py2ri(Dataset)    
    robjects.globalenv['Dataprecios'] = Dataprecios
    Dataprecios=pandas2ri.py2ri(Dataprecios)
    
    # Se crean variables para guardar los resultados de la optimizacion y de los 
    # errores que se van generando
    valores = dict() 
    Error_H = list() # Lista con los codigos de productos que arrojaron error en el solve(hessian)
    Error_K = dict() # Diccionario con los codigos de productos que arrojaron error en K-means y el cluster usado
    Error_Mu = dict() # solo pasa cuando Mu <= 0 con garmafit (al ingresar mu en las funciones qNBII o rNBII genera error si el parametro es <= 0  // se prueba primero cambiando modelo y luego cambiando cuartiles
    cuartil = dict() # Diccionario con los codigos de productos que arrojaron algun error anterior más de una ves y el cuartil utilizado
    arreglo = Dataset.rx(-1)
    # O=15400
    # S=999
    # H=0.29
    # cuart=50
    
    if prod:
        if prod in prod_act:
            prod_act=[prod]
        else:
            if prod[0:1] == 'p':
                prod='P'+prod[1::]
                if prod in prod_act:
                    prod_act=[prod]
                else:
                    raise NameError('Nombre no valido o inexistente')
            elif prod[0:1] =='P':
                prod='p'+prod[1::]
                if prod in prod_act:
                    prod_act=[prod]
                else:
                    raise NameError('Nombre no valido o inexistente')
            else:
                raise NameError('Nombre no valido o inexistente')
            
    for i in prod_act: ## usar len(prod_act) para todos los productos con informacion
        ## ojo que faltan precios de algunos productos en los datos de ejemplo
        try:
            valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,q=cuart)
        except Exception as RRTE:
            if ("Hessian" in str(RRTE)):
                if not (i in Error_H):
                    Error_H.append(i)
                try:
                    valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,ajuste='gamlss',q=cuart)
                except RRuntimeError:
                    for h in range (5,96,5):
                        try:
                            valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,ajuste='gamlss',q=h)
                            cuartil[i]=h
                            break
                        except RRuntimeError as error:
                            if h==95:
                                valores[i]=str(error)
                            pass
            elif ("kmeans" in str(RRTE)):
                Error_K[i] = None
                for k in range(9,1,-1):
                    try:
                        valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,k_centros=k,q=cuart)
                        Error_K[i] = k
                        break
                    except RRuntimeError as error:
                        if k==2:
                            valores[i]=str(error)
                        pass
            elif ("Mu <= 0" in str(RRTE)):
                if not (i in Error_Mu):
                    if ('qNBII' in str(RRTE)):
                        Error_Mu[i]='qNBII'
                    elif ('rNBII' in str(RRTE)):
                        Error_Mu[i]='rNBII'
                    else:
                        Error_Mu[i]=str(RRTE)
                try:
                    valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,ajuste='gamlss',q=cuart)
                except RRuntimeError:
                    for l in range (5,96,5):
                        try:
                            valores[i]=r_estocastica(arreglo[prod_act.index(i)],Dataprecios[col1][prod_act.index(i)],Dataprecios[col2][prod_act.index(i)],O,S,H,ajuste='gamlss',q=l)
                            cuartil[i]=l
                            break
                        except RRuntimeError as error:
                            if l==95:
                                valores[i]=str(error)
                            pass
            else:
                valores[i]=str(RRTE)
    
    df = pd.DataFrame(columns=valores.keys(), index=["Z","Q","I","B","CT","IS","MU","Sigma","EH","EK","EMU","CL","q","q1_glm"])
    for res in valores.keys():   
        if (res in Error_H):
            EH = "Si"
        else:
            EH = "No"
        if (res in Error_K):
            EK = "Si"
            CL = Error_K[res]
        else:
            EK = "No"
            CL = 10
        if (res in Error_Mu):
            EMU = Error_Mu[res]
        else:
            EMU = "No"
        if (res in cuartil):
            q = cuartil[res]
        else:
            q = cuart
        if (type(valores[res]) == str):
            df[res] = [valores[res].replace('\n',' ')," "," "," "," "," "," "," "," "," "," "," "," "," "]
        else:
            Z = valores[res][1][0]
            Q = valores[res][1][1]
            I = valores[res][1][2]
            B = valores[res][1][3]
            df[res] = list ([Z,Q,I,B,list(valores[res][3])[0],list(valores[res][6])[0],list(valores[res][7])[0],list(valores[res][8])[0],EH,EK,EMU,CL,q,list(valores[res][9])[0]])
    df['Parametros'] = list(['O='+str(O), 'S='+str(S), 'H='+str(H), 'threshold='+str(threshold)," "," "," "," "," "," "," "," "," "," "])
    df2 = df.T
    
    ## Se corrige el valor del Costo total para los casos en donde
    ## este valor dio negativo o un valor exacto a 1e+30. Son 
    ## reemplazados por: Z*O + P*Q + I*H + B*s
    ## Z = pedir o no
    ## O = costo por oden mensual
    ## P = precio del producto
    ## Q = cantidad del producto
    ## I = producto en inventario
    ## H = costo de almacenamiento 
    ## B = unidades producto desabastecidos
    ## s = costo por desabastecimiento (1010)
    df2['fit_CT']=None
    for correcion in valores:
        try:
            if (df2.loc[correcion,"CT"] == 1e+30 or df2.loc[correcion,"CT"] <= 0):
                df2.loc[correcion,"CT"] = df2.loc[correcion,"Z"]*O + Dataprecios[col1][prod_act.index(correcion)]*df2.loc[correcion,"Q"] + df2.loc[correcion,"I"]*H + df2.loc[correcion,"B"]*S
                df2.loc[correcion,"fit_CT"] = 'Si'
            else:
                df2.loc[correcion,"fit_CT"] = 'No'
        except:
            pass
            
        
    #Se exporta el archivo a formato csv ya que es más versatil a la hora de ser
    #usado por otras aplicaciones.
    df2.to_csv(guardar, sep=";")
    # df2.to_excel('resul_prueba_concon.xlsx', engine='xlsxwriter')
    return()

'''
Observaciones:
    Con las ultimas modificaciones se logro correr el algoritmo para
    todos los productos de Datos Concón que tienen información de 
    valor promedio y tope presupuesto compra.

'''
#%%