    ### Este codigo fue construido utilizando:   ###
    ### Anaconda: v2020.02
    ### Python:   v3.7.6
    ### OS: windows 10-64bits

#import sys                  ## Se agrega al archivo spec
#sys.setrecursionlimit(5000)

#2 Para salir
def menu_principal(): #0
    
    layout = [ [sg.Text('Seleccione la activdad que desea realizar', size=(30,3))],
               [sg.Frame('Funciones',  [
                    [sg.Text('Para',size=(12,1)), sg.Text('Para',size=(12,1)), sg.Text('Para',size=(4,1))],
                    [sg.Button('Tasas',size=(12,4), tooltip='Funcion para obtener demanda mensual\n de los productos, valor promedio y tope presupuesto compra'), sg.Button('Fusionar',size=(12,4), tooltip='Une las demandas de dos años consecutivos\n (solo para datos concon)'), sg.Button('Optimizar',size=(12,4), tooltip='Para ejecutar el algoritmo de programación estocastica')]],element_justification='center',title_location='n', )],
               [sg.Text(' ',size=(30,1)), sg.Cancel('Salir', size=(10,2))] ]
    
    window = sg.Window('Menu principal de funciones', layout)
    
    while True:
        event, values = window.read()
        if (event == sg.WIN_CLOSED or event == 'Salir'):
            change=2
            break
        if  event == 'Fusionar':
            change=3
            break
        if event == 'Optimizar':
            change=1
            break
        if event == 'Tasas':
            change=4
            break
    
    window.close()
    return(change)

def ejec_tasa(): #4
    
    choices = ('Datos Concon','Datos Hospitales','Datos Casinos')
    
    layout = [  [sg.Text('Datos de consumo', size=(16, 1)), sg.Input(), sg.FileBrowse()],
                [sg.Text('Datos de productos', size=(16, 1)), sg.Input(), sg.FileBrowse()],
                [sg.Text('*Para datos HCVB usar\n solo "Datos consumo"*',auto_size_text=True), sg.Listbox(choices, size=(25, len(choices)), key='-Datos-')],
                [sg.SaveAs('Guardar consumo\n como: (Opcional)', size=(16,2)), sg.InputText(size=(50,1))],
                [sg.SaveAs('Guardar información\n como: (Opcional)', size=(16,2)), sg.InputText(size=(50,1))],
                [sg.Button('Aceptar',size=(10,2), tooltip='Genera dos archivos excel, uno con la tasa de demanda mensual\n y otro excel con información de los productos.'), sg.Text(' ',size=(41,1)), sg.Cancel('Volver',size=(10,2))]]
    
    window = sg.Window('Capturando archivos excel', layout)
    
    while True:                  # the event loop
        event, values = window.read()
        if (event == sg.WIN_CLOSED):
            change=2
            break
        if  event == 'Volver':
            change=0
            break
        if event == 'Aceptar':
            if (values['-Datos-'] and values[0]):
                cont = 0
                v0=values[0].split('/')
                v1=values[1].split('/')
                if (len(v0)==len(v1)):
                    for aux in range(len(v0)):
                        if v0[aux] != v1[aux]:
                            cont +=1
                        if cont >= 2:
                            break
                if(len(values[0].split('/')) != len(values[1].split('/')) or cont >=2):
                    precaucion = sg.popup_yes_no('Archivos con distinto directorio: ¿Desea continuar?')
                    if precaucion == 'Yes':
                        if (values['-Datos-'][0] == 'Datos Concon'):
                            try:
                                aux=values[0].split('/')
                                aux2=values[1].split('/')
                                if values[2]:
                                    nombre_tasa=values[2]
                                else:
                                    nombre_tasa='Tasa demanda datos Concon '+aux[len(aux)-1][aux[len(aux)-1].find('2'):aux[len(aux)-1].find('2')+4]
                                if values[3]:
                                    nombre_info=values[3]
                                else:
                                    nombre_info='Codigo y precio productos Concon '+aux2[len(aux2)-1][aux2[len(aux2)-1].find('2'):aux2[len(aux2)-1].find('2')+4]
                                window.hide()
                                sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                                ct.t_d_concon(values[0],values[1],nombre_tasa,nombre_info)
                                window.UnHide()
                            except Exception as Error:
                                sg.popup(str(Error),keep_on_top=True,title='ERROR')
                                window.UnHide()
                        if (values['-Datos-'][0] == 'Datos Hospitales'):
                            try:
                                aux=values[0].split('/')
                                aux2=values[1].split('/')
                                nombre=values[0][values[0].find("CONSUMO")+8:len(values[0])-14]
                                if values[2]:
                                    nombre_tasa=values[2]
                                else:
                                    nombre_tasa='Tasa demanda datos '+nombre
                                if values[3]:
                                    nombre_info=values[3]
                                else:
                                    nombre_info='Codigo y precio productos '+nombre
                                window.hide()
                                sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                                ct.t_d_hospital(values[0],values[1],nombre_tasa,nombre_info)
                                window.UnHide()
                            except Exception as Error:
                                sg.popup(str(Error),keep_on_top=True,title='ERROR')
                                window.UnHide()
                        if (values['-Datos-'][0] == 'Datos Casinos'):
                            try:
                                if values[2]:
                                    nombre_tasa=values[2]
                                else:
                                    nombre_tasa='Tasa demanda '+values['-Datos-'][0]
                                if values[3]:
                                    nombre_info=values[3]
                                else:
                                    nombre_info='Codigo y precio productos '+values['-Datos-'][0]
                                window.hide()
                                sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                                ct.t_d_casino(values[0],nombre_tasa,nombre_info)
                                window.UnHide()
                            except Exception as Error:
                                sg.popup(str(Error),keep_on_top=True,title='ERROR')
                                window.UnHide()
                else:
                    if (values['-Datos-'][0] == 'Datos Concon'):
                        try:
                            aux=values[0].split('/')
                            aux2=values[1].split('/')
                            if values[2]:
                                nombre_tasa=values[2]
                            else:
                                nombre_tasa='Tasa demanda datos Concon '+aux[len(aux)-1][aux[len(aux)-1].find('2'):aux[len(aux)-1].find('2')+4]
                            if values[3]:
                                nombre_info=values[3]
                            else:
                                nombre_info='Codigo y precio productos Concon '+aux2[len(aux2)-1][aux2[len(aux2)-1].find('2'):aux2[len(aux2)-1].find('2')+4]
                            window.hide()
                            sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                            ct.t_d_concon(values[0],values[1],nombre_tasa,nombre_info)
                            window.UnHide()
                        except Exception as Error:
                            sg.popup(str(Error),keep_on_top=True,title='ERROR')
                            window.UnHide()
                    if (values['-Datos-'][0] == 'Datos Hospitales'):
                        try:
                            aux=values[0].split('/')
                            aux2=values[1].split('/')
                            nombre=values[0][values[0].find("CONSUMO")+8:len(values[0])-14]
                            if values[2]:
                                nombre_tasa=values[2]
                            else:
                                nombre_tasa='Tasa demanda datos '+nombre
                            if values[3]:
                                nombre_info=values[3]
                            else:
                                nombre_info='Codigo y precio productos '+nombre
                            window.hide()
                            sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                            ct.t_d_hospital(values[0],values[1],nombre_tasa,nombre_info)
                            window.UnHide()
                        except Exception as Error:
                            sg.popup(str(Error),keep_on_top=True,title='ERROR')
                            window.UnHide()
                    if (values['-Datos-'][0] == 'Datos Casinos'):
                        try:
                            aux=values[0].split('/')
                            nombre=aux[len(aux)-1]
                            if values[2]:
                                nombre_tasa=values[2]
                            else:
                                nombre_tasa='Tasa demanda '+nombre
                            if values[3]:
                                nombre_info=values[3]
                            else:
                                nombre_info='Codigo y precio productos '+nombre
                            window.hide()
                            sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                            ct.t_d_casino(values[0],nombre_tasa,nombre_info)
                            window.UnHide()
                        except Exception as Error:
                            sg.popup(str(Error),keep_on_top=True,title='ERROR') 
                            window.UnHide()
    
    window.close()
    return(change)
#%%
def ejec_algoritmo(): #1 Optimizar
    
    layout = [  [sg.Text('Datos de consumo', size=(15, 1)), sg.Input(size=(52,1)), sg.FileBrowse()],
                [sg.Text('Datos de productos', size=(15, 1)), sg.Input(size=(52,1)), sg.FileBrowse()],
                [sg.Frame('', [
                [sg.Text('Costo pedido mensual     (o)', size=(22,1)), sg.InputText(size=(9,1)), sg.Text('Ingrese decimales usando punto',size=(25,1))],
                [sg.Text('Costo desabastecimiento (s)', size=(22,1)),sg.InputText(size=(9,1)), sg.Text('Archivo csv guardado por defecto ',size=(25,1))],
                [sg.Text('Costo almacenamiento    (h)', size=(22,1)),sg.InputText(size=(9,1)), sg.Text('como: Resultados', size=(12,1))],
                [sg.Text(' ',size=(4,1)), sg.Text('Cuartil, default=50 (q)', size=(16,1)), sg.InputText(size=(9,1), key='cuartil')]], element_justification='left'),sg.Frame('Threshold', [[sg.Text('0<=T<100', size=(8,1))],[sg.Input(size=(4,1),key='thresh',tooltip='Indique el porcentaje de meses con demanda distinta de 0 exigido para trabajar')]])],
                [sg.Text('Ejecutar para producto:', size=(17,1)),sg.InputText(size=(6,1),key='prod_esp'),sg.Text('*EJ: "P1" solo corre codigo para ese producto,\n dejar en blanco para todos los productos', size=(40,2))],
                [sg.SaveAs('Guardar como:\n (opcional)',size=(12,2)), sg.InputText(size=(52,3))],
                [sg.Button('Optimizar', size=(8,2), tooltip='Genera un archivo csv con el resultado de la optimización'), sg.Text(' ',size=(50,1)), sg.Cancel('Volver', size=(8,2))]  ]
    
    window = sg.Window('Capturando archivos excel', layout)
    
    while (True):
        event2, values2 = window.read(100)
        if event2 == sg.WIN_CLOSED:
            change=2
            break
        if event2 == 'Volver':
            change=0
            break
        if event2 == 'Optimizar':
            if (values2[0] and values2[1] and values2[2] and values2[3] and values2[4]):
                try:
                    if values2['thresh']:
                        values2['thresh'] = float(values2['thresh'])
                    else:
                        values2['thresh'] = 0
                    if (values2['thresh'] < 0 or values2['thresh'] >= 100):
                        raise Exception('Treshold no valido')
                    Dataset_d = values2[0]
                    Dataprecios_d = values2[1]
                    if values2['cuartil']:
                        qu=values2['cuartil']
                    else:
                        qu=50
                    if values2[5]:
                        save_as=values2[5]
                    else:
                        save_as='Resultados'
                    for i in range(2,5):
                        values2[i]=float(values2[i])
                    if values2['prod_esp']:
                        p_ejec=values2['prod_esp']
                    else:
                        p_ejec=None
                    window.hide()
                    sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                    ct.optimizacion_est(Dataset_d,Dataprecios_d,values2[2],values2[3],values2[4],guardar=save_as,threshold=values2['thresh']/100,cuart=qu,prod=p_ejec)
                    window.UnHide()
                except Exception as Error:
                    sg.popup(str(Error),keep_on_top=True,title='ERROR')
                    window.UnHide()
            else:
                sg.popup("No se completo la informacion primordial") 
    
    window.close()
    return(change)

#%%
def ejec_fusion(): #3
    layout = [[sg.Text('Codigo y precio año anterior',size=(20,1)),sg.Input(), sg.FileBrowse()],
               [sg.Text('Codigo y precio año actual', size=(20,1)),sg.Input(),sg.FileBrowse()],
               [sg.Text('Tasa demanda año anterior', size=(20,1)),sg.Input(),sg.FileBrowse()],
               [sg.Text('Tasa demanda año actual', size=(20,1)),sg.Input(),sg.FileBrowse()],
               [sg.SaveAs('Guardar como:\n (Opcional)',size=(12,2)),sg.InputText(size=(56,1))],
               [sg.Text(' ', size=(15,1)), sg.Text('Para que funcione correctamente debe respetar el orden de los archivos. Si ya realizo una fusión, utilizar este archivo como año anterior y anexar el año siguiente.', size =(40,4))],
               [sg.Button('Fusionar',size=(8,2), tooltip='Genera un archivo excel que contiene la demanda mensual de los dos años ingresados.'), sg.Text(' ',size=(49,1)), sg.Button('Volver',size=(8,2))]]
    
    window =  sg.Window('Archivos de Concon a fusionar', layout)
    
    global mensaje1
    if mensaje1==0:
        sg.Popup('Solo para datos Concon, seleccione\n "Ok" para continuar')
        mensaje1 = 1
    
    while True:
        event, values =window.read()
        if event == 'Volver':
            change=0
            break
        if event == sg.WIN_CLOSED:
            change=2
            break
        if event == 'Fusionar':
            if (values[0] and values[1] and values[2] and values[3]):
                if values[4]:
                    try:
                        window.hide()
                        sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                        ct.fusion(values[0],values[1],values[2],values[3],values[4])
                        window.UnHide()
                    except Exception as error:
                        sg.popup(str(error),keep_on_top=True,title='ERROR')
                        window.UnHide()
                else:
                    try:
                        window.hide()
                        sg.PopupNoButtons('Ejecutando tarea, la ventana principal aparecerá\n cuando finalize el proceso.\n Cierre esta pestaña una vez reaparezca menu principal \n para evitar errores', title='Procesando', non_blocking=True, keep_on_top=True)
                        ct.fusion(values[0],values[1],values[2],values[3])
                        window.UnHide()
                    except Exception as error:
                        sg.popup(str(error),keep_on_top=True, title='ERROR')
                        window.UnHide()
    window.close()
    return(change)

#%%
if __name__=='__main__':
    import PySimpleGUI as sg
    import funciones as ct
    change=0
    mensaje1=0
    while change != 2:
        while change==0:
            change=menu_principal()   
        
        while change==1:
            change=ejec_algoritmo()
        
        while change==3:
            change=ejec_fusion()
        
        while change==4:
            change=ejec_tasa()
    
