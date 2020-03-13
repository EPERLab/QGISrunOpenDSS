# Load the Pandas libraries with alias 'pd' 
import pandas as pd
from matplotlib import pyplot as plt
import os
import sys
import traceback

#Función que grafica todas las columnas de un dataframe contra sus índices
def plot_dataframe( dataframe, titulo_grafico, output_file = "", titulo_ventana = "",  x_label = "", y_label = "", rotacion = 0, save_csv = False, legend = "" ):
    try:
        cantidad_filas = len( dataframe.index )
        columnas = dataframe.columns
        horas = list( dataframe.columns )
        fig = plt.figure( titulo_ventana )
        
        #Grafica uno a uno los elementos del dataframe
        for i in range( cantidad_filas ):
            dato = list( dataframe.iloc[i].values )
            plt.plot( horas, dato )
        #plt.autoscale(enable=False, axis="y", tight=False)
        plt.title( titulo_grafico )
        plt.ylabel( y_label )
        plt.xlabel( x_label )
        plt.xticks( rotation = rotacion )
        x1,x2,y1,y2 = plt.axis()
        diferencia_y = y2 - abs( y1 )
        lim_dif = 0.003 #Valor obtenido mediante prueba y error
        #Esto es necesario para que no se escriba un subíndice arriba del eje y
        if diferencia_y < lim_dif:
            y2 = y1 + lim_dif
        diferencia_y = y2 - y1
        plt.axis((x1,x2-0.1,y1,y2)) #+ 0.001))
        #Para no tener demasiadas etiquetas en el eje x
        my_xticks = plt.xticks()[0]
        tam_myxticks = len( my_xticks )
        nuevo_xticks = []
        if tam_myxticks > 20:
            for i in range( tam_myxticks ):
                if i % 12 == 0 or i == tam_myxticks - 1:
                    nuevo_xticks.append( my_xticks[i] )
            plt.xticks(nuevo_xticks, visible = True, rotation = rotacion)
        
        #Agregar leyenda al gráfico
        if legend != "":
            plt.legend( legend )
        plt.grid( True )
        #Mostrar gráfica
        #plt.show()
        #Mostrar figura maximizada
        mng = plt.get_current_fig_manager()
        mng.window.showMaximized()        
        #Se guarda la figura en el directorio solicitado
        if output_file != "":
            pdf = ".pdf" in output_file
            #Caso en que nombre no incluya el ".pdf"
            if pdf == False:
                output_file = str( output_file + ".pdf" )
            plt.savefig(output_file, format='pdf', dpi=6000)
        
        if save_csv == True and output_file != "":
            output_file = output_file.split(".pdf")
            output_file = str( output_file[0] )
            output_file = str( output_file + '.csv' )
            dataframe.to_csv( output_file )
			
        return 
    except:
        print("\n********************************************************************************")
        print("***************************************  ERROR *********************************")
        print("********************************************************************************")
        exc_info = sys.exc_info()
        print("Error: ", exc_info )
        print("\n********************************************************************************")
        print("*************************  Información detallada del error *********************")
        print("********************************************************************************")
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)
            
def dataframe_prom_max_and_min( dataframe  ):
    try:
        prom = dataframe.mean(axis = 0)
        maxim = dataframe.max(axis = 0)
        minim = dataframe.min(axis = 0)
        
        prom_data = prom.values
        max_data = maxim.values
        min_data = minim.values
        columnas = dataframe.columns
        data = [prom_data, max_data, min_data ]
        filas = ["Promedio", "Máximo", "Mínimo"]
        
        dataframe_prom_max_min = pd.DataFrame( data = data,
                                               index = filas,
                                               columns = columnas )
        return dataframe
    except:
        print("\n********************************************************************************")
        print("***************************************  ERROR *********************************")
        print("********************************************************************************")
        exc_info = sys.exc_info()
        print("Error: ", exc_info )
        print("\n********************************************************************************")
        print("*************************  Información detallada del error *********************")
        print("********************************************************************************")
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)
        return 0            

if __name__ == "__main__":
    nombre = "Daily_Voltages_MVphaseB_01_01_2016.csv" 
    dataframe = pd.read_csv(nombre, index_col = 0 )
    print( dataframe )
    fecha = "01/01/2016"
    titulo_ventana = str("Tensión pu lv fase a por bus. Fecha: " + fecha)
    titulo_grafico = titulo_ventana
    x_label = "Hora"
    y_label = "Tensión (pu)"
    output_file = "hola"
    #def plot_dataframe( dataframe, titulo_grafico, output_file = "",  = "",  x_label = "", y_label = "", rotacion = 0, save_csv = False ):
    plot_dataframe( dataframe, titulo_grafico, titulo_ventana = titulo_ventana, x_label = x_label, y_label = y_label, rotacion = 0, output_file = output_file, save_csv = True )

