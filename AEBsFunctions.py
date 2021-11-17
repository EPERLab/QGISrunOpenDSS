import numpy as np

"""
########################################################################
Función que descomenta la líneas que el programa pudo comentar anteriormente
y crea una lista ordenada de todos los dssname existentes.

Parámetros de entrada:
*file_path (str): dirección del archivo a ser analizado.

Parámetros de salida:
*dss_name_dict (dict): diccionario con la lista de dss de buses
"""
def uncomment_lines(file_path):
    with open(file_path, "r+") as f:
        lineList = f.readlines()
        f.seek(0)
        dss_name_dict = {}
        
        for line in lineList:
            
            if line[:2] == "!!": #aquí es donde se descomenta
                line = line[2:]
            f.write(line)
            in_dssname = line.find("storage.")
            
            if in_dssname == -1:
                f.write(line)
                continue
            in_dssname += len("storage.")
            fin_dssname = line.find(" ", in_dssname)
            fin_plantelname = line.rfind("_",0, fin_dssname)
            dssname = line[in_dssname:fin_dssname]
            dssplantelname = line[in_dssname:fin_plantelname]            
            if "monitor" not in line:
                if dssplantelname not in dss_name_dict:
                    dss_name_dict[dssplantelname] = [dssname]
                else:
                    vector_names = dss_name_dict[dssplantelname]
                    vector_names.append(dssname)
                    dss_name_dict[dssplantelname] = vector_names            
        f.truncate()
        
    return dss_name_dict

"""
########################################################################
Función 

Parámetros de entrada:
*dss_name_dict (dict): diccionario con la lista de dss de buses
*file_path (str): dirección del archivo a ser analizado.
*percentaje_aebs (int): porcentaje de AEBs a ser analizados

Parámetros de salida:
No retorna nada

"""
    
def comment_lines(dss_name_dict, file_path, percentaje_aebs):
    dss_name_dictnew = {}
    if float(percentaje_aebs) > 1:
        percentaje_aebs = float(percentaje_aebs)/100
    #Se reacomoda el vector de nombres según el porcentaje seleccionado
    for key in dss_name_dict:
        dssname_vect = dss_name_dict[key]     
        longitud_vector = len(dssname_vect)        
        len_nueva = longitud_vector*percentaje_aebs
        len_nueva = int(len_nueva)
        vector_nuevo = dssname_vect
        np.random.shuffle(vector_nuevo)
        vector_nuevo = vector_nuevo[:len_nueva]
        dss_name_dict[key] = vector_nuevo
        
    with open(file_path, "r+") as f:
        lineList = f.readlines()
        f.seek(0)
        
        for line in lineList:
            if line[0] == "!" and line[1] != "!": #comentario del usuario
                f.write(line)
                continue
            in_dssname = line.find("storage.")
            
            if in_dssname == -1:
                f.write(line)
                continue
            
            in_dssname += len("storage.")
            fin_dssname = line.find(" ", in_dssname)
            fin_plantelname = line.rfind("_",0, fin_dssname)
            dssname = line[in_dssname:fin_dssname]
            dssplantelname = line[in_dssname:fin_plantelname]
            
            vector_dss = dss_name_dict[dssplantelname]
            if dssname not in vector_dss:
                line = "!!" + line
            f.write(line)
        f.truncate()
            
    
    
    return 1
   
    
if __name__ == "__main__":
    file_path = "CAR_StorageBuses.dss"
    vector_dss = uncomment_lines(file_path)
    print( vector_dss )
    comment_lines( vector_dss, file_path, 1)
    
