
import sys
import traceback
"""
Función que busca en un csv el formato de fecha correcto, según los datos que introdujo el usuario
Básicamente agrega o elimina ceros al inicio del día y mes de ser necesario.

-Parámetros de entrada
*csv_path (str): dirección del csv
*date (str): fecha que introdujo el usuario

-Valores retornados
Fecha en el formato correcto, 0 en caso de que no se encuentre la fecha (formato dd/mm/yyyy o d/m/yyyy).
"""

def correct_date( csv_path, date ):
	try:
		if date == "" or date == None:
			return date
		date_temp = date.split('/')
		try:
			date_temp = date_temp.split('-')
		except:
			pass
		bandera = -1 #se utiliza para saber si en el día agregó ceros o eliminó posiciones
		
		#Búsqueda de alternativas a fecha
		dia = date_temp[0]
		mes = date_temp[1]
		ano = date_temp[2]
		
		#Día
		if dia[0] == "0": #Elimina el cero
			day_ = dia[1] 
			bandera = 1
		elif int(dia) < 10: #agrega ceros al principio si es menor que diez
			day_ = "0" + str(dia[0])
			bandera = 0
		elif int(dia) > 31: #Es un error
			return 0
		else:
			day_ = dia
		
		#Mes
		if mes[0] == "0" and bandera == 1: #Elimina el cero
			mes_ = mes[1]
		elif mes[0] != "0" and int(mes) < 10 and bandera == 0: #agrega ceros al principio si es menor que diez
			mes_ = "0" + mes[0]
		elif int(dia) > 9 and mes[0] == "0":
			mes_ = mes[1]
		elif int(mes) < 1 or int(mes) > 12 : #Es un error
			return 0
		else:
			mes_ = mes
			
		#Corrección de fecha original (casos d/mm/yyyy o dd/m/yyyy)
		if dia[0] != "0" and int(dia) <10:
			if mes[0] == "0":
				date = dia + "/" + mes[1] + "/" + ano
		elif dia[0] == "0" and int(mes) < 10 and mes[0] != "0":
				date =  dia + "/0" + mes + "/" + ano
		#Año
			
		date_ = day_  + "/" + mes_ + "/" + ano
		
		#Lectura del csv
		with open( csv_path ) as f:
			lineList = f.readlines()	
		
		#Se busca cuál es la fecha que está en el csv
		for line in lineList:
			if date in line:
				return date
			elif date_ in line:
				return date_
		
		#Si está aquí es porque terminó de recorrer el csv y no encontró la fecha
		return date
	except:
		print_error()
		return date	
	
"""
Función que busca en un csv el formato de fecha correcto, según los datos que introdujo el usuario
Básicamente agrega o elimina ceros al inicio del día y mes de ser necesario.

-Parámetros de entrada
*csv_path (str): dirección del csv
*time (str): hora que introdujo el usuario

-Valores retornados
Hora en el formato correcto, 0 en caso de que no se encuentre la hora (formato h:mm u hh:mm).
"""

def correct_hour( csv_path, time ):
	try:
		if time == "" or time == None:
			return date
		time_temp = time.split(':')
		
		#Búsqueda de alternativas a hora
		hora = time_temp[0]
		minut = time_temp[1]
		
		#Hora
		if hora[0] == "0": #Elimina el cero
			hora_ = hora[1] 
		elif int(hora) < 10: #agrega ceros al principio si es menor que diez
			hora_ = "0" + str(hora[0])
		elif int(hora) < 0 or int(hora) > 23: #Es un error
			return 0
		else:
			hora_ = hora
			
		time_ = hora_  + ":" + minut
		
		#Lectura del csv
		with open( csv_path ) as f:
			lineList = f.readlines()	
		
		#Se busca cuál es la fecha que está en el csv
		for line in lineList:
			if time in line:
				return time
			elif time_ in line:
				return time_
		
		#Si está aquí es porque terminó de recorrer el csv y no encontró la hora
		return time
	except:
		print_error()
		return time

def print_error():
	exc_info = sys.exc_info()
	print("\nError: ", exc_info )
	print("*************************  Información detallada del error ********************")
	for tb in traceback.format_tb(sys.exc_info()[2]):
		print(tb)

	
if __name__ == "__main__":
	csv_path = "Curva_Alimentador.csv"
	date_ = ["24/09/2016","09/09/2016","10/09/2016","01/09/2016","1/09/2016","1/9/2016","24/12/2016","33/09/2016","24/13/2016"]
	
	for date in date_:
		print("date = ", date )
		
		correcto = correct_date( csv_path, date )
		print( "Correcto = ", correcto )
	
	
	time_ = ["01:00", "1:00", "23:00", "23:59", "00:00"]
	print("*********" )
	
	for time in time_:
		correcto = correct_hour( csv_path, time )
		print( "Hora correcta = ", correcto )
	
	

	
