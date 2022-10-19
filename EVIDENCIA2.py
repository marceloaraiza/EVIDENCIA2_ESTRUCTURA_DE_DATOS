import datetime
import sys
from tabulate import tabulate
import csv
import openpyxl

reservaciones={}
clientes={}
salas={}
lista_reporte=[]
lista_reporte_excel=[]
list_reservaciones=[]
list_salas = []
list_clientes = []

def menu_principal():
    print("xxxxxxxxxxxxMENU PRINCIPALxxxxxxxxxxxx")
    print("[1] RESERVACIONES.")
    print("[2] REPORTES.")
    print("[3] REGISTRAR UNA NUEVA SALA.")
    print("[4] REGISTRAR UN NUEVO CLIENTE.")
    print("[5] SALIR.")

def listado_clientes():
    for clave, datos in list(clientes.items()):
        registro_cliente = clave,datos
        list_clientes.append(registro_cliente)

    print (tabulate(list(list_clientes),headers=["NUM CLIENTE","NOMBRE"],tablefmt='grid'))
    print("")

    list_clientes.clear()


def listado_reservaciones():
    for clave, datos in list(reservaciones.items()):
        registro_reservacion = clave,datos[0],datos[1],datos[2],datos[3],datos[4]
        list_reservaciones.append(registro_reservacion)

    print (tabulate(list(list_reservaciones),headers=["NUM RESERVACION","SALA","NOMBRE EVENTO","CLIENTE","TURNO","FECHA"],tablefmt='grid'))
    print("")

    list_reservaciones.clear()

def listado_salas():
    for llave, valor in list(salas.items()):
        registro_sala = llave,valor[0],valor[1]
        list_salas.append(registro_sala)

    print (tabulate(list(list_salas),headers=["NUM SALA","NOMBRE SALA","CAPACIDAD"],tablefmt='grid'))
    print("")
    list_salas.clear()


def capturafechareservacion():

    while True:
        cadena_fecha_reservacion = input("INGRESE LA FECHA DE RESERVACIÓN EN EL FORMATO (DD/MM/AAAA): \n")
        try:
            fecha_reservacion = datetime.datetime.strptime(cadena_fecha_reservacion, "%d/%m/%Y")
            fecha_reservacion_procesada = (fecha_reservacion - datetime.timedelta(days=+2)).date()
            fecha_actual = datetime.date.today()
            if fecha_reservacion_procesada>=fecha_actual:
                break
            else:
                print("LA RESERVACIONES SOLO SE PUEDEN REALIZAR COMO MINIMO CON DOS DIAS DE ANTICIPACIÓN. ")
                continue

        except Exception:
            print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
    print("FECHA DE RESERVACION CAPTURADA CORRECTAMENTE.")
    return fecha_reservacion.strftime('%d/%m/%Y')

def actualizar_reservacion():
    while True:
        print("LISTADO DE RESERVACIONES: ")
        listado_reservaciones()
        str_llave_reservaciones = input("INGRESE EL NUMERO DE RESERVACION QUE SE VA MODIFICAR: \n")
        try:
            llave_reservaciones = int(str_llave_reservaciones)
            if llave_reservaciones in reservaciones:
                for llave,valores in reservaciones.items():
                    if llave==llave_reservaciones:
                        nvo_nombre = input("INGRESE EL NUEVO NOMBRE DEL EVENTO \n").upper()
                        reservaciones[llave]=(valores[0],nvo_nombre,valores[2],valores[3],valores[4])
                        print("EL NOMBRE DE LA RESERVACIÓN SE MODIFICÓ.")
                        print("")
                        break
            else:
                print("RESERVACIÓN NO ENCONTRADA")
        except Exception:
            print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
            continue
        break


try:
    with open("clientes.csv","r", newline="") as archivo1:
        lector = csv.reader(archivo1)
        next(lector)
        for clave, nombre in lector:
            clientes[int(clave)] = (nombre)

    with open("salas.csv","r", newline="") as archivo2:
        lector = csv.reader(archivo2)
        next(lector)
        for clave, nombre,capacidad in lector:
            salas[int(clave)] = (nombre,int(capacidad))

    with open("reservaciones.csv","r", newline="") as archivo3:
        lector = csv.reader(archivo3)
        next(lector)
        for clave,num_sala, nombre_evento,nom_cliente,turno,fecha_reservacion in lector:
            reservaciones[int(clave)] = (int(num_sala),nombre_evento,nom_cliente,turno,fecha_reservacion)

except Exception:
    print("NO SE ENCONTRO INFORMACION ANTERIOR, SE CONSIDERARÁ A ESTA LA PRIMERA EJECUCION DEL DESARROLLO")

while True:
    menu_principal()
    while True:
        dato_menu = input("INGRESE EL NUMERO DE OPCION DEL MENU PRINCIPAL QUE SE DESEA REALIZAR, INGRESE SOLAMENTE NUMEROS: \n")
        try:
            opcion_principal = int(dato_menu)
            break
        except Exception:
            print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
    if opcion_principal==1:
        print("")
        print("*****SUBMENU RESERVACIONES*****")
        print("[1] REGISTRAR UNA RESERVACION.")
        print("[2] MODIFICAR NOMBRE DE UNA RESERVACION.")
        print("[3] CONSULTAR DISPONIBILIDAD DE SALAS PARA UNA FECHA.")
        while True:
            dato_submenu = input("INGRESE EL NUMERO DE OPCION DEL SUBMENU QUE SE DESEA REALIZAR, INGRESE SOLAMENTE NUMEROS: \n")
            try:
                opcion_submenu = int(dato_submenu)
                break
            except Exception:
                print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
        if opcion_submenu==1:
            generador_llave_reservaciones=max(list(reservaciones.keys()),default=0) + 1
            while True:
                print("SALAS REGISTRADAS: ")
                listado_salas()
                sala=input("INGRESE EL NUMERO DE LA SALA EN DONDE SE HARA LA RESERVACIÓN: \n")
                try:
                    num_sala = int(sala)
                    if num_sala in salas:
                        print("SALA EXISTENTE.")
                        print("")
                        break
                    else:
                        print("SALA NO EXISTE.")
                        continue
                except Exception:
                    print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
                    print("SOLO SE PERMITEN NUMEROS ENTEROS.")

            while True:
                print("CLIENTES REGISTRADOS: ")
                listado_clientes()
                cliente  = input("INGRESE EL NUMERO DE CLIENTE QUE REALIZA LA RESERVACIÓN: \n")
                try:
                    num_cliente = int(cliente)
                    if num_cliente in clientes:
                        print("CLIENTE EXISTE.")
                        nom_cliente = clientes[num_cliente]
                        print("")
                        break
                    else:
                        print("CLIENTE NO EXISTE.")
                        continue
                except Exception:
                    print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
                    print("SOLO SE PERMITEN NUMEROS ENTEROS.")
                    print("")

            while True:
                turno = input("INGRESE EL TURNO DEL EVENTO: [MATUTINO] [VESPERTINO] [NOCTURNO] \n").upper()
                if (turno ==""):
                    print("EL TURNO NO DEBE OMITIRSE")
                    continue
                elif (turno =="MATUTINO") or (turno =="VESPERTINO") or (turno =="NOCTURNO"):
                    print("TURNO CAPTURADO")
                    print("")
                    break
                else:
                    print("TURNO NO DISPONIBLE.")

            fecha_reservacion = capturafechareservacion()
            print("")

            while True:
                nombre_evento=input("INGRESE EL NOMBRE DEL EVENTO: \n").upper()
                if (nombre_evento ==""):
                    print("EL NOMBRE DEL EVENTO NO DEBE OMITIRSE")
                    continue
                else:
                    print("NOMBRE DEL EVENTO CAPTURADO CORRECTAMENTE.")
                    print("")
                    break

            for clave, datos in list(reservaciones.items()):
                if (num_sala,turno,fecha_reservacion) == (datos[0],datos[3],datos[4]):
                    print("NO SE PUEDE TENER DOS RESERVACIONES AL MISMO TIEMPO.")
                    print("")
                    break
            else:
                reservaciones[generador_llave_reservaciones]=(num_sala,nombre_evento,nom_cliente,turno,fecha_reservacion)
                print("RESERVACION REGISTRADA CORRECTAMENTE.")
                print("")
        elif opcion_submenu==2:
            actualizar_reservacion()
            print("")
        elif opcion_submenu==3:
            while True:
                cadena_fecha_disp = input("INGRESE LA FECHA PARA CONSULTAR DISPONIBILIDAD DE SALAS EN EL FORMATO (DD/MM/AAAA): \n")
                try:
                    fecha_valida = datetime.datetime.strptime(cadena_fecha_disp, "%d/%m/%Y")
                    fecha_disp = fecha_valida.date()
                    fecha_disp = fecha_disp.strftime('%d/%m/%Y')
                    break
                except Exception:
                    print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
                    print("")
                    continue


            salas_ocupadas = set()
            for clave,valor in reservaciones.items():
                if valor[4]==fecha_disp:
                    nom_sala = salas[clave][0]
                    tup_reservadas = clave,nom_sala,valor[3]
                    salas_ocupadas.add(tup_reservadas)

            posibles_salas_disp = set()
            turnos = ["MATUTINO","VESPERTINO","NOCTURNO",]

            for clave,valor in salas.items():
                for turno in turnos:
                    tup_salas = clave,valor[0],turno
                    posibles_salas_disp.add(tup_salas)

            salas_disponibles = posibles_salas_disp - salas_ocupadas


            list_salas_disp = list(sorted(salas_disponibles))
            print("")
            print(f"REPORTE DE SALAS DISPONIBLES PARA RESERVAR DE LA FECHA: {fecha_disp}.")
            print (tabulate(list(list_salas_disp),headers=["NUMERO SALA","NOMBRE SALA","TURNO"],tablefmt='grid'))

        else:
            print("OPCIÓN NO DISPONIBLE.")
    elif opcion_principal==2:
        print("SUBMENU REPEORTES.")
        print("[1] REPORTE EN PANTALLA PARA UNA FECHA ESPECIFICA. ")
        print("[2] EXPORTAR REPORTE A EXCEL PARA UNA FECHA ESPECIFICA. ")

        while True:
            dato_submenu2 = input("INGRESE EL NUMERO DE OPCION DEL SUBMENU QUE SE DESEA REALIZAR, INGRESE SOLAMENTE NUMEROS: \n")
            try:
                opcion_submenu2 = int(dato_submenu2)
                break
            except Exception:
                print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")


        if opcion_submenu2==1:
            while True:
                cadena_fecha_consulta = input("INGRESE LA FECHA DE LAS RESERVACIÓNES PARA CONSULTAR EN EL FORMATO (DD/MM/AAAA): \n")
                try:
                    fecha_valida = datetime.datetime.strptime(cadena_fecha_consulta, "%d/%m/%Y")
                    fecha_consulta = fecha_valida.date()
                    fecha_consulta = fecha_consulta.strftime('%d/%m/%Y')
                    break
                except Exception:
                    print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
                    print("")
                    continue
            lista_reporte.clear()
            for clave, datos in list(reservaciones.items()):
                if datos[4]==fecha_consulta:
                    lista_reporte.append(datos)


            print(f"REPORTE DE RESERVACIONES AL DIA {fecha_consulta}")
            print (tabulate(list(lista_reporte),headers=["SALA","NOMBRE EVENTO","CLIENTE","TURNO","FECHA"],tablefmt='grid'))
            print("")

        elif opcion_submenu2==2:
            while True:
                cadena_fecha_consulta = input("INGRESE LA FECHA EN EL FORMATO (DD/MM/AAAA): \n")
                try:
                    fecha_valida = datetime.datetime.strptime(cadena_fecha_consulta, "%d/%m/%Y")
                    fecha_consulta = fecha_valida.date()
                    fecha_consulta = fecha_consulta.strftime('%d/%m/%Y')
                    break
                except Exception:
                    print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")
                    print("")

            lista_reporte_excel.clear()
            for clave, datos in list(reservaciones.items()):
                if datos[4]==fecha_consulta:
                    lista_reporte_excel.append(datos)

            libro = openpyxl.Workbook()
            hoja = libro["Sheet"]


            hoja.title = "REPORTE"

            hoja.append(('NUM SALA','NOMBRE EVENTO','CLIENTE','TURNO','FECHA'))

            for reporte in lista_reporte_excel:
                hoja.append(reporte)

            libro.save("ReporteReservaciones.xlsx")
            print("SE EXPORTO EL REPORTE TABULAR A UN ARCHIVO DE EXCEL.")
        else:
            print("OPCION NO DISPONIBLE")

    elif opcion_principal==3:
        while True:
            nombre_sala = input("INGRESE EL NOMBRE DE LA SALA QUE SE VA REGISTRAR: \n").upper()
            if (nombre_sala==""):
                print("EL NOMBRE DE LA SALA NO DEBE OMITIRSE.")
                continue
            else:
                break

        while True:
            cap_sala = input("INGRESE LA CAPACIDAD DE PERSONAS DE LA SALA: \n")
            try:
                if cap_sala == "":
                    print("EL CUPO DE LA SALA NO DEBE OMITIRSE.")
                cupo_sala = int(cap_sala)
                if cupo_sala<1:
                    print("EL DATO NO DEBE SER MENOR QUE CERO.")
                    continue
                break
            except Exception:
                print(f"LO SIENTO :( OCURRIO UNA EXCEPCIÓN DE TIPO: {sys.exc_info()[0]}")

        generador_llave_sala=max(list(salas.keys()),default=0) + 1
        salas[generador_llave_sala]= (nombre_sala,cupo_sala)
        print("SALA REGISTRADA CORRECTAMENTE. ")
        print("")
    elif opcion_principal==4:
        while True:
            nombre_cliente = input("INGRESE EL NOMBRE DEL CLIENTE QUE SE VA REGISTRAR: \n").upper()
            if (nombre_cliente==""):
                print("EL NOMBRE DEL CLIENTE NO DEBE OMITIRSE.")
                continue
            else:
                generador_llave_cliente=max(list(clientes.keys()),default=0) + 1
                clientes[generador_llave_cliente]= (nombre_cliente)
                print("CLIENTE REGISTRADADO CORRECTAMENTE. ")
                print("")
                break
    elif opcion_principal==5:
        with open("reservaciones.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("NUM RESERVACION","NUM SALA","NOMBRE EVENTO","NOMBRBE CLIENTE","TURNO","FECHA"))
            grabador.writerows([(clave, datos[0], datos[1],datos[2],datos[3],datos[4]) for clave, datos in reservaciones.items()])

        with open("clientes.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("KEY","NOMBRE CLIENTE"))
            grabador.writerows([(clave, datos) for clave, datos in clientes.items()])

        with open("salas.csv","w", newline="") as archivo:
            grabador = csv.writer(archivo)
            grabador.writerow(("KEY","NOMBRE SALA","CAPACIDAD"))
            grabador.writerows([(clave, datos[0],datos[1]) for clave, datos in salas.items()])

        break
    else:
        print("OPCION DEL MENU NO DISPONIBLE.")