#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import time
import optparse
import ConfigParser
import xlrd
import xlwt
from clases import CursoParte, Registro, Asignatura
from utils import normalize_str, normalize_num

#para cargar codificacion utf-8 en python
reload(sys)
sys.setdefaultencoding('utf-8')

def procesar_argumentos(argv):
    """Procesa los argumentos o parametros de entrada

    Esta  funcion genera los  parametros de configuracion en funcion de  los
    argumentos recibidos por linea de comando y el archivo de configuracion.

    :param argv: lista de argumentos
    :return: settings, args
    """

    if argv is None:
        argv = sys.argv[1:]

    parser = optparse.OptionParser()

    #Definir opciones (argumentos) esperados
    parser.add_option("-i","--input", dest="entrada",
        help="Ruta del archivo de entrada")
    parser.add_option("-p","--optativos", dest="optativos",
        help="Ruta del archivo de optativos")
    parser.add_option("-a","--attributes", dest="atributos",
        help="Ruta del archivo de atributos")
    parser.add_option("-o", "--output", dest="salida",
        help="Ruta del archivo de salida")
    parser.add_option("-s", "--sede", dest="sede",
        help="Sede")
    parser.add_option("-e", "--escuela", dest="escuela",
        help="Escuela")
    parser.add_option("-r", "--regimen", dest="regimen",
        help="Regimen")
    parser.add_option("-w", "--semanas", dest="semanas",
        help="Semanas planificadas por curso")
    parser.add_option("-c", "--config", dest="config",
        help="Archivo de configuracion")

    #Cargar parametros recibidos
    settings, args = parser.parse_args(argv)

    #Cargar datos de archivo de configuracion
    if settings.config:
        config_file = settings.config
    else:
        config_file = "asignaturas.cfg"

    cfg = ConfigParser.ConfigParser()
    cfg.read(["asignaturas.cfg"])

    #Verificar argumentos
    if not settings.entrada:
        if cfg.has_option("archivos", "entrada"):
            settings.entrada = cfg.get("archivos", "entrada")
        else:
            parser.error('No se ha especificado un archivo de entrada')
    if not settings.optativos:
        if cfg.has_option("archivos", "optativos"):
            settings.optativos = cfg.get("archivos", "optativos")
        else:
            settings.optativos = ''
    if not settings.atributos:
        if cfg.has_option("archivos", "atributos"):
            settings.atributos = cfg.get("archivos", "atributos")
        else:
            parser.error('No se ha especificado un archivo de atributos')
    if not settings.salida:
        if cfg.has_option("archivos", "salida"):
            settings.salida = cfg.get("archivos", "salida")
        else:
            parser.error('No se ha especificado un archivo de salida')
    if not settings.sede:
        if cfg.has_option("basicos", "sede"):
            settings.sede = cfg.get("basicos", "sede")
        else:
            settings.sede = 'SEDE'
    if not settings.escuela:
        if cfg.has_option("basicos", "escuela"):
            settings.escuela = cfg.get("basicos", "escuela")
        else:
            settings.escuela = 'ESCUELA'
    if not settings.regimen:
        if cfg.has_option("basicos", "regimen"):
            settings.regimen = cfg.get("basicos", "regimen")
        else:
            settings.regimen = 'REGIMEN'
    if not settings.semanas:
        if cfg.has_option("basicos", "semanas"):
            settings.semanas = cfg.get("basicos", "semanas")
        else:
            settings.semanas = 18

    #Fila inicial de lectura de datos (= numero de filas a omitir)
    if cfg.has_option("basicos", "ini_filas"):
        settings.ini_filas = cfg.get("basicos", "ini_filas")
    else:
        settings.ini_filas = 3 #por omision se omiten 3 filas de encabezado

    #Columna inicial de lectura de datos (= numero de columnas a omitir)
    if cfg.has_option("basicos", "ini_columnas"):
        settings.ini_columnas = cfg.get("basicos", "ini_columnas")
    else:
        settings.ini_columnas = 0 #por omision no se omiten columnas

    #Maximo numero de filas vacias consecutivas permitidas
    if cfg.has_option("basicos", "max_filas_vacias"):
        settings.max_filas_vacias = cfg.get("basicos", "max_filas_vacias")
    else:
        settings.max_filas_vacias = 20

    #Normalizar argumentos
    settings.semanas          = normalize_num(settings.semanas)
    settings.ini_filas        = normalize_num(settings.ini_filas)
    settings.ini_columnas     = normalize_num(settings.ini_columnas)
    settings.max_filas_vacias = normalize_num(settings.max_filas_vacias)

    return settings, args

def leer_atributos(nombre_archivo):
    """ Leer archivo de atributos de sala"""

    attr_book = xlrd.open_workbook(nombre_archivo, encoding_override='LATIN1')

    log = open('log', 'a')

    diccionario = {}
    #Recorrer las hojas
    for s in attr_book.sheets():
        for row in range(1,s.nrows): #omitir encabezado
            if s.cell(row,0).value.encode('utf-8').strip().upper() == 'SALA':
                clave = s.cell(row,2).value.encode('utf-8').strip()
                valor = s.cell(row,1).value.encode('utf-8').strip()
                diccionario[clave] = valor
            else:
                log.write(nombre_archivo+": Fila "+str(row+1)+" no es tipo sala")

    log.close()

    return diccionario

def leer_bloques(planilla, settings, atributos):
    """ Genera la lista de todos los posibles bloques de la planilla """

    #Arreglo para almacenar los posibles bloque ('D-MAN','V-14',..)
    bloques = []

    #Leer nombres de hojas de la planilla
    hojas_planilla = planilla.sheet_names()

    for worksheet_name in hojas_planilla:
        worksheet = planilla.sheet_by_name(worksheet_name)

        num_rows = worksheet.nrows
        num_cols = worksheet.ncols

        #se leen filas de hoja
        for current_row in range(settings.ini_filas, num_rows):

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            registro_actual = Registro()
            registro_actual.set_from_list(fila_actual, settings)

            bloque_teorico, bloque_practico = registro_actual.procesar_bloques(atributos)
            if len(bloque_teorico) > 0:
                bloques.append(bloque_teorico)
            if len(bloque_practico) > 0:
                bloques.append(bloque_practico)

    #Unicuificar y ordenar el arreglo de bloques
    bloques = set(bloques)
    bloques = list(bloques)
    bloques.sort()

    return bloques

def procesar_fila(registro, settings, atributos, bloques, optativo = 0):
    """Procesa un registro del archivo de entrada

    Esta funcion procesa un registro del archivo de entrada, devolviendo uno o
    dos objetos tipo Asignatura, dependiendo de si existe parte teorica y/o
    practica.

    :param registro: objeto tipo Registro
    :param settings: arreglo con parametros de configuracion
    :param atributos: arreglo con los valores del archivo de atributos
    """

    salida = []

    #Si tiene parte teorica
    if(registro.tiene_parte_teorica(atributos)):
        rowT = Asignatura()
        rowT.set_from_registro(registro, bloques, atributos, settings, 'TEO', optativo)
        salida.append(rowT)

    #Si tiene parte practica
    if(registro.tiene_parte_practica(atributos)):
        rowT = Asignatura()
        rowT.set_from_registro(registro, bloques, atributos, settings, 'PRA', optativo)
        salida.append(rowT)

    return salida

def escribir_arreglo(arreglo, hoja, fila, columna_inicial = 0, estilo = None):
    """Escribe un arreglo a la hoja de calculo recibida

    :param arreglo: Arreglo que se quiere escribir en la hoja
    :param hoja: Hoja de calculo
    :param fila: fila en la que se escribe
    :param columna_inicial: columna en la que se comienza a escribir
    :param estilo: XFStyle para escribir la celda
    :return: settings, args
    """

    if estilo is None:
        for i in range(0, len(arreglo)):
            hoja.write(fila, columna_inicial+i, arreglo[i])
    else:
        for i in range(0, len(arreglo)):
            hoja.write(fila, columna_inicial+i, arreglo[i], estilo)

def escribir_encabezado(arreglo, bloques, salas, semanas, hoja, fila):
    """Escribe la linea de encabezado del archivo de salida

    Aqui se configuran los indices y colores de cada bloque de columnas

    :param arreglo: Lista que se quiere escribir en la hoja
    :param bloques: Lista de nombres de bloques
    :param salas: Lista de nombres de salas
    :param semanas: Numero de semanas
    :param hoja: hoja en donde se va a esribir
    :param fila: fila en la que se escribe
    :return: datos_basicos
    """

    #Bloque 1: Descripcion de asignatura
    font1       = xlwt.Font()
    font1.bold  = True
    font1.color = 'color white'
    pattern1 = xlwt.Pattern()
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['blue']
    style1         = xlwt.XFStyle()
    style1.pattern = pattern1
    style1.font    = font1
    pi = 0
    pf = 10
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style1)

    #Bloque 2: Configuracion
    font2       = xlwt.Font()
    font2.bold  = True
    font2.color = 'color white'
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = xlwt.Style.colour_map['red']
    style2         = xlwt.XFStyle()
    style2.pattern = pattern2
    style2.font    = font2
    pi = 10
    pf = 18
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style2)

    #Bloque 3: Atributos de sala
    font4       = xlwt.Font()
    font4.bold  = True
    font4.color = 'color white'
    pattern4 = xlwt.Pattern()
    pattern4.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern4.pattern_fore_colour = xlwt.Style.colour_map['purple_ega']
    style4         = xlwt.XFStyle()
    style4.pattern = pattern4
    style4.font    = font4
    pi = 18
    pf = pi + len(salas)
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style4)

    #Bloque 4: Atributos de bloque
    font3       = xlwt.Font()
    font3.bold  = True
    font3.color = 'color white'
    pattern3 = xlwt.Pattern()
    pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern3.pattern_fore_colour = xlwt.Style.colour_map['turquoise']
    style3         = xlwt.XFStyle()
    style3.pattern = pattern3
    style3.font    = font3
    pi = pf
    pf = pi + len(bloques)
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style3)

    #Bloque 5: Semanas
    font5       = xlwt.Font()
    font5.bold  = True
    font5.color = 'color white'
    pattern5 = xlwt.Pattern()
    pattern5.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern5.pattern_fore_colour = xlwt.Style.colour_map['olive_ega']
    style5         = xlwt.XFStyle()
    style5.pattern = pattern5
    style5.font    = font5
    pi = pf
    pf = pi + semanas
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style5)

def procesar_planilla(planilla, hoja_salida, bloques, atributos, output_row, settings, optativo):
    """ Procesa una planilla de plan de estudio o de optativos"""

    #Registros procesados
    registros_procesados = []

    #Leer nombres de hojas de la planilla
    hojas_planilla = planilla.sheet_names()

    #Abrir archivo de log
    log = open('log', 'a')
    if optativo == 0:
        log.write('\nPlanilla de Plan de Estudio\n')
    else:
        log.write('\nPlanilla de Optativos\n')

    #Procesamiento de la planilla de plan de estudio
    for worksheet_name in hojas_planilla:
        worksheet = planilla.sheet_by_name(worksheet_name)

        num_rows = worksheet.nrows
        num_cols = worksheet.ncols

        #numero de registros totales
        num_registros = 0
        #numero de registros procesados
        num_procesados = 0

        #numero de registros vacios (consecutivamente)
        num_vacios = 0

        log.write('Hoja: '+worksheet_name+'\n')
        log.write('--Filas: '+str(num_rows)+', Columnas: '+str(num_cols)+'\n')

        #se leen filas de hoja (omitiendo las primeras 3 filas de encabezado)
        for current_row in range(settings.ini_filas, num_rows):

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            registro_actual = Registro()
            registro_actual.set_from_list(fila_actual, settings)

            num_registros = num_registros+1

            #Omitir fila si no trae campos obligatorios: curriculum, jornada, nivel, siglas o asignatura
            if registro_actual.curriculum == '' or registro_actual.jornada == '' or registro_actual.nivel == 0 \
               or registro_actual.siglas == '' or registro_actual.asignatura == '':
                log.write("Fila "+str(current_row+1)+" omitida por no tener campos obligatorios\n")
                num_vacios = num_vacios + 1

                #Si se alcanza el maximo de filas vacias consecutivas permitidas: terminar hoja
                if num_vacios == settings.max_filas_vacias:
                    log.write('Se alcanzo el maximo permitido de filas vacias consecutivamente (' \
                        +str(settings.max_filas_vacias)+') en la hoja '+worksheet_name+'\n')
                    break

                continue
            else: 
                #Omitir fila si ya se proceso una fila con igual curriculum-jornada-asignatura
                num_vacios = 0
                clave = registro_actual.curriculum + '-' + registro_actual.jornada + '-' + registro_actual.siglas
                if clave in registros_procesados:
                    log.write('Fila '+str(current_row+1)+' omitida por duplicidad de Curriculum-Jornada-Asignatura ('+clave+')\n')
                    continue

            filas_salida = procesar_fila(registro_actual, settings, atributos, bloques, optativo)
            if len(filas_salida) == 0:
                log.write('Fila '+str(current_row+1)+' sin parte teorica ni practica validas\n')
            else:
                registros_procesados.append(clave)
                num_procesados = num_procesados + 1

            #Puede existir una o 2 asignaturas (teoria y/o practica)
            for rw in filas_salida:
                escribir_arreglo(rw.export_list(), hoja_salida, output_row)
                output_row = output_row + 1

    log.write('Total filas: '+str(num_registros)+'\n')
    porcentaje_procesados = round(100.0*num_procesados/(num_registros), 2)
    log.write('Total filas procesadas: '+str(num_procesados)+' ('+str(porcentaje_procesados)+'%)\n')

    log.close()

    return output_row


def main(argv = None):

    settings, args = procesar_argumentos(argv)

    #Leer planilla
    planilla = xlrd.open_workbook(settings.entrada, encoding_override='LATIN1')

    #Leer archivo de atributos
    atributos = leer_atributos(settings.atributos)
    lista_atributos = atributos.items()
    lista_atributos.sort(key= lambda tup: tup[1])

    #Crear archivo salida
    salida = xlwt.Workbook(encoding='LATIN-1')

    #Crear hoja de archivo de salida
    hoja_salida = salida.add_sheet('Sheet1')

    #Crear archivo de log
    log = open("log", "w")
    log.write(time.strftime('Archivo generado el: %d/%m/%Y - %H:%M:%S\n'))
    log.write('salida: '+settings.salida+'\n')
    log.close()

    #Preprocesamiento de la planilla: leer bloques
    bloques = leer_bloques(planilla, settings, atributos)

    #Escribir encabezado
    encabezado = Asignatura().export_header_list(bloques, lista_atributos, settings.semanas)
    escribir_encabezado(encabezado, bloques, atributos, settings.semanas, hoja_salida, 0)

    #Contador para las filas del archivo de salida
    output_row = 1

    #Procesamiento de planilla de plan de estudio
    output_row = procesar_planilla(planilla, hoja_salida, bloques, atributos, output_row, settings, 0)

    #Procesamiento de planilla de optativos (si existe)
    if settings.optativos != '':
        planilla_optativos = xlrd.open_workbook(settings.optativos, encoding_override='LATIN1')
        output_row = procesar_planilla(planilla_optativos, hoja_salida, bloques, atributos, output_row, settings, 1)
    
    #Guardar archivo de salida
    salida.save(settings.salida)

    return 0

if __name__ == '__main__':
    status = main()
    sys.exit(status)

