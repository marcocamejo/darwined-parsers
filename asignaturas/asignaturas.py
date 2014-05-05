#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
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

    #Cargar parametros recibidos
    settings, args = parser.parse_args(argv)

    #Cargar datos de archivo de configuracion
    cfg = ConfigParser.ConfigParser()
    cfg.read(["asignaturas.cfg"])

    if not (cfg.has_option("indices", "bi") and cfg.has_option("indices", "bf")
        and cfg.has_option("indices", "ti") and cfg.has_option("indices", "tf")
        and cfg.has_option("indices", "pi") and cfg.has_option("indices", "pf")
    ):
        parser.error("El archivo de configuracion esta incompleto.")

    #Indices de datos basicos
    settings.bi = cfg.get("indices", "bi")
    settings.bf = cfg.get("indices", "bf")
    #Indices de parte teorica
    settings.ti = cfg.get("indices", "ti")
    settings.tf = cfg.get("indices", "tf")
    #Indices de parte practica
    settings.pi = cfg.get("indices", "pi")
    settings.pf = cfg.get("indices", "pf")

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


    #Normalizar argumentos
    settings.semanas = normalize_num(settings.semanas)
    settings.bi  = normalize_num(settings.bi)
    settings.bf  = normalize_num(settings.bf)

    settings.ti  = normalize_num(settings.ti)
    settings.tf  = normalize_num(settings.tf)

    settings.pi = normalize_num(settings.pi)
    settings.pf = normalize_num(settings.pf)

    return settings, args

def leer_atributos(nombre_archivo):
    """ Leer archivo de atributos de sala"""

    attr_book = xlrd.open_workbook(nombre_archivo, encoding_override='LATIN1')

    diccionario = {}
    #Recorrer las hojas
    for s in attr_book.sheets():
        for row in range(1,s.nrows): #omitir encabezado
            clave = s.cell(row,1).value.encode('utf-8').strip()
            valor = s.cell(row,0).value.encode('utf-8').strip()
            diccionario[clave] = valor
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
        for current_row in range(3, num_rows):

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            registro_actual = Registro()
            registro_actual.set_datos_basicos(fila_actual[settings.bi:settings.bf+1])
            registro_actual.set_parte_teorica(fila_actual[settings.ti:settings.tf+1], settings.semanas)
            registro_actual.set_parte_practica(fila_actual[settings.pi:settings.pf+1], settings.semanas)

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
    pf = 17
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
    pi = 17
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

    #Bloque 6: Optativo
    font6       = xlwt.Font()
    font6.bold  = True
    font6.color = 'color white'
    pattern6 = xlwt.Pattern()
    pattern6.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern6.pattern_fore_colour = xlwt.Style.colour_map['yellow']
    style6         = xlwt.XFStyle()
    style6.pattern = pattern6
    style6.font    = font6
    pi = pf
    pf = pi + 1
    escribir_arreglo(arreglo[pi:pf], hoja, fila, pi, style6)


def procesar_planilla(planilla, hoja_salida, bloques, atributos, output_row, settings, optativo):
    """ Procesa una planilla de plan de estudio o de optativos"""

    #Leer nombres de hojas de la planilla
    hojas_planilla = planilla.sheet_names()

    #Procesamiento de la planilla de plan de estudio
    for worksheet_name in hojas_planilla:
        worksheet = planilla.sheet_by_name(worksheet_name)

        num_rows = worksheet.nrows
        num_cols = worksheet.ncols

        print('Hoja: {}'.format(worksheet_name))
        print('--Filas: {}, Columnas: {}'.format(num_rows, num_cols))

        #se leen filas de hoja (omitiendo las primeras 3 filas de encabezado)
        for current_row in range(3, num_rows):

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            registro_actual = Registro()
            registro_actual.set_datos_basicos(fila_actual[settings.bi:settings.bf+1])
            registro_actual.set_parte_teorica(fila_actual[settings.ti:settings.tf+1], settings.semanas)
            registro_actual.set_parte_practica(fila_actual[settings.pi:settings.pf+1], settings.semanas)

            filas_salida = procesar_fila(registro_actual, settings, atributos, bloques, optativo)
            # if len(filas_salida) > 0 and optativo == 1:
            #     print 1
            

            #Puede existir una o 2 asignaturas (teoria y/o practica)
            for rw in filas_salida:
                escribir_arreglo(rw.export_list(), hoja_salida, output_row)
                output_row = output_row + 1

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
        print settings.optativos
        planilla_optativos = xlrd.open_workbook(settings.optativos, encoding_override='LATIN1')
        output_row = procesar_planilla(planilla_optativos, hoja_salida, bloques, atributos, output_row, settings, 1)
    
    #Guardar archivo de salida
    salida.save(settings.salida)

    return 0

if __name__ == '__main__':
    status = main()
    sys.exit(status)

