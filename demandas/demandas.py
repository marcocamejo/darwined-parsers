import os
import sys
import optparse
import ConfigParser
from string import upper
import xlrd
import xlwt
from clases import Cursable, Demanda, Asignatura
from utils import normalize_str, normalize_num

#para cargar codificacion utf-8 en python
reload(sys)
sys.setdefaultencoding('utf-8')

def procesar_argumentos(argv):
    """Procesa los argumentos o parametros de entrada

    Esta funcion genera los parametros de configuracion en
    funcion de los argumentos recibidos por linea de comando.

    :param argv: lista de argumentos
    :return: settings, args
    """

    if argv is None:
        argv = sys.argv[1:]

    parser = optparse.OptionParser()

    #Definir opciones (argumentos) esperados
    parser.add_option("-l", "--log", dest="log",
        help="Ruta del archivo de log")
    parser.add_option("-i","--input", dest="entrada",
        help="Ruta del archivo de entrada")
    parser.add_option("-f", "--referencias", dest="referencias",
        help="Ruta del archivo de referencias")
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

    # Se toma el primer argumento (no parametro) como el directorio base
    if len(args) > 0:
        settings.dir = args[0]
    else:
        settings.dir = ''

    #Cargar datos de archivo de configuracion
    if settings.config:
        config_file = settings.dir+settings.config
    else:
        config_file = settings.dir+"demandas.cfg"

    cfg = ConfigParser.ConfigParser()
    cfg.read([config_file])

    #Verificar argumentos
    if not settings.log:
        if cfg.has_option("archivos", "log"):
            settings.log = cfg.get("archivos", "log")
        else:
            parser.error('No se ha especificado un archivo de entrada')
    if not settings.entrada:
        if cfg.has_option("archivos", "entrada"):
            settings.entrada = cfg.get("archivos", "entrada")
        else:
            parser.error('No se ha especificado un archivo de entrada')
    if not settings.referencias:
        if cfg.has_option("archivos", "referencias"):
            settings.referencias = cfg.get("archivos", "referencias")
        else:
            parser.error('No se ha especificado un archivo de referencias')
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

    settings.semanas = normalize_num(settings.semanas)

    settings.log       = settings.dir+settings.log
    settings.entrada   = settings.dir+settings.entrada
    settings.salida    = settings.dir+settings.salida
    settings.referencias = settings.dir+settings.referencias

    return settings, args


def leer_llaves(planilla, hojas_planilla, settings):
    """Pre-Procesamiento de la planilla: obtener llaves

       Esta funcion lee todas las llaves de la planilla y las devuelve en
       un arreglo (ordenado). El arreglo contiene claves repetidas, a fin
       de contar el numero total de apariciones de cada clave
    """
    llaves = []
    for worksheet_name in hojas_planilla:
        worksheet = planilla.sheet_by_name(worksheet_name)

        num_rows = worksheet.nrows
        num_cols = worksheet.ncols

        #se leen filas de hoja
        for current_row in range(1, num_rows): #omitir 1 columna de encabezado

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            cursable_actual = Cursable(fila_actual)

            #Filtrar filas con otype = ST
            if cursable_actual.otype.upper() != 'ST':
                demanda = Demanda()
                demanda.set_from_cursable(cursable_actual, settings)
                llaves.append(demanda.llave)

    llaves.sort()
    return llaves

def leer_referencias(nombre_archivo):
    """ Leer archivo de asignaturas de referencias"""

    ref_book = xlrd.open_workbook(nombre_archivo, encoding_override='LATIN1')

    asignaturas = []

    #Recorrer las hojas
    for s in ref_book.sheets():
        for row in range(1,s.nrows): #omitir encabezado
            fila_actual = []

            #Recorrer las columnas
            for current_col in range(0,s.ncols):
                fila_actual.append(normalize_str(s.cell(row, current_col).value))

            #Asignatura temporal (con todos sus alias)
            asignatura_actual = Asignatura(fila_actual)
  
            #Buscar la asignatura en las asignaturas ya revisadas
            asignatura_encontrada = False

            nasignaturas = len(asignaturas)
            for i in range(0,nasignaturas):
                #Si la asignatura ya fue revisada --> agregar posibles alias
                if asignatura_actual.codigo == asignaturas[i].codigo:
                    asignaturas[i].add_alias(asignatura_actual.alias)
                    asignatura_encontrada = True
                    break

            #Si la asignatura no ha sido revisada --> agregar  asignatura
            #                                         (y todos sus alias)
            if not asignatura_encontrada:
                asignaturas.append(asignatura_actual)

    return asignaturas


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


def escribir_encabezado(arreglo, hoja, fila):
    """Escribe la linea de encabezado del archivo de salida

    :param arreglo: Arreglo que se quiere escribir en la hoja
    :param hoja: hoja en donde se va a esribir
    :param fila: fila en la que se escribe
    :return: datos_basicos
    """

    #datos basicos
    font1       = xlwt.Font()
    font1.bold  = True
    font1.color = 'color white'
    pattern1 = xlwt.Pattern()
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern1.pattern_fore_colour = xlwt.Style.colour_map['blue']
    style1         = xlwt.XFStyle()
    style1.pattern = pattern1
    style1.font    = font1
    escribir_arreglo(arreglo[0:8], hoja, fila, 0, style1)

    #datos alumnos
    font2       = xlwt.Font()
    font2.bold  = True
    font2.color = 'color white'
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = xlwt.Style.colour_map['green']
    style2         = xlwt.XFStyle()
    style2.pattern = pattern2
    style2.font    = font2
    escribir_arreglo(arreglo[8:10], hoja, fila, 8, style2)

    #lista cruzada
    font3       = xlwt.Font()
    font3.bold  = True
    font3.color = 'color white'
    pattern3 = xlwt.Pattern()
    pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern3.pattern_fore_colour = xlwt.Style.colour_map['purple_ega']
    style3         = xlwt.XFStyle()
    style3.pattern = pattern3
    style3.font    = font3
    escribir_arreglo(arreglo[10:11], hoja, fila, 10, style3)


def main(argv = None):

    settings, args = procesar_argumentos(argv)

    # Crear los paths de los archivos de salida y log si no existen
    if not os.path.exists(os.path.dirname(settings.log)):
        os.makedirs(os.path.dirname(settings.log))
    if not os.path.exists(os.path.dirname(settings.salida)):
        os.makedirs(os.path.dirname(settings.salida))

    #Leer planilla
    planilla = xlrd.open_workbook(settings.entrada, encoding_override='LATIN1')

    #Leer hojas de la planilla
    hojas_planilla = planilla.sheet_names()

    #Crear archivo salida
    salida = xlwt.Workbook(encoding='LATIN-1')

    #Cear hoja de archivo de salida
    hoja_salida = salida.add_sheet('Sheet1')

    #Contador para las filas del archivo de salida
    output_row = 0

    #Lista de llaves
    llaves = leer_llaves(planilla, hojas_planilla, settings)
    #Conjunto de llaves (no repetidas) a fines de optimizar la busqueda
    set_llaves = set(llaves)
    #Lista de llaves insertadas
    llaves_insertadas = []

    #Obtener lista de posibles asignaturas y sus alias
    asignaturas = leer_referencias(settings.referencias)

    #Procesamiento de la planilla
    for worksheet_name in hojas_planilla:
        worksheet = planilla.sheet_by_name(worksheet_name)

        num_rows = worksheet.nrows
        num_cols = worksheet.ncols

        print('Hoja: {}'.format(worksheet_name))
        print('--Filas: {}, Columnas: {}'.format(num_rows, num_cols))

        #se leen filas de hoja
        for current_row in range(1, num_rows): #omitir 1 columna de encabezado

            fila_actual = []

            #Recorrer las columnas
            for current_col in range(num_cols):
                fila_actual.append(normalize_str(worksheet.cell(current_row, current_col).value))

            cursable_actual = Cursable(fila_actual)

            #Filtrar filas con otype = ST
            if cursable_actual.otype.upper() != 'ST':
                demanda = Demanda()
                demanda.set_from_cursable(cursable_actual, settings)

                #primera fila -> agregar encabezado
                if(output_row == 0):
                    escribir_encabezado(demanda.export_header_list(), hoja_salida, output_row)
                    output_row = output_row + 1

                #Insertar el registro en el archivo de salida si aun no ha sido insertado
                #un registro con la misma clave (CARRERA-JORNADA-ASIGNATURA)
                if demanda.llave in set_llaves:
                    if demanda.llave not in llaves_insertadas:

                        #Calcular el numero de alumnos
                        demanda.alumnos = llaves.count(demanda.llave)

                        #Calcular la llave cruzada de acuerdo al valor de
                        #demanda.asignatura
                        for asignatura in asignaturas:
                            lista_cruzada = asignatura.get_lista_cruzada(demanda.asignatura)
                            if lista_cruzada != '':
                                demanda.lista_cruzada = lista_cruzada
                                break

                        escribir_arreglo(demanda.export_list(), hoja_salida, output_row)
                        llaves_insertadas.append(demanda.llave)
                        output_row = output_row + 1

    #Guardar archivo
    salida.save(settings.salida)

    return 0

if __name__ == '__main__':
    status = main()
    sys.exit(status)

