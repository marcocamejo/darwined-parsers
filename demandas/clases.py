import hashlib
from utils import normalize_str, normalize_num

class Cursable:
    """ Cursable representa una fila del archivo de entrada."""
    def __init__(self, fila):

        self.sede             = normalize_str(fila[0])
        self.st_objid         = fila[1]
        self.rut_estudiante   = fila[2]
        self.zplan_estudios   = fila[3]
        self.zasignatura_cod  = fila[4]
        self.otype            = fila[5]
        self.zsede            = fila[6]
        self.zjornada         = fila[7]
        self.zmodo_calculo    = fila[8]
        self.peryr            = fila[9]
        self.perid            = fila[10]
        self.curriculum       = fila[11]
        self.sl_text          = fila[12]
        self.codigo_asignatura= fila[13]
        self.o_stext          = fila[14]
        self.stgbeg           = fila[15]
        self.wexp_modcpmin    = fila[16]
        self.wexp_modcpopt    = fila[17]
        self.wexp_modcpamx    = fila[18]
        self.wexp_modcpunt    = fila[19]
        self.estado_st        = fila[20]
        self.estados_desc     = fila[21]
        self.uno              = fila[22]
        self.suma_optativo    = fila[23]
        self.peryr_proy       = fila[24]
        self.perid_proy       = fila[25]
        self.zinstitucion     = fila[26]


class Demanda:
    """ Demanda representa una fila del archivo de salida."""
    def __init__(self):

        self.sede              = ''
        self.escuela           = ''
        self.carrera           = ''
        self.plan_estudio      = ''
        self.regimen           = ''
        self.jornada           = ''
        self.asignatura        = ''
        self.nivel             = ''
        self.alumnos           = 0
        self.alumnos_estimados = 0
        self.lista_cruzada     = ''

        #Variable artificial para almacenar las claves CARRERA-JORNADA-CODIGO
        self.llave             = ''

    def set_from_cursable(self, cursable, settings):

        self.sede              = normalize_str(settings.sede)
        self.escuela           = normalize_str(settings.escuela)
        self.carrera           = normalize_str(cursable.curriculum)
        self.plan_estudio      = normalize_str(cursable.curriculum)
        self.regimen           = normalize_str(settings.regimen)
        self.jornada           = normalize_str(cursable.zjornada)
        self.asignatura        = normalize_str(cursable.codigo_asignatura)
        self.nivel             = ''
        self.llave             = normalize_str(cursable.curriculum+'-'+cursable.zjornada+'-'+cursable.codigo_asignatura)
        self.alumnos           = 0
        self.alumnos_estimados = 0
        self.lista_cruzada     = ''

    def export_list(self):
        salida = []
        salida.append(self.sede)
        salida.append(self.escuela)
        salida.append(self.carrera)
        salida.append(self.plan_estudio)
        salida.append(self.regimen)
        salida.append(self.jornada)
        salida.append(self.asignatura)
        salida.append(self.nivel)
        salida.append(self.alumnos)
        salida.append(self.alumnos_estimados)
        salida.append(self.lista_cruzada)

        return salida

    def export_header_list(self):
        salida = []
        salida.append('SEDE')
        salida.append('ESCUELA')
        salida.append('CARRERA')
        salida.append('PLAN DE ESTUDIO')
        salida.append('REGIMEN')
        salida.append('JORNADA')
        salida.append('CODIGO ASIGNATURA')
        salida.append('NIVEL')
        salida.append('ALUMNOS')
        salida.append('ALUMNOS ESTIMADOS')
        salida.append('LISTA CRUZADA')

        return salida






class AliasAsignatura:
    """ AliasAsignatura representa un Alias de una Asignatura a planificar."""
    def __init__(self, fila):
        self.id_asignatura      = normalize_str(fila[0])
        self.codigo_asignatura  = normalize_str(fila[1])
        self.asignatura         = normalize_str(fila[2])

class Asignatura:
    """ Asignatura representa una Asignatura a planificar.

        La clase Asignatura guarda los datos de la asignatura
        (filas A-E del archivo de referencia), asi como una lista
        con los posibles alias
    """

    def __init__(self, fila):
        self.codigo_escuela      = normalize_str(fila[0])
        self.escuela             = normalize_str(fila[1])
        self.id                  = normalize_str(fila[2])
        self.codigo              = normalize_str(fila[3])
        self.asignatura          = normalize_str(fila[4])

        #Arreglo para almacenar los posibles alias de la asignatura
        self.alias = []

        #Atributo virtual para almacenar el numero actual de alias
        self.nalias = 0

        #Calculo de la llave cruzada de la asignatura
        #(depende exclusivamente del codigo base)
        m = hashlib.md5()
        m.update(self.codigo)
        self.llave_cruzada = self.codigo+'-'+m.hexdigest()

        #Primer Alias: codigo en fila G (6)
        if normalize_str(fila[6]) != '':
            aliasasignatura = AliasAsignatura(fila[5:8])
            self.alias.append(aliasasignatura)
            self.nalias = self.nalias + 1

        #Segundo Alias: codigo en fila J (9)
        if normalize_str(fila[9]) != '':
            aliasasignatura = AliasAsignatura(fila[8:11])
            self.alias.append(aliasasignatura)
            self.nalias = self.nalias + 1
            
        #Tercer Alias: codigo en fila M (12)
        if normalize_str(fila[12]) != '':
            aliasasignatura = AliasAsignatura(fila[11:14])
            self.alias.append(aliasasignatura)
            self.nalias = self.nalias + 1

    def has_alias(self, codigo):
        """ Verifica si el codigo recibido se encuentra entre los alias

            Si el alias se encuentra, devuelve la posicion dentro del arreglo
            de alias, en caso contrario, devuelve -1
        """
        
        for i in range(self.nalias):
            if self.alias[i].codigo_asignatura == codigo:
                return True
        return False

    def get_lista_cruzada(self, codigo):
        """ Devuelve el valor de lista cruzada segun el codigo recibido

            Si el codigo se encuentra, devuelve el valor de la lista cruzada
            en caso contrario, devuelve ''
        """
        if self.codigo == codigo or self.has_alias(codigo):
            return self.llave_cruzada
        return ''


    def add_alias(self, alias_list):
        """ Agrega la lista de alias recibida a la lista interna de alias

            Se valida que el alias no se encuentre ya en la lista, para evitar
            repetidos.
        """

        for i in range(len(alias_list)):
            if not self.has_alias(alias_list[i].codigo_asignatura):
                self.alias.append(alias_list[i])
                self.nalias = self.nalias + 1



