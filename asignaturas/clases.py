from utils import normalize_str, normalize_num
from string import upper, split, strip

class CursoParte:
    """ CursoParte representa la parte teorica o practica de un Registro."""
    def __init__(self):

        self.tipo_sala        = []
        self.capacidad        = 0
        self.sesiones         = 0
        self.franja_horaria   = ''
        self.restriccion_dias = [0]*6
        self.modulos_semana   = []
        self.semanas          = 0

    def set_curso(self, datos_basicos, datos_semanas, semanas):

        tipos_sala = normalize_str(datos_basicos[0])
        tipos_sala = tipos_sala.split('/')
        
        self.tipo_sala = []

        for tipo_sala in tipos_sala:
            tipo_sala = tipo_sala.strip().encode('utf-8').upper()
            self.tipo_sala.append(tipo_sala)

        self.capacidad        = normalize_num(datos_basicos[1])
        self.sesiones         = normalize_num(datos_basicos[2])
        self.franja_horaria   = normalize_str(datos_basicos[3])
        self.semanas          = semanas

        #Asignar restricciones de dias de semana (0: lunes - 5: sabado)
        for i in range(0,6):
            self.restriccion_dias[i] = normalize_num(datos_basicos[i+4])

        #Asignar modulos por semanas
        for semana in datos_semanas:
            self.modulos_semana.append(normalize_num(semana))

    def get_bloque(self, jornada):
        bloque = ''
        franjas  = ['D-MAN', 'D-TAR', 'V-SEM', 'V-SAB']

        dias = ''
        for i in range(0,len(self.restriccion_dias)):
            if self.restriccion_dias[i] == 1:
                dias = dias + normalize_str(i+1)

        #Caso 1: Franja no vacia y valida
        if len(self.franja_horaria) > 0 and self.franja_horaria in franjas:
 
            #Si franja corresponde con jornada (Ej: 'V-SEM' y 'V')
            # --> Estudiar dias
            if jornada.upper() == self.franja_horaria[0].upper():

                if len(dias) == 0:
                    bloque = self.franja_horaria
                elif len(dias) > 0:
                    #Si franja vespertina sem y no aparece sabado(6)   --> OK
                    if self.franja_horaria == 'V-SEM' and '6' not in dias:
                        bloque = self.franja_horaria + '-' + dias
                    #Si franja vespertina sab y solo aparece sabado(6) --> OK
                    if self.franja_horaria == 'V-SAB' and dias == '6':
                        bloque = self.franja_horaria + '-' + dias
                    #Si franja diurna (manana o tarde) y no aparece sabado(6) --> OK
                    if (self.franja_horaria == 'D-MAN' or self.franja_horaria == 'D-TAR') and '6' not in dias:
                        bloque = self.franja_horaria + '-' + dias

        #Caso 2: Franja Vacia
        else:
            if len(dias) > 0:
                if (jornada == 'V') or (jornada == 'D' and '6' not in dias):
                    bloque = jornada + '-' + dias
            #elif len(dias) == 0:
            #    bloque = jornada

        return bloque.encode('utf-8')

    def is_valid(self, atributos):
        """ Define si una ParteCurso es valida o no.

            Un curso es valido si tiene al menos una sala asignada y esa sala
            se encuentra dentro de la lista de salas posibles.
        """

        if sum(self.modulos_semana) <= 0:
            return False

        if self.sesiones <= 0:
            return False

        if len(self.tipo_sala) > 0:
            lista_atributos = atributos.items()
            lista_atributos.sort(key= lambda tup: tup[1])

            for tipo_sala in self.tipo_sala:
                for clave, valor in lista_atributos:
                        if clave.upper() == tipo_sala.upper():
                            return True
        return False


class Registro:
    """ Registro representa una fila de la planilla de entrada.

        Cada fila de entrada contiene la parte teorica y practica de un curso,
        por lo tanto un objeto tipo Registro tiene una serie de datos basicos,
        un objeto tipo CursoParte que  representa la parte teorica y otro  que
        representa la parte practica.
    """
    def __init__(self):

        self.curriculum = ''
        self.jornada    = ''
        self.nivel      = 0
        self.siglas     = ''
        self.asignatura = ''
        self.creditos   = 0
        self.horas      = 0

        self.parte_teorica  = CursoParte()
        self.parte_practica = CursoParte()

    def set_datos_basicos(self, entrada):
        self.curriculum = normalize_str(entrada[0])
        self.jornada    = normalize_str(entrada[1])
        self.nivel      = normalize_num(entrada[2])
        self.siglas     = normalize_str(entrada[3])
        self.asignatura = normalize_str(entrada[4])
        self.creditos   = normalize_num(entrada[5])
        self.horas      = normalize_num(entrada[6])

    def set_parte_teorica(self, entrada, semanas):
        i_semanas = len(entrada)-semanas
        datos_basicos = entrada[:i_semanas]
        datos_semanas = entrada[i_semanas:]
        self.parte_teorica.set_curso(datos_basicos, datos_semanas, semanas)

    def set_parte_practica(self, entrada, semanas):
        i_semanas = len(entrada)-semanas
        datos_basicos = entrada[:i_semanas]
        datos_semanas = entrada[i_semanas:]
        self.parte_practica.set_curso(datos_basicos, datos_semanas, semanas)

    def tiene_parte_teorica(self, atributos):
        return self.parte_teorica.is_valid(atributos)

    def tiene_parte_practica(self, atributos):
        return self.parte_practica.is_valid(atributos)

    def procesar_bloques(self, atributos):
        bloque_teorico  = ''
        bloque_practico = ''

        if self.parte_teorica.is_valid(atributos):
            bloque_teorico = self.parte_teorica.get_bloque(self.jornada)
        if self.parte_practica.is_valid(atributos):
            bloque_practico = self.parte_practica.get_bloque(self.jornada)

        return bloque_teorico, bloque_practico




class Asignatura:
    """ Asignatura representa una fila de la planilla de salida.

        Un objeto de clase asignatura,  representa la parte teorica o
        practica de un registro del archivo de entrada, con todos sus
        atributos base.
    """
    def __init__(self):

        self.sede         = ''
        self.escuela      = ''
        self.carrera      = ''
        self.plan_estudio = ''
        self.regimen      = ''
        self.jornada      = ''
        self.codigo       = ''
        self.asignatura   = ''
        self.nivel        = ''
        self.categoria    = ''

        self.num_bloques    = 0
        self.num_sesiones   = 0
        self.num_horas      = 0
        self.estandar       = 0
        self.num_profesores = 1
        self.anual          = 0
        self.online         = 0
        self.optativo       = 0

        self.bloques = []
        self.salas   = []
        self.semanas = []

    def set_from_registro(self, registro, bloques, atributos, settings, categoria = 'TEO', optativo = 0):
        self.sede         = settings.sede
        self.escuela      = settings.escuela
        self.carrera      = registro.curriculum
        self.plan_estudio = registro.curriculum
        self.regimen      = settings.regimen
        self.jornada      = registro.jornada
        self.codigo       = registro.siglas
        self.asignatura   = registro.asignatura
        self.nivel        = registro.nivel
        self.categoria    = categoria

        if categoria == 'TEO':
            curso = registro.parte_teorica
        else:
            curso = registro.parte_practica

        #Determinar el numero de bloques por semana:
        #num_bloques es primer elemento > 0 del arreglo de modulos por semana
        num_bloques = 0
        for semana in curso.modulos_semana:
            if semana > 0:
                num_bloques = semana
                break

        self.num_bloques    = num_bloques
        self.num_sesiones   = curso.sesiones
        self.num_horas      = 0
        self.estandar       = curso.capacidad
        self.num_profesores = 1
        self.anual          = 0
        self.online         = 0
        self.optativo       = optativo

        #Ajustar arreglo de atributos de bloque
        bloque_curso = curso.get_bloque(registro.jornada)
        for bloque in bloques:
            if bloque_curso == bloque:
                self.bloques.append(1)
            else:
                self.bloques.append(0)

        #Ajustar arreglo de atributos de sala
        lista_atributos = atributos.items()
        lista_atributos.sort(key= lambda tup: tup[1])
        for clave, valor in lista_atributos:
            found = False
            for pr, tipo_sala in enumerate(curso.tipo_sala):
                if clave.upper() == tipo_sala.upper():
                    found = True
                    prioridad = pr+1


            if found is True:
                self.salas.append(prioridad)
            else:
                self.salas.append(0)


        #Ajustar arreglo de semanas
        #Para cada elemento distinto de 0 el arreglo de modulos por semana
        #guardar un 1 en la columna respectiva de salida.
        #guardar un 0 en caso contrario
        for semana in curso.modulos_semana:
            if semana != 0:
                self.semanas.append(1)
            else:
                self.semanas.append(0)

    def export_list(self):
        salida = []
        salida.append(self.sede)
        salida.append(self.escuela)
        salida.append(self.carrera)
        salida.append(self.plan_estudio)
        salida.append(self.regimen)
        salida.append(self.jornada)
        salida.append(self.codigo)
        salida.append(self.asignatura)
        salida.append(self.nivel)
        salida.append(self.categoria)
        salida.append(self.num_bloques)
        salida.append(self.num_sesiones)
        salida.append(self.num_horas)
        salida.append(self.estandar)
        salida.append(self.num_profesores)
        salida.append(self.anual)
        salida.append(self.online)

        for sala in self.salas:
            salida.append(sala)

        for bloque in self.bloques:
            salida.append(bloque)

        for semana in self.semanas:
            salida.append(semana)

        salida.append(self.optativo)

        return salida

    def export_header_list(self, bloques, salas, semanas):
        salida = []
        salida.append('SEDE')
        salida.append('ESCUELA')
        salida.append('CARRERA')
        salida.append('PLAN ESTUDIO')
        salida.append('REGIMEN')
        salida.append('JORNADA')
        salida.append('CODIGO ASIGNATURA')
        salida.append('NOMBRE ASIGNATURA')
        salida.append('NIVEL')
        salida.append('CATEGORIA')
        salida.append('NUM BLOQUES')
        salida.append('NUM SESIONES')
        salida.append('NUM HORAS')
        salida.append('ESTANDAR')
        salida.append('NUM PROFESORES')
        salida.append('ANUAL')
        salida.append('ONLINE')
        salida.append('OPTATIVO')

        for clave, valor in salas:
            salida.append(valor)

        for bloque in bloques:
            salida.append(bloque)

        for i in range(0,semanas):
            salida.append('S'+str(i+1))

        return salida


