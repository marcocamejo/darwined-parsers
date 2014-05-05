
***********************************************************
Esquema
***********************************************************

 *************			*****************
 * Atributos * --------------->	*		*
 *************			*		*
				*		*
*************************	*		*
*			*	*		*
*  *******************	*	*		*	*******************
*  * Plan de estudio *	*	*asignaturas.py * --->	* Asignaturas.xls *
*  *******************	*	*		*	*******************
*			* ---->	*		*
*  *******************	*	*		*
*  *    Optativos    *	*	*		*
*  *******************	*	*		*
*			*	*		*
*************************	*****************

Atributos: hoja con todos los posibles valores de tipo de
sala, definidas por su código y su nombre

Plan de estudio y Optativos: ambos poseen el mismo formato
tal como se describe a continuación.


***********************************************************
Descripción de la planilla de entrada
***********************************************************

Está formado por 3 bloques de datos:
1) Datos basicos (descripcion de la asignatura)
2) Parte Teórica
3) Parte Práctica

Tanto el bloque de parte teórica como el bloque de parte
práctica, están divididos en 3 sub-bloques:
A) Información Básica: Tipo Sala, Capacidad, Nro sesiones
B) Información de Bloque: Franja Horaria, Días de semana
C) Información de Semanas: Una columna por semana


***********************************************************
Descripción del archivo de salida
***********************************************************

Está formado por 5 bloques de datos:
1) Descripción de asignatura
2) Configuración
3) Atributos de bloque
4) Atributos de sala
5) Semanas


***********************************************************
Notas Varias
***********************************************************

- No se está  validando la existencia de  los archivos  de
  entrada.

- Las comparaciones de códigos de tipo de sala se realizan
  siempre en mayúsculas. Por tanto "Sala Silla" es igual a
  "SALA SILLA".

- Se asume que un curso  tiene parte teorica o práctica si
  la columna respectiva "Tipo SALA"  no está vacía y tiene
  un valor que aparece en el archivo de atributos.
  Por ejemplo, en los archivos de entrada algunas columnas
  tienen "Sin sala", cuyo código no aparece en el  archivo
  de atributos, por lo tanto, esa fila es descartada.


***********************************************************
Argumentos recibidos
***********************************************************
    [-i ó --input]      Archivo de entrada
    [-a ó --attributes] Archivo de atributos
    [-o ó --output]     Archivo de salida
    [-s ó --sede]       Sede
    [-e ó --escuela]    Escuela
    [-r ó --regimen]    Regimen
    [-w ó --semanas]    Semanas

***********************************************************
Ejemplo de uso
***********************************************************

python excel_parser.py -i alameda.xlsx -o matriz.xls
   -a alameda_atributos.xls -s "Sede A" -e "Escuela B"
   -r "Regimen C" -w 18



