
***********************************************************
Notas Varias
***********************************************************

- Este script realiza 2 pre-procesamientos:
  1) Calcula el numero total de claves CARRERA-JORN-CODIGO
     existentes en la planilla de entrada.

  2) Lee del archivo de referencia todos los posibles
     codigos de asignaturas y los almacena en una lista
     para realizar comparaciones al momento de procesar
     la planilla de entrada.

- Si una fila de la planilla tiene un código que no se
  encuentra en el archivo de referencias, la fila
  correspondiente en el archivo de salida tendra la
  columna LISTA_CRUZADA vacía.

- No se valida la existencia de lOs archivos de entrada.


***********************************************************
Argumentos recibidos
***********************************************************
    [-i ó --input]       Archivo de entrada
    [-f ó --referencias] Archivo de referencia
    [-o ó --output]      Archivo de salida
    [-s ó --sede]        Sede
    [-e ó --escuela]     Escuela
    [-r ó --regimen]     Regimen

***********************************************************
Ejemplo de uso
***********************************************************

python proyecciones.py -i alameda.xlsx -o proyecciones.xls
  -f referencias.xslx -s "Sede A" -e "Escuela B"
  -r "Regimen C"

