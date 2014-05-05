"""Funciones Utiles
   :author: Felipe Urra
   :version: 1.0

   Funciones utiles: normalizacion de cadenas y de enteros
""" 

def normalize_str(value):

    if value is None:
        value = ''
    elif type(value) is float or type(value) is int:
        value = str(int(value))
    elif type(value) is unicode or type(value) is str:
        value = value.strip()

    return value

def normalize_num(value):
    try:
        if value is None or value == '':
            value = 0
        elif type(value) is str or type(value) is float or type(value) is unicode:
            value = int(value)
    except ValueError:
        value = -1

    return value

