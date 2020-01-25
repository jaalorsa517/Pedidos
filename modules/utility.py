# -*- encoding: utf-8 -*-

import click


def getSheet(book):
    '''
    Función que devuelve una hoja de calculo seleccionada por el usuario.
    @param book: Libro xlsx
    @return: Hoja de calculo
    '''
    #INICIO DEL MENU
    op = ''
    i = 0
    for sheet in book.get_sheet_names():
        i += 1
        op += '{}.{}\n'.format(str(i), sheet)

    while True:
        try:
            r = click.prompt('Ingrese la opción deseada:\n{}'.format(op))
            sheet = book[book.get_sheet_names()[int(r) - 1]]
            break
        except ValueError:
            continue
        except IndexError:
            continue
    #FIN DEL MENU
    return sheet


def getPedido(sheet):
    '''
    Función que obtiene todos los pedidos a un diccionario
    @param sheet: Hoja de calculo.
    @return dic: Diccionario {row:string}.
    '''
    list = {}
    #Ciclo que recorre desde la fila ini_row y la columna col
    for col in sheet.iter_cols(min_row=1, max_col=1):
        for cell in col:
            #Evalua si contiene dato
            if (not (cell.value == None)):
                #Evalua si el contendio esta en negrita (encabezado)
                if (not cell.font.b):
                    list[cell.row] = cell.value
    return list


def getDataColumn(sheet,ini_row=3,col=1):
    '''
    Función que obtiene todos los pedidos a un diccionario
    @param sheet: Hoja de calculo.
    @paramDefault ini_row: Primera fila de la iteracion.
    @paramDefault col: Columna a la cual iterar.
    @return dic: Diccionario {row:string}.
    '''
    list = {}
    #Ciclo que recorre desde la fila ini_row y la columna col
    for col in sheet.iter_cols(min_row=ini_row, max_col=col):
        for cell in col:
            #Evalua si contiene dato
            if (not (cell.value == None)):
                #Evalua si el contendio esta en negrita (encabezado)
                if (not cell.font.b):
                    list[cell.row] = cell.value
    return list

def get_unidad(product):
    '''
    Función encargada de poner la unidad de medida de productos.
    @param product: Nombre del producto
    @return: String con la medida correspondiente
    '''
    PAQUETE = ('Avena*6', 'Frescolanta*6', 'Kumis*6', 'Tampico*6','Tampico*125*30' , 'Yagur*6')
    SIXPACK = ('Leche deslactosada*1100ml*6', 'Leche entera*1100ml*6',
               'Leche montefrio*900*6', 'Leche prolinco*900*6',
               'Leche prolinco deslactosada*6', 'Leche semidescremada*1100*6')
    CAJA = ('Leche Ricura CAJA', 'Leche polvo prolinco*780 CAJA')
    RISTRA = ('Leche en polvo RISTRA','Arequipe RISTRA')

    for p in PAQUETE:
        if (p == product):
            return 'Paquete'

    for p in SIXPACK:
        if (p == product):
            return 'Sixpack'

    for p in CAJA:
        if (p == product):
            return 'Caja'

    for p in RISTRA:
        if (p == product):
            return 'Ristra'

    return 'Unidad'