# -*- coding: utf-8 -*-
import click
from openpyxl import load_workbook


#METODO
@click.command()
@click.argument('xlsx')
def main(xlsx):
    """
    Script que resumirá los pedidos.
    Se requiere un archivo de excel, ocupando las 3 primeras columnas con la información 
    de la siguiente estructura:
    Nombre-Cantidad-Unidad.
    Donde Nombre es un String, Cantidad un float de 1 punto y Unidad es un String
    """
    _workbook = load_workbook(xlsx)

    #INICIO DEL MENU
    op = ''
    i = 0
    for sheet in _workbook.get_sheet_names():
        i += 1
        op += '{}.{}\n'.format(str(i), sheet)

    while True:
        try:
            r = click.prompt('Ingrese la opción deseada:\n{}'.format(op))
            _sheet = _workbook[_workbook.get_sheet_names()[int(r) - 1]]
            break
        except ValueError:
            continue
        except IndexError:
            continue
    #FIN DEL MENU

    #Obtener Dic {row:producto}
    _list_productos = _getProductos(_sheet)

    #Obtener conjunto
    _con_produtos = set(_list_productos.values())

    #Escribir titulo y encabezados
    row_max = sorted(_list_productos)[-1]
    cont = row_max + 2
    _sheet.cell(row=cont, column=1, value='RESUMEN')
    cont += 2
    _sheet.cell(row=cont, column=1, value='Nombre')
    _sheet.cell(row=cont, column=2, value='Cantidad')
    _sheet.cell(row=cont, column=3, value='Unidad')

    #Escribir el resumen
    for product in _con_produtos:
        cont += 1
        _sheet.cell(row=cont, column=1, value=product)
        _sheet.cell(row=cont,
                    column=2,
                    value='=SUMIF(A5:A{rm},A{c},B5:B{rm})'.format(
                        rm=row_max, c=cont))

    _workbook.save(xlsx)


#FUNCION
def _getProductos(sheet):
    '''
    Función que obtiene todos los pedidos a un diccionario
    @param sheet: Hoja de calculo
    @return dic: Diccionario con la fila y el producto pedido
    '''
    list = {}
    for col in sheet.iter_cols(min_row=1, max_col=1):
        for cell in col:
            if (not (cell.value == None)):
                if (not cell.font.b):
                    list[cell.row] = cell.value
    return list


#PUNTO DE ENTRADA
if __name__ == "__main__":
    main()