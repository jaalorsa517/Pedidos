import click
from openpyxl import load_workbook


@click.command()
@click.argument('xlsx')
@click.option('--sheet',
              prompt=True,
              default='Hoja1',
              help='Nombre de la hoja a trabajar')
def main(xlsx, sheet):
    """
    Script que resumirá los pedidos.
    Se requiere un archivo de excel, ocupando las 3 primeras columnas con la información 
    de la siguiente estructura:
    Nombre-Cantidad-Unidad.
    Donde Nombre es un String, Cantidad un float de 1 punto y Unidad es un String
    """
    _workbook = load_workbook(xlsx)
    _sheet = _workbook[sheet]
    _list_productos = _getProductos(_sheet)


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


if __name__ == "__main__":
    main()