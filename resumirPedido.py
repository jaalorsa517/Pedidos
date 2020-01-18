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


if __name__ == "__main__":
    main()