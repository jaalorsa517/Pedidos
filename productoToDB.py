"""
App que copiará datos de un xlxs a una base de datos.
"""
from openpyxl import load_workbook
import click


@click.group()
def main():
    """Programa para la manipulación de archivos xlsx"""
    pass


@main.command()
@click.argument('workbook')
@click.option('--sheetsrc',
              default='Hoja1',
              prompt=True,
              help='Nombre de la hoja origen')
@click.option('--range',
              default='A2',
              prompt=True,
              help='Rango de celdas a copiar')
# @click.option('--sheetdest', prompt=True, help='Nombre de la hoja destino')
# @click.option('--rangedest', prompt=True, help='Rango de celdas a copiar')
def copyRange(workbook, sheetsrc, range):
    """Copia un rango de celdas a otro"""
    w = load_workbook(workbook)
    sheet = w[sheetsrc]
    sheet[range] = 'Hello'
    w.save(workbook)


if __name__ == "__main__":
    main()