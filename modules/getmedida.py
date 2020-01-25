# -*- coding: utf-8 -*-

import click
from modules.principal import main
from openpyxl import load_workbook
import modules.utility as org


@main.command()
@click.argument('xlsx')
def getMedida(xlsx):
    """
    Completará la columna de medida.
    Se requiere un archivo de excel, ocupando las 3 primeras columnas con la información 
    de la siguiente estructura:
    Nombre-Cantidad-Unidad.
    Donde Nombre es un String, Cantidad un float de 1 punto y Unidad es un String
    """
    _workbook = load_workbook(xlsx)

    _sheet = org.getSheet(_workbook)

    #Obtener Dic {row:producto}
    _list_productos = org.getDataColumn(_sheet)

    for key, product in _list_productos.items():
        _sheet.cell(row=key, column=3, value=org.get_unidad(product))

    _workbook.save(xlsx)
