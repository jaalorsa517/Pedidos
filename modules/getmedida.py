# -*- coding: utf-8 -*-

import click
from openpyxl import load_workbook
import modules.utility as org


def medida(xlsx):
    _workbook = load_workbook(xlsx)

    _sheet = org.getSheet(_workbook)

    #Obtener Dic {row:producto}
    _list_productos = org.getPedido(_sheet)

    for key, product in _list_productos.items():
        _sheet.cell(row=key, column=3, value=org.get_unidad(product))

    _workbook.save(xlsx)
