# -*- coding: utf-8 -*-
import click
from datetime import datetime
from openpyxl import load_workbook
import modules.utility as ut


def order(inventario, pedido):
    _workbook = load_workbook(inventario)
    _fecha = click.prompt("Ingrese fecha #Mes(3)")
    _mes_dict = {
        'ene': 1,
        'feb': 2,
        'mar': 3,
        'abr': 4,
        'may': 5,
        'jun': 6,
        'jul': 7,
        'ago': 8,
        'sep': 9,
        'oct': 10,
        'nov': 11,
        'dic': 12
    }
    _mes = 0
    for key, val in _mes_dict.items():
        if (key == _fecha[2::]):
            _mes = val
            break
    _fechaDate = datetime(2020, _mes, int(_fecha[0:2]))
    inv=[]
    cliente={}
    for hoja in _workbook.sheetnames:
        'Recorrer las columnas B1:W1'
        for fila in _workbook[hoja].iter_rows(min_col=2, max_row=1):
            for cell in fila:
                if (cell != None):
                    if (cell == _fecha or cell == _fechaDate):
                        #cell.column
                        #cell.row
                        #cell.comment
                        cliente['id']
                        cliente['nombre']
                        cliente['telefono']
                        cliente['email']
                        #cliente['pedido']=ut.getDataColumn(hoja,3,cell.column+1)
            inv.append(cliente)
