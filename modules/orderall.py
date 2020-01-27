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
    inv = []
    for hoja in _workbook.sheetnames:
        a1=_workbook[hoja][1][0]
        cliente = {}
        if (a1.comment is not None):
            cliente = ut.datosInComment(a1.comment.text)
            cliente['nom']=a1.value
        else:
            click.echo("No existe comentario en la hoja {}.".format(hoja))
            cliente['nom'] = a1.value
        'Recorrer las columnas B1:W1'
        for fila in _workbook[hoja].iter_rows(min_col=2, max_row=1):
            for cell in fila:
                if (cell.value != None):
                    if (cell.value == _fecha or cell.value == _fechaDate):
                        #cell.column
                        #cell.row
                        #cell.comment
                        #_workbook[hoja][1][0].comment.text
                        cliente['pedido'] = ut.getDataColumn(_workbook[hoja],
                                                             ini_row=3,
                                                             col=cell.column +
                                                             1)
        if ('pedido' in cliente):
            #[{'nom':'','pedido':{}}]
            inv.append(cliente)

    #Se carga el xlsx pedido
    _workbook = load_workbook(pedido)
    #Se crea la nueva hoja
    _sheet = _workbook.create_sheet(_fecha)

    _row=2
    for i in inv:
        if('id'in i and 'tel'in i and 'email'in i):
            ut.cabecera(_sheet,_row, i['nom'],i['id'],i['tel'],i['email'])
        else:
            ut.cabecera(_sheet,_row, i['nom'])

        _row+=5
        ut.setPedido(_sheet,_row,i['pedido'])
        _row+= len(i['pedido']+1)
    
    _workbook.save(pedido)
        
