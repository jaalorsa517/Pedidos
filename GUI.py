# -*- coding:utf-8 -*-

import datetime
import os
# import gi
# gi.require_version('Gtk', '3.0')
# from gi.repository import Gtk
import tkinter as tk
from modules import orderall, getmedida
from openpyxl import Workbook, load_workbook

hoy = datetime.date.today()
selfpath = os.getcwd()
path = None
fecha = None

win = tk.Tk()
win.title("Pedido")
win.geometry('350x200')

# calendar = builder.get_object("calendar")
# calendar.day = hoy.day
# calendar.month = hoy.month
# calendar.year = hoy.year

# dlInfo = builder.get_object('dlInfo')

# def day_select(day):
#     global fecha
#     fecha = datetime.date(day.get_date().year,
#                           day.get_date().month + 1,
#                           day.get_date().day)

# def on_btPath_file_set(file):
#     global path
#     path = file.get_file().get_path()

# def on_btProcesar_clicked(boton):
#     global path
#     global fecha
#     global selfpath

#     if (path is not None and fecha is not None):

#         if ('.xlsx' in path):

#             _selfpath = selfpath + '/pedido.xlsx'

#             if (os.path.isfile(_selfpath)):
#                 os.remove(_selfpath)

#             pedido = Workbook()
#             pedido.save(_selfpath)

#             orderall.order(path, _selfpath, fecha)
#             getmedida.medida(_selfpath, True)

#             pedido = load_workbook(_selfpath)
#             pedido.remove(pedido.worksheets[0])
#             pedido.save(_selfpath)

#             dlInfo.set_property('text', 'Finalizó el proceso')
#             dlInfo.set_property('secondary_text',
#                                 'El proceso terminó con éxito')

#         else:
#             dlInfo.set_property('text', 'Archivo erróneo')
#             dlInfo.set_property('secondary_text',
#                                 'Por favor seleccione un archivo XLSX')

#     else:
#         dlInfo.set_property('text', 'Falta seleccionar la fecha o el archivo.')
#         dlInfo.set_property('secondary_text',
#                             'Por favor seleccione una fecha y/o un archivo')

#     dlInfo.run()

# def on_btDialog_clicked(boton):
#     dlInfo.hide()

# handlers = {
#     "on_main_destroy": Gtk.main_quit,
#     "on_btPath_file_set": on_btPath_file_set,
#     "on_btProcesar_clicked": on_btProcesar_clicked,
#     "day_select": day_select,
#     "on_btDialog_clicked": on_btDialog_clicked
# }
# builder.connect_signals(handlers)

if __name__ == "__main__":
    win.mainloop()