# Programa Pedido

## Síntesis

Este programa está hecho con python 3.
En la rama CLI, es un programa de consola.
En la rama GUI-TK es un programa de interfaz con tkinter.

## Descripción

Este programa procesa una hoja de excel específica, obtiene unos datos y estructura dichos datos en otro archivo excel.

## Obtener el ejecutable con Pyinstaller

Para obtener el ejecutable, abra una terminal, ubiquese en la raíz del proyecto y ejecute:
pyinstaller -p modules/ --hidden-import babel.numbers -Fw GUI.py
