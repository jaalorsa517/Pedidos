# -*- coding: utf-8 -*-

#datetime.datetime(2020,1,10,0,0)
#cell.font.b:bool negrita
#cell.font.name:'Arial'
#cell.font.sz:10
#cell.border.bottom.border_style:'hair'
#cell.border.bottom.style:'hair'

from modules.principal import main
'''
Archivo Inventario
1. Buscar en la primera fila la fecha deseada !
2. Obtenida la columna, guardar todas las filas donde "Ped" sea != 0 or Null !
3.Guardar en un dict {id:,nombre:,tel:,email:,pedido:{}}
4.Repetir los pasos anteriores en cada hoja de calculo, guardandolos dentro una list

Archivo Pedido
1. Crear una nueva hoja con el nombre de la fecha
2. Hacer encabezado y descargar la list dict
3. Aplicar formato
'''

#PUNTO DE ENTRADA
if __name__ == "__main__":
    main()