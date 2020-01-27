# -*- coding: utf-8 -*-
import click
from modules.getmedida import medida
from modules.getresume import resume
from modules.orderall import order


@click.group()
def main():
    """Programa que organizará los pedidos"""
    pass


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
    medida(xlsx)


@main.command()
@click.argument('xlsx')
def getResume(xlsx):
    """
    Resumirá los pedidos.
    Se requiere un archivo de excel, ocupando las 3 primeras columnas con la información 
    de la siguiente estructura:
    Nombre-Cantidad-Unidad.
    Donde Nombre es un String, Cantidad un float de 1 punto y Unidad es un String
    """
    resume(xlsx)


@main.command()
@click.argument('inventario')
@click.argument('pedido')
def orderAll(inventario, pedido):
    '''
    Método que organizará todo el pedido.
    :param inventario: archivo xlsx con el inventario
    :param pedido: archivo xlsx donde se hace el pedido
    '''
    order(inventario, pedido)