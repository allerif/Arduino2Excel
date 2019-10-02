# -*- coding: utf-8 -*-
"""
Programa: Comunicação Arduino e Excel com Python
Autor: Izabelle Aller

"""

import serial
import time
import xlwt

arquivo= xlwt.Workbook()
tabela= arquivo.add_sheet(u'Tabela 1')
tabela.write(0,0, u'Dia')
tabela.write(0,1, u'Hora')
tabela.write(0,2, u'Temperatura')

coluna = 1

porta = serial.Serial('COM5', baudrate = 9600, timeout = 3)
time.sleep(3)

while 1:
 arduinoData = porta.readline().decode() 
 Dia = arduinoData[0:10]
 Hora = arduinoData[11:19]
 Temperatura = arduinoData[20:25]
 temperaturaVirgula = Temperatura.replace('.',',')
 
 tabela.write(coluna,0, u'%s' %Dia)
 tabela.write(coluna,1, u'%s' %Hora)
 tabela.write(coluna,2, u'%s' %temperaturaVirgula )
 coluna = coluna + 1 
 tabela.save('Arduino2Excel.xls')