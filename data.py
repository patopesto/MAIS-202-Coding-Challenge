#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys, csv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Series, Reference
from operator import itemgetter

reload(sys)
sys.setdefaultencoding('utf-8')


def data(fileToReadFrom, chosenFieldnames, worksheetToWrite):

    csv.register_dialect('semicolons', delimiter=';')
    fieldnamesreader = csv.DictReader(fileToReadFrom)
    fieldnameslist = fieldnamesreader.fieldnames
    #print fieldnameslist
    L=[]

    for index in range(len(chosenFieldnames)):
       L.append(get_column_letter(fieldnameslist.index(chosenFieldnames[index]) + 1))

    for item, index in enumerate(L):
        print item, index
    print "\n"

    reader = csv.reader(fileToReadFrom)

    #boucle permettant d'inscrire les colonnes dont j'ai besoin dans l'excel (chosenfieldnames au dessus)
    for row_index, row in enumerate(reader):
        for column_index, cell in enumerate(row):
            column_letter = get_column_letter((column_index + 1))
            if column_letter in L:
                columnletter = get_column_letter((L.index(column_letter)) +1)
                if cell in (None, " "):
                    worksheetToWrite["%s%d" % (columnletter, (row_index + 2))].value = None
                else:
                    try:
                        worksheetToWrite["%s%d" % (columnletter, (row_index + 2))].value = cell.lower()
                    except:
                        worksheetToWrite["%s%d" % (columnletter, (row_index + 2))].value = cell.replace(u"\u0003","").lower()
                        
    for index in range(len(chosenFieldnames)):
        worksheetToWrite.cell(row = 1, column = (index+1), value=(chosenFieldnames[index]))
    

def average(worksheetToRead):

    purposes = []
    values = []

    for row_index, index in enumerate(worksheetToRead.iter_rows()):
        purpose = worksheetToRead["%s%d" % ('A', (row_index + 2))].value
        if purpose is None:
            break
        value = float(worksheetToRead["%s%d" % ('B', (row_index + 2))].value)
        print purpose, value
        if purpose in purposes:
            i = purposes.index(purpose)
            values[i].append(value)
        else:
            purposes.append(purpose)
            a = [value]
            values.append(a)

    averages = []

    for l in values:
        numElements = len(l)
        result = 0
        for element in l:
            result += element
        averages.append(result/numElements)
    print purposes
    print averages

    return purposes, averages



def graph(purposes, averages, worksheet):

    worksheet.cell(row = 1, column = 1, value = "purposes")
    worksheet.cell(row = 1, column = 2, value = "avg_rate")

    for index, item in enumerate(purposes):
        worksheet.cell(row = (index+2), column = 1, value = item)
        worksheet.cell(row = (index+2), column = 2, value = averages[index])


    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 13
    chart1.height = 20
    chart1.width = 20
    chart1.title = "Bar Chart"
    chart1.y_axis.title = 'Average Rates'
    chart1.x_axis.title = 'Purpose'

    data = Reference(worksheet, min_col=2, max_col=2, min_row=2, max_row=13)
    cats = Reference(worksheet, min_col=1, max_col=1, min_row=2, max_row=13)
    chart1.add_data(data)
    chart1.set_categories(cats)
    chart1.shape = 4
    worksheet.add_chart(chart1, "E4")


wb = Workbook()
ws = wb.active

file = open('data.csv', 'rb')
chosenFieldnames = ['purpose','int_rate']
data(file, chosenFieldnames, ws)

purposes, averages = average(ws)

ws_chart = wb.create_sheet(title = "chart")
graph(purposes, averages, ws_chart)
wb.remove(ws)
wb.save("result.xlsx")


