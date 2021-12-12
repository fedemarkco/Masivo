# coding=latin1
from time import gmtime, strftime, sleep
from win32com.client import Dispatch

import PySimpleGUI as sg
import mysql.connector
import subprocess
import xlsxwriter
import logging
import base64
import re
import os


def connectMysql():
  con = mysql.connector.connect(
    host = "host",
    user = "user",
    passwd = "passwd",
    database = "database"
  )

  return con

def queryMysql():
  global cells

  try:
    con = connectMysql()
    cur = con.cursor()

    cur.execute("""select hostname, 
          ipaddress, 
          group_concat(distinct siteid order by categoria separator ', ') 
          from inventario_sites 
          where sitioOld in ("+cells+") or 
          siteid in ("+cells+") and 
          hostname <> '' and 
          ipaddress <> '' group by hostname order by hostname"""
        )

    table = """
      <table style="border-collapse: collapse; font-family: Calibri;" 
        border="1" cellspacing="0" cellpadding="0">
        <tr style="background-color: #FFC75F;">
          <td style="padding: 5px;">Cabecera IPRAN</td>
          <td style="padding: 5px;">IP</td>
          <td style="padding: 5px;">Celdas</td>
        <tr>"""

    rows = cur.fetchall()

    if cur.rowcount == 0:
      return None

    for c in rows:
      table += '<tr style="background-color: #DFF1F8;">\
                  <td style="padding: 5px;">'+c[0]+'</td>\
                  <td style="padding: 5px;">'+c[1]+'</td>\
                  <td style="padding: 5px;">'+c[2]+'</td>\
                </tr>'

    table += '</table>'

    return table
  except:
    return None

def getGreeting():
  import datetime

  now = datetime.datetime.now()

  if now.hour >= 20 or now.hour <= 6:
    text = "Buenas noches estimados,"
  if now.hour >= 13 and now.hour <= 19:
    text = "Buenas tardes estimados,"
  if now.hour >= 7 and now.hour <= 12:
    text = "Buenos d&iacute;as estimados,"

  return text

def createSubjectAlarma():
  subject = "Alarma de "

  return subject

def createBodyAlarma():
  from xlsx2html import xlsx2html

  global filename
  global CountCells

  countCellsTotal = CountCells[0] + CountCells[1] + CountCells[2]
  out_stream = xlsx2html(os.path.abspath(filename))
  out_stream.seek(0)
  output = out_stream.read()

  table = re.search(
    '<table(\s+\S+){1,}.*?<\/td>\s<\/tr>\s<\/table>', 
    output).group(0)
  table = re.sub(
    r'<table style="border-collapse: collapse" \
      border="0" cellspacing="0" cellpadding="0"><colgroup>', 
    '<table  style="border-collapse: collapse; font-family: Calibri;" \
      border="0" cellspacing="0" cellpadding="0"><colgroup>', 
    table
  )
  table = re.sub(r'<col\s.*', '', table)
  table = re.sub(r'<colgroup>\s+<\/colgroup>', '', table)
  table = re.sub(r'font-size: 11.0px;', 'font-size: 15.0px;', table)
  table = re.sub(
    r'background-color: #FFC75F;', 
    'padding-left: 4px;padding-right: 15px;background-color: #FFC75F;', 
    table
  )
  table = re.sub(
    r'background-color: #DFF1F8;', 
    'padding-left: 4px;padding-right: 15px;background-color: #DFF1F8;', 
    table
  )
  if CountCells == 1:
    body = getGreeting() + \
      "<br><br>" + \
      "Se genera incidencia por la siguiente alarma: <br><br>" + \
      table
  else:
    body = getGreeting() + \
      "<br><br>" + \
      "Se genera incidencia por las siguientes alarmas: <br><br>" + \
      table

  return body

def createSubjectMasiva():
  global CountCells

  countCellsTotal = CountCells[0] + CountCells[1] + CountCells[2]
  subject = ""

  if CountCells[0] != 0:
    subject += str(CountCells[0])
    if CountCells[0] == 1:
      subject += " sitio 2G"
    else:
      subject += " sitios 2G"
  if CountCells[1] != 0:
    if subject == "":
      subject += str(CountCells[1])
      if CountCells[1] == 1:
        subject += " sitio 3G"
      else:
        subject += " sitios 3G"
    else:
      if CountCells[1] != 0:
        subject += ", " + str(CountCells[1])
        if CountCells[1] == 1:
          subject += " sitio 3G"
        else:
          subject += " sitios 3G"
      else:
        subject += " y " + str(CountCells[1])
        if CountCells[1] == 1:
          subject += " sitio 3G"
        else:
          subject += " sitios 3G"
  if CountCells[2] != 0:
    if subject == "":
      subject += str(CountCells[2])
      if CountCells[2] == 1:
        subject += " sitio 4G"
      else:
        subject += " sitios 4G"
    else:
      subject += " y " + str(CountCells[2])
      if CountCells[2] == 1:
        subject += " sitio 4G"
      else:
        subject += " sitios 4G"
  if countCellsTotal == 1:
    subject += " fuera de servicio en "
  else:
    subject += " fuera de servicio en "

  return subject

def createBodyMasiva():
  try:
    from xlsx2html import xlsx2html

    global filename
    global cells

    out_stream = xlsx2html(os.path.abspath(filename))
    out_stream.seek(0)
    output = out_stream.read()

    table = re.search(
      '<table(\s+\S+){1,}.*?<\/td>\s<\/tr>\s<\/table>', 
      output).group(0)
    table = re.sub(
      r'<table  style="border-collapse: collapse" \
        border="0" cellspacing="0" cellpadding="0"><colgroup>', 
      '<table  style="border-collapse: collapse; font-family: Calibri;" \
        border="0" cellspacing="0" cellpadding="0"><colgroup>', 
      table
    )
    table = re.sub(r'<col\s.*', '', table)
    table = re.sub(r'<colgroup>\s+<\/colgroup>', '', table)
    table = re.sub(r'font-size: 11.0px;', 'font-size: 15.0px;', table)
    table = re.sub(
      r'background-color: #FFC75F;', 
      'padding-left: 4px;padding-right: 15px;background-color: #FFC75F;', 
      table
    )
    table = re.sub(
      r'background-color: #DFF1F8;', 
      'padding-left: 4px;padding-right: 15px;background-color: #DFF1F8;', 
      table
    )

    tableIPRAN = queryMysql()

    if tableIPRAN is not None:
      body = getGreeting() + \
        "<br><br>" + \
        "Tenemos " + \
        createSubjectMasiva() + \
        "<br><br>Incidencia NOC: <br>Tarea TX: <br><br>" + \
        table + \
        "<br><br><br>" + \
        tableIPRAN
    else:
      body = getGreeting() + \
        "<br><br>" + \
        "Tenemos " + \
        createSubjectMasiva() + \
        "<br><br>Incidencia NOC: <br>Tarea TX: <br><br>" + \
        table

    return body

  except:
    import sys

    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)

#type: 0 corresponde a Masiva y 1 corresonde a Alarmas
def sendEmail(type):
    import win32com.client as win32
    import pythoncom
    import pathlib

    pythoncom.CoInitialize ()

    outlook = win32.Dispatch(
                      'outlook.application', 
                      clsctx=pythoncom.CLSCTX_LOCAL_SERVER
                    )
    mail = outlook.CreateItemFromTemplate(
                      pathlib.Path().absolute().joinpath('firma.oft'))

    if type == 0:
      mail.SentOnBehalfOfName = "NOC@claro.com.ar"
      mail.To = "SoportedeTransmision@claro.com.ar"
      mail.CC = "NOC@claro.com.ar"
      mail.Subject = createSubjectMasiva()
      mail.HTMLbody = "<div style='font-family: Calibri'>" + \
        createBodyMasiva() + \
        "</div>"
    if type == 1:
      mail.SentOnBehalfOfName = "NOC@claro.com.ar"
      mail.To = "BORA@claro.com.ar"
      mail.CC = "NOC@claro.com.ar"
      mail.Subject = createSubjectAlarma()
      mail.HTMLbody = "<div style='font-family: Calibri'>" + \
        createBodyAlarma() + \
        "</div>"
    mail.Display()
    os.remove('firma.oft')

def CountTechnology(tecs):
  global CountCells

  tecnologias = ['2G', '3G', '4G']

  CountTech2G = tecs.count(tecnologias[0])
  CountTech3G = tecs.count(tecnologias[1])
  CountTech4G = tecs.count(tecnologias[2])

  CountCells = [CountTech2G, CountTech3G, CountTech4G]

  return CountTech2G, CountTech3G, CountTech4G

def CreateXlsMasivo(listAllPart1, listAllPart2, tecs, email=False):
  global window
  global filename

  try:
    cabeceras = [
      'CATEGORY', 
      'CELLID', 
      'Original Event Time', 
      'Managed Object', 
      'Additional Text'
    ]
    tecnologias = ['2G', '3G', '4G']
    CountTech2G, CountTech3G, CountTech4G = CountTechnology(tecs)
    Sheet = createSubjectMasiva()[:-4]
    filename = Sheet+strftime(" - %Y-%m-%d %H.%M.%S", gmtime())+'.xlsx'
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Hoja 1')
    formatoCabecera = workbook.add_format(
      {
        'bold': True, 
        'bg_color': '#FFC75F', 
        'border_color': 'black', 
        'border': 1, 
        'text_justlast': True
      }
    )
    formatoNoCabecera = workbook.add_format(
      {
        'bg_color': '#DFF1F8', 
        'border_color': 'black', 
        'border': 1, 
        'text_justlast': True
      }
    )
    row = 0
    col = 0
    pos = 0
    com = 0
    maxLength1 = 0
    maxLength2 = 0
    maxLength3 = 0
    maxLength4 = 0

    while len(tecnologias) != (pos):
      enc = 0
      for data in range(len(tecs)):
        if(tecs[data].upper().find(tecnologias[pos])) != -1:
          if com == 0 and \
            ((tecnologias[pos] == '2G' and \
            CountTech2G >0) or \
            (tecnologias[pos] == '3G' and \
            CountTech3G >0) or \
            (tecnologias[pos] == '4G' and \
            CountTech4G >0)):
            for cabecera in cabeceras:
              worksheet.write(row, col, cabecera, formatoCabecera)
              col +=1
          row +=1
          col = 0
          com = 1
          enc = 1

          maxLength1 = max(
                        len(listAllPart1[data]["cellId"])+3, 
                        maxLength1
                      )
          maxLength2 = max(
                        len(listAllPart2[data]["originalEventTime"])+3, 
                        maxLength2
                      )
          maxLength3 = max(
                        len(listAllPart1[data]["managedObject"])+3, 
                        maxLength3
                      )
          maxLength4 = max(
                        len(listAllPart2[data]["additionalText"])+3, 
                        maxLength4
                      )

          worksheet.write(
                      row, 
                      col, 
                      listAllPart2[data]["category"], 
                      formatoNoCabecera
                    )
          worksheet.set_column(row, col, 15)
          worksheet.write(
                      row, 
                      (col+1), 
                      listAllPart1[data]["cellId"], 
                      formatoNoCabecera
                    )
          worksheet.set_column(row, (col+1), maxLength1)
          worksheet.write(
                      row, 
                      (col+2), 
                      listAllPart2[data]["originalEventTime"], 
                      formatoNoCabecera
                    )
          worksheet.set_column(row, (col+2), maxLength2)
          worksheet.write(
                      row, 
                      (col+3), 
                      listAllPart1[data]["managedObject"], 
                      formatoNoCabecera
                    )
          worksheet.set_column(row, (col+3), maxLength3)
          worksheet.write(
                      row, 
                      (col+4), 
                      listAllPart2[data]["additionalText"], 
                      formatoNoCabecera
                    )
          worksheet.set_column(row, (col+4), maxLength4)

      if enc == 1:
        row +=2
        com = 0
      pos += 1

    workbook.close()

    if email == False:
      window.Hide()
      confirm = sg.Popup(
        "El archivo fue creado correctamente.\n¿Desea abrir el archivo generado?", 
        title="Información", 
        button_type=1
      )
      if confirm == 'Yes':
        os.system('start excel.exe "'+filename+'"')
        window.UnHide()
      if confirm == 'No':
        window.UnHide()
    else:
      sendEmail(0)
  except:
    import sys

    window.Hide()

    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
    confirm = sg.Popup(
      "La operación no pudo ser realizada correctamente.\nIntente nuevamente.", 
      title="Advertencia", 
      button_type=0
    )
    if confirm == 'OK':
      window.UnHide()

def CreateXlsAlarmas(listAllPart1, listAllPart2, tecs, email=False):
  global window
  global filename

  try:
    Sheet = strftime("%Y-%m-%d %H.%M.%S", gmtime())
    CountTech2G, CountTech3G, CountTech4G = CountTechnology(tecs)
    cabeceras = [
      'CATEGORY', 
      'CELLID', 
      'Original Event Time', 
      'Managed Object', 
      'Additional Text'
    ]
    filename = Sheet+'.xlsx'
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Hoja 1')
    formatoCabecera = workbook.add_format(
      {
        'bold': True, 
        'bg_color': '#FFC75F', 
        'border_color': 'black', 
        'border': 1, 
        'text_justlast': True
      }
    )
    formatoNoCabecera = workbook.add_format(
      {
        'bg_color': '#DFF1F8', 
        'border_color': 'black', 
        'border': 1, 
        'text_justlast': True
      }
    )
    row = 0
    col = 0

    for cabecera in cabeceras:
      worksheet.write(row, col, cabecera, formatoCabecera)
      col +=1

    row = 1
    col = 0
    maxLength1 = 0
    maxLength2 = 0
    maxLength3 = 0
    maxLength4 = 0

    for data in range(len(tecs)):

      maxLength1 = max(
                    len(listAllPart1[data]["cellId"])+3, 
                    maxLength1
                  )
      maxLength2 = max(
                    len(listAllPart2[data]["originalEventTime"])+3, 
                    maxLength2
                  )
      maxLength3 = max(
                    len(listAllPart1[data]["managedObject"])+3, 
                    maxLength3
                  )
      maxLength4 = max(
                    len(listAllPart2[data]["additionalText"])+3, 
                    maxLength4
                  )

      worksheet.write(
                  row, 
                  col, 
                  listAllPart2[data]["category"], 
                  formatoNoCabecera
                )
      worksheet.set_column(row, col, 15)
      worksheet.write(
                  row, 
                  (col+1), 
                  listAllPart1[data]["cellId"], 
                  formatoNoCabecera
                )
      worksheet.set_column(row, (col+1), maxLength1)
      worksheet.write(
                  row, 
                  (col+2), 
                  listAllPart2[data]["originalEventTime"], 
                  formatoNoCabecera
                )
      worksheet.set_column(row, (col+2), maxLength2)
      worksheet.write(
                  row, 
                  (col+3), 
                  listAllPart1[data]["managedObject"], 
                  formatoNoCabecera
                )
      worksheet.set_column(row, (col+3), maxLength3)
      worksheet.write(
                  row, 
                  (col+4), 
                  listAllPart2[data]["additionalText"], 
                  formatoNoCabecera
                )
      worksheet.set_column(row, (col+4), maxLength4)

      row +=1
    workbook.close()
    if email == False:
      window.Hide()
      confirm = sg.Popup(
        "El archivo xls fue creado correctamente.\n¿Desea abrir el archivo generado?", 
        title="Información", 
        button_type=1
      )
      if confirm == 'Yes':
        os.system('start excel.exe "'+filename+'"')
        window.UnHide()
      if confirm == 'No':
        window.UnHide()
    else:
      sendEmail(1)
  except:
    import sys

    window.Hide()

    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
    confirm = sg.Popup(
      "La operación no pudo ser realizada correctamente.\nIntente nuevamente.", 
      title="Advertencia", 
      button_type=0
    )
    if confirm == 'OK':
      window.UnHide()

def GetData(s, option):
  global cells

  cellId = ""
  cells = ""
  tec = ""
  category = ""
  managedObject = ""
  originalEventTime = ""
  listAdditionalText = ""
  listAllPart1 = []
  listAllPart2 = []
  listAllPart3= []
  tecs = []

  frags = re.findall("OPERATION_CONTEXT[\s\S]+?CELLID=.*", s)

  for frag in frags:

    cellId = re.search('CELLID=(.*)', frag, re.IGNORECASE)
    if cellId:
      cellId = cellId[1]

    managedObject = re.search('Managed Object=(.*)', frag, re.IGNORECASE)
    if managedObject:
      if "mrbts" in managedObject[1].lower() or "4g" in managedObject[1].lower():
        tec = "4G"
        managedObject = managedObject[1]
      if "rnc" in managedObject[1].lower() or "3g" in managedObject[1].lower():
        tec = "3G"
        managedObject = managedObject[1]
      if "bsc" in managedObject[1].lower() or "2g" in managedObject[1].lower():
        tec = "2G"
        managedObject = managedObject[1]

    if {"cellId": cellId, "tec": tec} not in listAllPart3:
      listAllPart1.append({"cellId": cellId, "managedObject": managedObject})
      listAllPart3.append({"cellId": cellId, "tec": tec})
      cells += "'" + cellId + "', "
      tecs.append(tec)
    else:
      if option == 1 or option == 3:
        continue
      else:
        listAllPart1.append({"cellId": cellId, "managedObject": managedObject})
        listAllPart3.append({"cellId": cellId, "tec": tec})
        cells += "'" + cellId + "', "
        tecs.append(tec)

    additionalText = re.search('Additional Text=(.*)', frag, re.IGNORECASE)
    if additionalText:
      additionalText = additionalText[1]

    originalEventTime = re.search('Original Event Time=(.*)', frag, re.IGNORECASE)
    if originalEventTime:
      originalEventTime = originalEventTime[1]

    category = re.search('Categoria=(.*)', frag, re.IGNORECASE)
    if category:
      category = category[1]

    listAllPart2.append(
      {
        "additionalText": additionalText, 
        "originalEventTime": originalEventTime, 
        "category": category
      }
    )

  cells = cells[:-2]

  if option == 0:
    CreateXlsAlarmas(listAllPart1, listAllPart2, tecs)
  if option == 1:
    CreateXlsMasivo(listAllPart1, listAllPart2, tecs)
  if option == 2:
    CreateXlsAlarmas(listAllPart1, listAllPart2, tecs, True)
  if option == 3:
    CreateXlsMasivo(listAllPart1, listAllPart2, tecs, True)

def ObtenerCeldas(s):
  listCells = []
  res = ''
  cells = re.findall('CELLID=(.*)', s)
  if cells:
    for cell in cells:
      listCells.append(cell)
  else:
    return ''
  listCells = list(dict.fromkeys(listCells))
  for cell in listCells:
    res = res + cell + '\n'
  return res

if __name__ == '__main__':
  global CountCells
  global filename
  global window
  global cells

  layout =  [
              [
                sg.Multiline(
                  default_text='', 
                  background_color="#ddd", 
                  text_color="black", 
                  size=(50, 15), 
                  key='_INPUT_'
                ),
                sg.Multiline(
                  default_text='', 
                  background_color="#ddd", 
                  text_color="black", 
                  size=(50, 15), 
                  key='_OUTPUT_'
                )
              ],
              [
                sg.Submit('Celdas'), 
                sg.Button('Alarmas'), 
                sg.Button('Masiva'), 
                sg.Button('Email Alarmas'), 
                sg.Button('Email Masiva'), 
                sg.Button('Limpiar'), 
                sg.Button('Cancelar')
              ]
            ]

  iconfile= open("icon.ico","wb")
  iconfile.write(icondata)
  iconfile.close()

  window = sg.Window('Masivo', icon=tempFile).Layout(layout).Finalize()

  while True:
    event, values = window.Read()
    if event is None or event == 'Cancelar':
        break
    if event == 'Celdas':
      if len(window.Element('_INPUT_').Get()) != 1:
        window.Element('_OUTPUT_').Update(
                                    value=ObtenerCeldas(
                                      window.Element('_INPUT_').Get()
                                    )
                                  )
        window.Element('_OUTPUT_').SetFocus()
    if event == 'Alarmas':
      if len(window.Element('_INPUT_').Get()) != 1:
        GetData(window.Element('_INPUT_').Get(), 0)
    if event == 'Masiva':
      if len(window.Element('_INPUT_').Get()) != 1:
        GetData(window.Element('_INPUT_').Get(), 1)
    if event == 'Email Alarmas':
      if len(window.Element('_INPUT_').Get()) != 1:
        GetData(window.Element('_INPUT_').Get(), 2)
    if event == 'Email Masiva':
      if len(window.Element('_INPUT_').Get()) != 1:
        GetData(window.Element('_INPUT_').Get(), 3)
    if event == 'Limpiar':
      window.Element('_OUTPUT_').Update(value='')
      window.Element('_INPUT_').Update(value='')
      window.Element('_INPUT_').SetFocus()
  window.Close()
